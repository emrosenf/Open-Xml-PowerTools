
#[cfg(test)]
mod tests {
    use crate::wml::document::WmlDocument;
    use crate::wml::comparer::WmlComparer;
    use crate::wml::settings::WmlComparerSettings;
    use crate::wml::atom_list::create_comparison_unit_atom_list;
    use crate::wml::document::find_document_body;
    use crate::wml::comparison_unit::{get_comparison_unit_list, WordSeparatorSettings, ComparisonCorrelationStatus};
    use crate::wml::lcs_algorithm::{lcs, flatten_to_atoms};
    use crate::wml::preprocess::{preprocess_markup, PreProcessSettings};

    /// Deep trace test for WC-1010 (Digits test)
    /// 
    /// Expected: 4 revisions (digit modifications)
    /// Doc1: 12.34, 12,34, Ab,cd, Test., .Test.123
    /// Doc2: 12.34, 12,4, Ab,cd, st., .Test.123
    /// 
    /// Changes:
    /// - Para 2: "12,34" -> "12,4" (delete "3")
    /// - Para 3: "Ab,cd" (same text but split across runs) 
    /// - Para 4: "Test." -> "st." (delete "Te")
    #[test]
    fn trace_wc1010_data_flow() {
        // Recreate the documents from their actual content
        let xml1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>12.34</w:t></w:r></w:p>
    <w:p><w:r><w:t>12,34</w:t></w:r></w:p>
    <w:p><w:r><w:t>Ab,cd</w:t></w:r></w:p>
    <w:p><w:r><w:t>Test.</w:t></w:r></w:p>
    <w:p><w:r><w:t>.Test.123</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

        let xml2 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>12.34</w:t></w:r></w:p>
    <w:p><w:r><w:t>12,4</w:t></w:r></w:p>
    <w:p><w:r><w:t>Ab,</w:t></w:r><w:r><w:t>cd</w:t></w:r></w:p>
    <w:p><w:r><w:t>st.</w:t></w:r></w:p>
    <w:p><w:r><w:t>.Test.123</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

        let doc1 = WmlDocument::from_main_xml(xml1.as_bytes()).unwrap();
        let doc2 = WmlDocument::from_main_xml(xml2.as_bytes()).unwrap();

        let mut main1 = doc1.main_document().unwrap();
        let mut main2 = doc2.main_document().unwrap();

        let body1 = find_document_body(&main1).unwrap();
        let body2 = find_document_body(&main2).unwrap();

        // Preprocess
        let preprocess_settings = PreProcessSettings::for_comparison();
        preprocess_markup(&mut main1, body1, &preprocess_settings).unwrap();
        preprocess_markup(&mut main2, body2, &preprocess_settings).unwrap();

        let settings = WmlComparerSettings::default();

        // Extract atoms
        let atoms1 = create_comparison_unit_atom_list(&mut main1, body1, "main", &settings);
        let atoms2 = create_comparison_unit_atom_list(&mut main2, body2, "main", &settings);

        println!("\n=== ATOM EXTRACTION ===");
        println!("Doc1 atoms: {}", atoms1.len());
        println!("Doc2 atoms: {}", atoms2.len());

        // Print atoms with their content
        println!("\n--- Doc1 Atoms ---");
        let mut doc1_text = String::new();
        for (i, atom) in atoms1.iter().enumerate() {
            let display = atom.content_element.display_value();
            doc1_text.push_str(&display);
            if i < 50 {
                println!("  [{}] {:?} hash={}", i, atom.content_element, &atom.sha1_hash[..8]);
            }
        }
        println!("Doc1 text: {:?}", doc1_text);

        println!("\n--- Doc2 Atoms ---");
        let mut doc2_text = String::new();
        for (i, atom) in atoms2.iter().enumerate() {
            let display = atom.content_element.display_value();
            doc2_text.push_str(&display);
            if i < 50 {
                println!("  [{}] {:?} hash={}", i, atom.content_element, &atom.sha1_hash[..8]);
            }
        }
        println!("Doc2 text: {:?}", doc2_text);

        // Build comparison units
        let word_settings = WordSeparatorSettings::default();
        let units1 = get_comparison_unit_list(atoms1.clone(), &word_settings);
        let units2 = get_comparison_unit_list(atoms2.clone(), &word_settings);

        println!("\n=== COMPARISON UNITS ===");
        println!("Doc1 units: {}", units1.len());
        println!("Doc2 units: {}", units2.len());

        for (i, unit) in units1.iter().enumerate() {
            if i < 20 {
                println!("  [{}] {:?}", i, unit);
            }
        }

        // Run LCS
        let correlated = lcs(units1, units2, &settings);

        println!("\n=== LCS RESULT ===");
        println!("Correlated sequences: {}", correlated.len());
        for (i, seq) in correlated.iter().enumerate() {
            println!("  [{}] status={:?} len1={} len2={}", 
                i, seq.status, seq.len1(), seq.len2());
        }

        // Flatten
        let flattened = flatten_to_atoms(&correlated);

        println!("\n=== FLATTENED ATOMS ===");
        println!("Total flattened atoms: {}", flattened.len());

        let equal_count = flattened.iter().filter(|a| a.correlation_status == ComparisonCorrelationStatus::Equal).count();
        let deleted_count = flattened.iter().filter(|a| a.correlation_status == ComparisonCorrelationStatus::Deleted).count();
        let inserted_count = flattened.iter().filter(|a| a.correlation_status == ComparisonCorrelationStatus::Inserted).count();

        println!("Equal: {}, Deleted: {}, Inserted: {}", equal_count, deleted_count, inserted_count);

        // Print non-equal atoms
        println!("\n--- Non-Equal Atoms ---");
        for (i, atom) in flattened.iter().enumerate() {
            if atom.correlation_status != ComparisonCorrelationStatus::Equal {
                println!("  [{}] {} {:?}", i, atom.correlation_status, atom.content_element);
            }
        }

        let result = WmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
        
        println!("\n=== FINAL RESULT ===");
        println!("Insertions: {}, Deletions: {}, Format: {}, Total: {}", 
            result.insertions, result.deletions, result.format_changes, result.revision_count);

        match WmlDocument::from_bytes(&result.document) {
            Ok(res_doc) => {
                match res_doc.main_document() {
                    Ok(xml) => {
                        match crate::xml::builder::serialize(&xml) {
                            Ok(xml_str) => {
                                println!("\n=== RESULT XML ===");
                                println!("{}", xml_str);
                            }
                            Err(e) => println!("Serialize error: {:?}", e),
                        }
                    }
                    Err(e) => println!("Main document error: {:?}", e),
                }
            }
            Err(e) => println!("From bytes error: {:?}", e),
        }
        
        assert!(result.revision_count > 0, "Should detect changes");
    }

    #[test]
    fn test_mismatched_paragraphs_with_same_start() {
        let xml1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>The quick brown fox jumps over the lazy dog.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let xml2 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>The quick brown fox jumps over the active cat.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let doc1 = WmlDocument::from_main_xml(xml1.as_bytes()).unwrap();
        let doc2 = WmlDocument::from_main_xml(xml2.as_bytes()).unwrap();
        
        let settings = WmlComparerSettings::default();
        let result = WmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
        
        println!("Insertions: {}, Deletions: {}, FormatChanges: {}", result.insertions, result.deletions, result.format_changes);
        
        if result.insertions == 0 || result.deletions == 0 {
            if let Ok(res_doc) = WmlDocument::from_bytes(&result.document) {
                if let Ok(xml) = res_doc.main_document() {
                    if let Ok(xml_str) = crate::xml::builder::serialize(&xml) {
                        println!("Main Document XML:\n{}", xml_str);
                    }
                }
            }
        }

        assert!(result.insertions > 0, "Should have insertions");
        assert!(result.deletions > 0, "Should have deletions");
    }

    /// Trace test specifically for footnoteReference handling
    /// Tests that footnoteReference is preserved when surrounding text changes
    #[test]
    fn trace_footnote_reference_flow() {
        use crate::wml::comparison_unit::ContentElement;
        
        // Simplified test case with footnoteReference
        let xml1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t>The original text that will change.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let xml2 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t>The modified text that is different.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let doc1 = WmlDocument::from_main_xml(xml1.as_bytes()).unwrap();
        let doc2 = WmlDocument::from_main_xml(xml2.as_bytes()).unwrap();

        let mut main1 = doc1.main_document().unwrap();
        let mut main2 = doc2.main_document().unwrap();

        let body1 = find_document_body(&main1).unwrap();
        let body2 = find_document_body(&main2).unwrap();

        // Preprocess
        let preprocess_settings = PreProcessSettings::for_comparison();
        preprocess_markup(&mut main1, body1, &preprocess_settings).unwrap();
        preprocess_markup(&mut main2, body2, &preprocess_settings).unwrap();

        let settings = WmlComparerSettings::default();

        println!("\n=== STEP 1: CREATE ATOMS ===");
        let atoms1 = create_comparison_unit_atom_list(&mut main1, body1, "main", &settings);
        let atoms2 = create_comparison_unit_atom_list(&mut main2, body2, "main", &settings);

        println!("Doc1 atoms: {}", atoms1.len());
        for (i, atom) in atoms1.iter().enumerate() {
            let is_footnote = matches!(atom.content_element, ContentElement::FootnoteReference { .. });
            println!("  [{}] {} {:?} hash={}", i, if is_footnote { "**FOOTNOTE**" } else { "" }, 
                atom.content_element, &atom.sha1_hash[..8]);
        }

        println!("\nDoc2 atoms: {}", atoms2.len());
        for (i, atom) in atoms2.iter().enumerate() {
            let is_footnote = matches!(atom.content_element, ContentElement::FootnoteReference { .. });
            println!("  [{}] {} {:?} hash={}", i, if is_footnote { "**FOOTNOTE**" } else { "" }, 
                atom.content_element, &atom.sha1_hash[..8]);
        }

        println!("\n=== STEP 2: GROUP INTO WORDS ===");
        let word_settings = WordSeparatorSettings::default();
        let units1 = get_comparison_unit_list(atoms1.clone(), &word_settings);
        let units2 = get_comparison_unit_list(atoms2.clone(), &word_settings);

        println!("Doc1 units (words/groups): {}", units1.len());
        for (i, unit) in units1.iter().enumerate() {
            if let Some(word) = unit.as_word() {
                let has_footnote = word.atoms.iter().any(|a| matches!(a.content_element, ContentElement::FootnoteReference { .. }));
                println!("  [{}] Word ({}atoms) {} hash={:.8}", i, word.atoms.len(), 
                    if has_footnote { "**HAS FOOTNOTE**" } else { "" }, word.sha1_hash);
                for atom in word.atoms.iter() {
                    println!("       - {:?}", atom.content_element);
                }
            } else if let Some(group) = unit.as_group() {
                println!("  [{}] Group {:?} hash={:.8}", i, group.group_type, group.sha1_hash);
            }
        }

        println!("\nDoc2 units (words/groups): {}", units2.len());
        for (i, unit) in units2.iter().enumerate() {
            if let Some(word) = unit.as_word() {
                let has_footnote = word.atoms.iter().any(|a| matches!(a.content_element, ContentElement::FootnoteReference { .. }));
                println!("  [{}] Word ({}atoms) {} hash={:.8}", i, word.atoms.len(), 
                    if has_footnote { "**HAS FOOTNOTE**" } else { "" }, word.sha1_hash);
            } else if let Some(group) = unit.as_group() {
                println!("  [{}] Group {:?} hash={:.8}", i, group.group_type, group.sha1_hash);
            }
        }

        println!("\n=== STEP 3: RUN LCS ===");
        let correlated = lcs(units1, units2, &settings);

        println!("Correlated sequences: {}", correlated.len());
        for (i, seq) in correlated.iter().enumerate() {
            println!("  [{}] status={:?} len1={} len2={}", 
                i, seq.status, seq.len1(), seq.len2());
        }

        println!("\n=== STEP 4: FLATTEN TO ATOMS ===");
        let flattened = flatten_to_atoms(&correlated);

        println!("Total flattened atoms: {}", flattened.len());
        
        // Find and report on footnote atoms specifically
        println!("\n--- Footnote Reference Status ---");
        for (i, atom) in flattened.iter().enumerate() {
            if matches!(atom.content_element, ContentElement::FootnoteReference { .. }) {
                println!("  [{}] FOOTNOTE status={:?} element={:?}", 
                    i, atom.correlation_status, atom.content_element);
            }
        }

        println!("\n--- All Non-Equal Atoms ---");
        for (i, atom) in flattened.iter().enumerate() {
            if atom.correlation_status != ComparisonCorrelationStatus::Equal {
                println!("  [{}] {} {:?}", i, atom.correlation_status, atom.content_element);
            }
        }

        // The footnote should be Equal, not Deleted or Inserted
        let footnote_atoms: Vec<_> = flattened.iter()
            .filter(|a| matches!(a.content_element, ContentElement::FootnoteReference { .. }))
            .collect();
        
        println!("\n=== ASSERTION ===");
        println!("Found {} footnote atom(s)", footnote_atoms.len());
        for atom in &footnote_atoms {
            println!("  Status: {:?}", atom.correlation_status);
        }

        assert!(!footnote_atoms.is_empty(), "Should have footnote atoms");
        assert!(
            footnote_atoms.iter().all(|a| a.correlation_status == ComparisonCorrelationStatus::Equal),
            "Footnote should be Equal, not Deleted/Inserted"
        );
    }
}
