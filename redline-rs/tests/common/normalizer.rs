use std::collections::HashSet;

pub struct Normalizer {
    ignore_attributes: HashSet<String>,
    ignore_elements: HashSet<String>,
}

impl Normalizer {
    pub fn new() -> Self {
        let mut ignore_attributes = HashSet::new();
        ignore_attributes.insert("pt:Unid".to_string());
        ignore_attributes.insert("w:rsidR".to_string());
        ignore_attributes.insert("w:rsidRPr".to_string());
        ignore_attributes.insert("w:rsidP".to_string());
        ignore_attributes.insert("w:rsidRDefault".to_string());
        ignore_attributes.insert("w14:textId".to_string());
        ignore_attributes.insert("w14:paraId".to_string());

        Self {
            ignore_attributes,
            ignore_elements: HashSet::new(),
        }
    }

    pub fn normalize(&self, xml: &str) -> String {
        let mut result = xml.to_string();

        for attr in &self.ignore_attributes {
            let pattern = format!(r#" {}="[^"]*""#, attr);
            result = regex_replace(&result, &pattern, "");
        }

        result
    }

    pub fn add_ignore_attribute(&mut self, attr: &str) {
        self.ignore_attributes.insert(attr.to_string());
    }

    pub fn add_ignore_element(&mut self, element: &str) {
        self.ignore_elements.insert(element.to_string());
    }
}

impl Default for Normalizer {
    fn default() -> Self {
        Self::new()
    }
}

fn regex_replace(s: &str, pattern: &str, replacement: &str) -> String {
    let re = regex::Regex::new(pattern).unwrap();
    re.replace_all(s, replacement).to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalizer_removes_unid_attributes() {
        let normalizer = Normalizer::new();
        let xml = r#"<w:p pt:Unid="12345"><w:r><w:t>text</w:t></w:r></w:p>"#;
        let normalized = normalizer.normalize(xml);
        assert!(!normalized.contains("pt:Unid"));
        assert!(normalized.contains("<w:p>"));
    }

    #[test]
    fn normalizer_removes_rsid_attributes() {
        let normalizer = Normalizer::new();
        let xml = r#"<w:p w:rsidR="00123456" w:rsidP="00654321">text</w:p>"#;
        let normalized = normalizer.normalize(xml);
        assert!(!normalized.contains("w:rsidR"));
        assert!(!normalized.contains("w:rsidP"));
    }
}
