use std::collections::HashSet;
use super::settings::PmlComparerSettings;
use super::slide_matching::{
    ShapeSignature, SlideSignature, PmlShapeType,
};

// ==================================================================================
// Shape Matching Classes
// ==================================================================================

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeMatchType {
    Matched,
    Inserted,
    Deleted,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeMatchMethod {
    Placeholder,
    NameAndType,
    NameOnly,
    Fuzzy,
}

#[derive(Debug, Clone)]
pub struct ShapeMatch {
    pub match_type: ShapeMatchType,
    pub old_shape: Option<ShapeSignature>,
    pub new_shape: Option<ShapeSignature>,
    pub score: f64,
    pub method: Option<ShapeMatchMethod>,
}

// ==================================================================================
// Shape Match Engine
// ==================================================================================

/// Matches shapes within slides.
/// Ported from C# PmlComparer.cs lines 1446-1777
pub struct PmlShapeMatchEngine;

impl PmlShapeMatchEngine {
    pub fn match_shapes(
        slide1: &SlideSignature,
        slide2: &SlideSignature,
        settings: &PmlComparerSettings,
    ) -> Vec<ShapeMatch> {
        let mut matches = Vec::new();
        let mut used1 = HashSet::new();
        let mut used2 = HashSet::new();

        // Pass 1: Match by placeholder
        Self::match_by_placeholder(slide1, slide2, &mut matches, &mut used1, &mut used2);

        // Pass 2: Match by name and type
        Self::match_by_name_and_type(slide1, slide2, &mut matches, &mut used1, &mut used2);

        // Pass 3: Match by name only
        Self::match_by_name_only(slide1, slide2, &mut matches, &mut used1, &mut used2);

        // Pass 4: Fuzzy matching
        if settings.enable_fuzzy_shape_matching {
            Self::fuzzy_match(slide1, slide2, &mut matches, &mut used1, &mut used2, settings);
        }

        // Remaining unmatched
        Self::add_unmatched_as_inserted_deleted(slide1, slide2, &mut matches, &used1, &used2);

        matches
    }

    fn get_shape_key(shape: &ShapeSignature) -> String {
        format!("{}:{}", shape.id, shape.name)
    }

    fn match_by_placeholder(
        slide1: &SlideSignature,
        slide2: &SlideSignature,
        matches: &mut Vec<ShapeMatch>,
        used1: &mut HashSet<String>,
        used2: &mut HashSet<String>,
    ) {
        let placeholders1: Vec<&ShapeSignature> = slide1
            .shapes
            .iter()
            .filter(|s| s.placeholder.is_some())
            .collect();

        for shape1 in placeholders1 {
            let key1 = Self::get_shape_key(shape1);
            if used1.contains(&key1) {
                continue;
            }

            let match_shape = slide2.shapes.iter().find(|s2| {
                s2.placeholder.is_some()
                    && !used2.contains(&Self::get_shape_key(s2))
                    && s2.placeholder == shape1.placeholder
            });

            if let Some(match_shape) = match_shape {
                let key2 = Self::get_shape_key(match_shape);
                matches.push(ShapeMatch {
                    match_type: ShapeMatchType::Matched,
                    old_shape: Some(shape1.clone()),
                    new_shape: Some(match_shape.clone()),
                    score: 1.0,
                    method: Some(ShapeMatchMethod::Placeholder),
                });
                used1.insert(key1);
                used2.insert(key2);
            }
        }
    }

    fn match_by_name_and_type(
        slide1: &SlideSignature,
        slide2: &SlideSignature,
        matches: &mut Vec<ShapeMatch>,
        used1: &mut HashSet<String>,
        used2: &mut HashSet<String>,
    ) {
        for shape1 in &slide1.shapes {
            let key1 = Self::get_shape_key(shape1);
            if used1.contains(&key1) {
                continue;
            }

            if shape1.name.is_empty() {
                continue;
            }

            let match_shape = slide2.shapes.iter().find(|s2| {
                !used2.contains(&Self::get_shape_key(s2))
                    && s2.name == shape1.name
                    && s2.type_ == shape1.type_
            });

            if let Some(match_shape) = match_shape {
                let key2 = Self::get_shape_key(match_shape);
                matches.push(ShapeMatch {
                    match_type: ShapeMatchType::Matched,
                    old_shape: Some(shape1.clone()),
                    new_shape: Some(match_shape.clone()),
                    score: 0.95,
                    method: Some(ShapeMatchMethod::NameAndType),
                });
                used1.insert(key1);
                used2.insert(key2);
            }
        }
    }

    fn match_by_name_only(
        slide1: &SlideSignature,
        slide2: &SlideSignature,
        matches: &mut Vec<ShapeMatch>,
        used1: &mut HashSet<String>,
        used2: &mut HashSet<String>,
    ) {
        for shape1 in &slide1.shapes {
            let key1 = Self::get_shape_key(shape1);
            if used1.contains(&key1) {
                continue;
            }

            if shape1.name.is_empty() {
                continue;
            }

            let match_shape = slide2
                .shapes
                .iter()
                .find(|s2| !used2.contains(&Self::get_shape_key(s2)) && s2.name == shape1.name);

            if let Some(match_shape) = match_shape {
                let key2 = Self::get_shape_key(match_shape);
                matches.push(ShapeMatch {
                    match_type: ShapeMatchType::Matched,
                    old_shape: Some(shape1.clone()),
                    new_shape: Some(match_shape.clone()),
                    score: 0.8,
                    method: Some(ShapeMatchMethod::NameOnly),
                });
                used1.insert(key1);
                used2.insert(key2);
            }
        }
    }

    fn fuzzy_match(
        slide1: &SlideSignature,
        slide2: &SlideSignature,
        matches: &mut Vec<ShapeMatch>,
        used1: &mut HashSet<String>,
        used2: &mut HashSet<String>,
        settings: &PmlComparerSettings,
    ) {
        let remaining1: Vec<&ShapeSignature> = slide1
            .shapes
            .iter()
            .filter(|s| !used1.contains(&Self::get_shape_key(s)))
            .collect();

        let remaining2: Vec<&ShapeSignature> = slide2
            .shapes
            .iter()
            .filter(|s| !used2.contains(&Self::get_shape_key(s)))
            .collect();

        for shape1 in remaining1 {
            let key1 = Self::get_shape_key(shape1);
            if used1.contains(&key1) {
                continue;
            }

            let mut best_score = 0.0;
            let mut best_match: Option<&ShapeSignature> = None;

            for shape2 in &remaining2 {
                let key2 = Self::get_shape_key(shape2);
                if used2.contains(&key2) {
                    continue;
                }

                let score = Self::compute_shape_match_score(shape1, shape2, settings);
                if score > best_score && score >= settings.shape_similarity_threshold {
                    best_score = score;
                    best_match = Some(shape2);
                }
            }

            if let Some(best_match) = best_match {
                let key2 = Self::get_shape_key(best_match);
                matches.push(ShapeMatch {
                    match_type: ShapeMatchType::Matched,
                    old_shape: Some(shape1.clone()),
                    new_shape: Some(best_match.clone()),
                    score: best_score,
                    method: Some(ShapeMatchMethod::Fuzzy),
                });
                used1.insert(key1);
                used2.insert(key2);
            }
        }
    }

    fn add_unmatched_as_inserted_deleted(
        slide1: &SlideSignature,
        slide2: &SlideSignature,
        matches: &mut Vec<ShapeMatch>,
        used1: &HashSet<String>,
        used2: &HashSet<String>,
    ) {
        // Deleted shapes
        for shape in &slide1.shapes {
            if !used1.contains(&Self::get_shape_key(shape)) {
                matches.push(ShapeMatch {
                    match_type: ShapeMatchType::Deleted,
                    old_shape: Some(shape.clone()),
                    new_shape: None,
                    score: 0.0,
                    method: None,
                });
            }
        }

        // Inserted shapes
        for shape in &slide2.shapes {
            if !used2.contains(&Self::get_shape_key(shape)) {
                matches.push(ShapeMatch {
                    match_type: ShapeMatchType::Inserted,
                    old_shape: None,
                    new_shape: Some(shape.clone()),
                    score: 0.0,
                    method: None,
                });
            }
        }
    }

    fn compute_shape_match_score(
        s1: &ShapeSignature,
        s2: &ShapeSignature,
        settings: &PmlComparerSettings,
    ) -> f64 {
        let mut score = 0.0;

        // Same type (required)
        if s1.type_ != s2.type_ {
            return 0.0;
        }
        score += 0.2;

        // Position similarity
        if let (Some(t1), Some(t2)) = (&s1.transform, &s2.transform) {
            if t1.is_near(t2, settings.position_tolerance) {
                score += 0.3;
            } else {
                // Partial credit for nearby positions
                let distance = ((t1.x - t2.x).pow(2) as f64 + (t1.y - t2.y).pow(2) as f64).sqrt();
                if distance < (settings.position_tolerance * 5) as f64 {
                    score += 0.1;
                }
            }
        }

        // Content similarity
        match s1.type_ {
            PmlShapeType::Picture => {
                if let (Some(img1), Some(img2)) = (&s1.image_hash, &s2.image_hash) {
                    if img1 == img2 {
                        score += 0.5;
                    }
                }
            }
            _ => {
                if let (Some(tb1), Some(tb2)) = (&s1.text_body, &s2.text_body) {
                    if tb1.plain_text == tb2.plain_text {
                        score += 0.5;
                    } else {
                        let text_sim = Self::compute_text_similarity(&tb1.plain_text, &tb2.plain_text);
                        score += text_sim * 0.5;
                    }
                } else if s1.content_hash == s2.content_hash {
                    score += 0.5;
                }
            }
        }

        score
    }

    fn compute_text_similarity(s1: &str, s2: &str) -> f64 {
        if s1.is_empty() && s2.is_empty() {
            return 1.0;
        }
        if s1.is_empty() || s2.is_empty() {
            return 0.0;
        }
        if s1 == s2 {
            return 1.0;
        }

        // Levenshtein-based similarity
        let max_len = s1.len().max(s2.len());
        if max_len == 0 {
            return 1.0;
        }

        let distance = Self::levenshtein_distance(s1, s2);
        1.0 - (distance as f64 / max_len as f64)
    }

    fn levenshtein_distance(s1: &str, s2: &str) -> usize {
        let m = s1.len();
        let n = s2.len();
        let mut d = vec![vec![0; n + 1]; m + 1];

        for i in 0..=m {
            d[i][0] = i;
        }
        for j in 0..=n {
            d[0][j] = j;
        }

        let s1_chars: Vec<char> = s1.chars().collect();
        let s2_chars: Vec<char> = s2.chars().collect();

        for j in 1..=n {
            for i in 1..=m {
                let cost = if s1_chars[i - 1] == s2_chars[j - 1] {
                    0
                } else {
                    1
                };
                d[i][j] = (d[i - 1][j] + 1)
                    .min(d[i][j - 1] + 1)
                    .min(d[i - 1][j - 1] + cost);
            }
        }

        d[m][n]
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::super::slide_matching::PlaceholderInfo;

    fn create_test_shape(id: u32, name: &str, type_: PmlShapeType) -> ShapeSignature {
        ShapeSignature {
            name: name.to_string(),
            id,
            type_,
            placeholder: None,
            transform: None,
            z_order: 0,
            geometry_hash: None,
            text_body: None,
            image_hash: None,
            table_hash: None,
            chart_hash: None,
            children: None,
            content_hash: String::new(),
        }
    }

    #[test]
    fn test_match_by_placeholder() {
        let mut slide1 = SlideSignature {
            index: 0,
            relationship_id: "rId1".to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: vec![],
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        let mut shape1 = create_test_shape(1, "Title 1", PmlShapeType::TextBox);
        shape1.placeholder = Some(PlaceholderInfo {
            type_: "title".to_string(),
            index: None,
        });
        slide1.shapes.push(shape1);

        let mut slide2 = SlideSignature {
            index: 1,
            relationship_id: "rId2".to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: vec![],
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        let mut shape2 = create_test_shape(2, "Title 2", PmlShapeType::TextBox);
        shape2.placeholder = Some(PlaceholderInfo {
            type_: "title".to_string(),
            index: None,
        });
        slide2.shapes.push(shape2);

        let settings = PmlComparerSettings::default();
        let matches = PmlShapeMatchEngine::match_shapes(&slide1, &slide2, &settings);

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].match_type, ShapeMatchType::Matched);
        assert_eq!(matches[0].method, Some(ShapeMatchMethod::Placeholder));
        assert_eq!(matches[0].score, 1.0);
    }

    #[test]
    fn test_match_by_name_and_type() {
        let slide1 = SlideSignature {
            index: 0,
            relationship_id: "rId1".to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: vec![create_test_shape(1, "MyShape", PmlShapeType::AutoShape)],
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        let slide2 = SlideSignature {
            index: 1,
            relationship_id: "rId2".to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: vec![create_test_shape(2, "MyShape", PmlShapeType::AutoShape)],
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        let settings = PmlComparerSettings::default();
        let matches = PmlShapeMatchEngine::match_shapes(&slide1, &slide2, &settings);

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].match_type, ShapeMatchType::Matched);
        assert_eq!(matches[0].method, Some(ShapeMatchMethod::NameAndType));
        assert_eq!(matches[0].score, 0.95);
    }

    #[test]
    fn test_levenshtein_distance() {
        assert_eq!(PmlShapeMatchEngine::levenshtein_distance("kitten", "sitting"), 3);
        assert_eq!(PmlShapeMatchEngine::levenshtein_distance("hello", "hello"), 0);
        assert_eq!(PmlShapeMatchEngine::levenshtein_distance("", "test"), 4);
    }

    #[test]
    fn test_compute_text_similarity() {
        assert_eq!(PmlShapeMatchEngine::compute_text_similarity("hello", "hello"), 1.0);
        assert_eq!(PmlShapeMatchEngine::compute_text_similarity("", ""), 1.0);
        assert_eq!(PmlShapeMatchEngine::compute_text_similarity("", "test"), 0.0);
        
        let sim = PmlShapeMatchEngine::compute_text_similarity("hello", "hallo");
        assert!(sim > 0.7 && sim < 1.0);
    }

    #[test]
    fn test_inserted_and_deleted_shapes() {
        let slide1 = SlideSignature {
            index: 0,
            relationship_id: "rId1".to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: vec![create_test_shape(1, "Shape1", PmlShapeType::AutoShape)],
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        let slide2 = SlideSignature {
            index: 1,
            relationship_id: "rId2".to_string(),
            layout_relationship_id: None,
            layout_hash: None,
            shapes: vec![create_test_shape(2, "Shape2", PmlShapeType::Picture)], // Different type to prevent fuzzy match
            notes_text: None,
            title_text: None,
            content_hash: String::new(),
            background_hash: None,
        };

        let settings = PmlComparerSettings::default();
        let matches = PmlShapeMatchEngine::match_shapes(&slide1, &slide2, &settings);

        // Should have 2 matches: 1 deleted, 1 inserted (fuzzy matching won't match due to different types)
        assert_eq!(matches.len(), 2);
        
        let deleted = matches.iter().find(|m| m.match_type == ShapeMatchType::Deleted);
        let inserted = matches.iter().find(|m| m.match_type == ShapeMatchType::Inserted);
        
        assert!(deleted.is_some());
        assert!(inserted.is_some());
    }
}
