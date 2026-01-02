use super::settings::PmlComparerSettings;
use base64::{engine::general_purpose, Engine as _};
use sha2::{Digest, Sha256};
use std::collections::{HashMap, HashSet};
use std::hash::Hash;

// ==================================================================================
// Utility Classes
// ==================================================================================

pub struct PmlHasher;

impl PmlHasher {
    pub fn compute_hash(content: &str) -> String {
        if content.is_empty() {
            return String::new();
        }
        let mut hasher = Sha256::new();
        hasher.update(content.as_bytes());
        general_purpose::STANDARD.encode(hasher.finalize())
    }

    // Note: compute_hash_stream skipped for now as we deal with in-memory structs mostly
}

// ==================================================================================
// Signature Classes
// ==================================================================================

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum PmlShapeType {
    Unknown,
    TextBox,
    AutoShape,
    Picture,
    Table,
    Chart,
    SmartArt,
    Group,
    Connector,
    OleObject,
    Media,
}

impl std::fmt::Display for PmlShapeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{:?}", self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PlaceholderInfo {
    pub type_: String,
    pub index: Option<u32>,
}

#[derive(Debug, Clone)]
pub struct TransformSignature {
    pub x: i64,
    pub y: i64,
    pub cx: i64,
    pub cy: i64,
    pub rotation: i32,
    pub flip_h: bool,
    pub flip_v: bool,
}

impl TransformSignature {
    pub fn is_near(&self, other: &TransformSignature, tolerance: i64) -> bool {
        (self.x - other.x).abs() <= tolerance && (self.y - other.y).abs() <= tolerance
    }

    pub fn is_same_size(&self, other: &TransformSignature, tolerance: i64) -> bool {
        (self.cx - other.cx).abs() <= tolerance && (self.cy - other.cy).abs() <= tolerance
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RunPropertiesSignature {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_name: Option<String>,
    pub font_size: Option<i32>,
    pub font_color: Option<String>,
}

#[derive(Debug, Clone)]
pub struct RunSignature {
    pub text: String,
    pub properties: Option<RunPropertiesSignature>,
    pub content_hash: String,
}

#[derive(Debug, Clone)]
pub struct ParagraphSignature {
    pub runs: Vec<RunSignature>,
    pub plain_text: String,
    pub alignment: Option<String>,
    pub has_bullet: bool,
}

#[derive(Debug, Clone)]
pub struct TextBodySignature {
    pub paragraphs: Vec<ParagraphSignature>,
    pub plain_text: String,
}

#[derive(Debug, Clone)]
pub struct ShapeSignature {
    pub name: String,
    pub id: u32,
    pub type_: PmlShapeType,
    pub placeholder: Option<PlaceholderInfo>,
    pub transform: Option<TransformSignature>,
    pub z_order: i32,
    pub geometry_hash: Option<String>,
    pub text_body: Option<TextBodySignature>,
    pub image_hash: Option<String>,
    pub table_hash: Option<String>,
    pub chart_hash: Option<String>,
    pub children: Option<Vec<ShapeSignature>>,
    pub content_hash: String,
}

#[derive(Debug, Clone)]
pub struct SlideSignature {
    pub index: usize,
    pub relationship_id: String,
    pub layout_relationship_id: Option<String>,
    pub layout_hash: Option<String>,
    pub shapes: Vec<ShapeSignature>,
    pub notes_text: Option<String>,
    pub title_text: Option<String>,
    pub content_hash: String,
    pub background_hash: Option<String>,
}

impl SlideSignature {
    pub fn compute_fingerprint(&self) -> String {
        let mut sb = String::new();
        if let Some(title) = &self.title_text {
            sb.push_str(title);
        }
        sb.push('|');

        let mut shapes: Vec<&ShapeSignature> = self.shapes.iter().collect();
        shapes.sort_by_key(|s| s.z_order);

        for shape in shapes {
            sb.push_str(&shape.name);
            sb.push(':');
            sb.push_str(&shape.type_.to_string());
            sb.push(':');
            if let Some(tb) = &shape.text_body {
                sb.push_str(&tb.plain_text);
            }
            sb.push('|');
        }

        PmlHasher::compute_hash(&sb)
    }
}

#[derive(Debug, Clone)]
pub struct PresentationSignature {
    pub slide_cx: i64,
    pub slide_cy: i64,
    pub slides: Vec<SlideSignature>,
    pub theme_hash: Option<String>,
}

// ==================================================================================
// Matching Classes
// ==================================================================================

#[derive(Debug, Clone, PartialEq, Copy)]
pub enum SlideMatchType {
    Matched,
    Inserted,
    Deleted,
}

#[derive(Debug, Clone)]
pub struct SlideMatch {
    pub match_type: SlideMatchType,
    pub old_index: Option<usize>,
    pub new_index: Option<usize>,
    pub old_slide: Option<SlideSignature>,
    pub new_slide: Option<SlideSignature>,
    pub similarity: f64,
}

impl SlideMatch {
    pub fn was_moved(&self) -> bool {
        self.old_index.is_some() && self.new_index.is_some() && self.old_index != self.new_index
    }
}

pub struct PmlSlideMatchEngine;

impl PmlSlideMatchEngine {
    pub fn match_slides(
        sig1: &PresentationSignature,
        sig2: &PresentationSignature,
        settings: &PmlComparerSettings,
    ) -> Vec<SlideMatch> {
        let mut matches = Vec::new();
        let mut used1 = HashSet::new();
        let mut used2 = HashSet::new();

        // Pass 1: Match by title text (exact match)
        Self::match_by_title_text(sig1, sig2, &mut matches, &mut used1, &mut used2);

        // Pass 2: Match by content fingerprint
        Self::match_by_fingerprint(sig1, sig2, &mut matches, &mut used1, &mut used2, settings);

        // Pass 3: Match by position (remaining slides)
        if settings.use_slide_alignment_lcs {
            Self::match_by_lcs(sig1, sig2, &mut matches, &mut used1, &mut used2, settings);
        } else {
            Self::match_by_position(sig1, sig2, &mut matches, &mut used1, &mut used2);
        }

        // Remaining unmatched = inserted/deleted
        Self::add_unmatched_as_inserted_deleted(sig1, sig2, &mut matches, &used1, &used2);

        // Sort by new index for consistent ordering
        matches.sort_by(|a, b| {
            let idx_a = a.new_index.unwrap_or(usize::MAX);
            let idx_b = b.new_index.unwrap_or(usize::MAX);
            if idx_a != idx_b {
                idx_a.cmp(&idx_b)
            } else {
                let old_idx_a = a.old_index.unwrap_or(usize::MAX);
                let old_idx_b = b.old_index.unwrap_or(usize::MAX);
                old_idx_a.cmp(&old_idx_b)
            }
        });

        matches
    }

    fn match_by_title_text(
        sig1: &PresentationSignature,
        sig2: &PresentationSignature,
        matches: &mut Vec<SlideMatch>,
        used1: &mut HashSet<usize>,
        used2: &mut HashSet<usize>,
    ) {
        for slide1 in &sig1.slides {
            if used1.contains(&slide1.index) {
                continue;
            }

            if slide1.title_text.as_ref().map_or(true, |t| t.is_empty()) {
                continue;
            }

            let match_slide = sig2
                .slides
                .iter()
                .find(|s2| !used2.contains(&s2.index) && s2.title_text == slide1.title_text);

            if let Some(match_slide) = match_slide {
                matches.push(SlideMatch {
                    match_type: SlideMatchType::Matched,
                    old_index: Some(slide1.index),
                    new_index: Some(match_slide.index),
                    old_slide: Some(slide1.clone()),
                    new_slide: Some(match_slide.clone()),
                    similarity: 1.0,
                });
                used1.insert(slide1.index);
                used2.insert(match_slide.index);
            }
        }
    }

    fn match_by_fingerprint(
        sig1: &PresentationSignature,
        sig2: &PresentationSignature,
        matches: &mut Vec<SlideMatch>,
        used1: &mut HashSet<usize>,
        used2: &mut HashSet<usize>,
        _settings: &PmlComparerSettings,
    ) {
        let fingerprints1: HashMap<usize, String> = sig1
            .slides
            .iter()
            .filter(|s| !used1.contains(&s.index))
            .map(|s| (s.index, s.compute_fingerprint()))
            .collect();

        let fingerprints2: HashMap<usize, String> = sig2
            .slides
            .iter()
            .filter(|s| !used2.contains(&s.index))
            .map(|s| (s.index, s.compute_fingerprint()))
            .collect();

        for slide1 in &sig1.slides {
            if used1.contains(&slide1.index) {
                continue;
            }
            let fp1 = &fingerprints1[&slide1.index];

            let match_slide = sig2
                .slides
                .iter()
                .find(|s2| !used2.contains(&s2.index) && fingerprints2.get(&s2.index) == Some(fp1));

            if let Some(match_slide) = match_slide {
                matches.push(SlideMatch {
                    match_type: SlideMatchType::Matched,
                    old_index: Some(slide1.index),
                    new_index: Some(match_slide.index),
                    old_slide: Some(slide1.clone()),
                    new_slide: Some(match_slide.clone()),
                    similarity: 1.0,
                });
                used1.insert(slide1.index);
                used2.insert(match_slide.index);
            }
        }
    }

    fn match_by_lcs(
        sig1: &PresentationSignature,
        sig2: &PresentationSignature,
        matches: &mut Vec<SlideMatch>,
        used1: &mut HashSet<usize>,
        used2: &mut HashSet<usize>,
        settings: &PmlComparerSettings,
    ) {
        let remaining1: Vec<&SlideSignature> = sig1
            .slides
            .iter()
            .filter(|s| !used1.contains(&s.index))
            .collect();
        let remaining2: Vec<&SlideSignature> = sig2
            .slides
            .iter()
            .filter(|s| !used2.contains(&s.index))
            .collect();

        if remaining1.is_empty() || remaining2.is_empty() {
            return;
        }

        let mut similarities = vec![vec![0.0; remaining2.len()]; remaining1.len()];
        for (i, r1) in remaining1.iter().enumerate() {
            for (j, r2) in remaining2.iter().enumerate() {
                similarities[i][j] = Self::compute_slide_similarity(r1, r2);
            }
        }

        let mut matched1 = HashSet::new();
        let mut matched2 = HashSet::new();

        while matched1.len() < remaining1.len() && matched2.len() < remaining2.len() {
            let mut best_sim = 0.0;
            let mut best_i = -1;
            let mut best_j = -1;

            for i in 0..remaining1.len() {
                if matched1.contains(&i) {
                    continue;
                }
                for j in 0..remaining2.len() {
                    if matched2.contains(&j) {
                        continue;
                    }
                    if similarities[i][j] > best_sim {
                        best_sim = similarities[i][j];
                        best_i = i as i32;
                        best_j = j as i32;
                    }
                }
            }

            // Note: C# uses settings.SlideSimilarityThreshold (0.4 by default)
            let threshold = settings.slide_similarity_threshold;

            if best_i < 0 || best_sim < threshold {
                break;
            }

            let idx_i = best_i as usize;
            let idx_j = best_j as usize;

            matches.push(SlideMatch {
                match_type: SlideMatchType::Matched,
                old_index: Some(remaining1[idx_i].index),
                new_index: Some(remaining2[idx_j].index),
                old_slide: Some(remaining1[idx_i].clone()),
                new_slide: Some(remaining2[idx_j].clone()),
                similarity: best_sim,
            });

            used1.insert(remaining1[idx_i].index);
            used2.insert(remaining2[idx_j].index);
            matched1.insert(idx_i);
            matched2.insert(idx_j);
        }
    }

    fn match_by_position(
        sig1: &PresentationSignature,
        sig2: &PresentationSignature,
        matches: &mut Vec<SlideMatch>,
        used1: &mut HashSet<usize>,
        used2: &mut HashSet<usize>,
    ) {
        let remaining1: Vec<&SlideSignature> = sig1
            .slides
            .iter()
            .filter(|s| !used1.contains(&s.index))
            .collect(); // Already sorted by index naturally if vector is ordered

        let remaining2: Vec<&SlideSignature> = sig2
            .slides
            .iter()
            .filter(|s| !used2.contains(&s.index))
            .collect();

        let count = std::cmp::min(remaining1.len(), remaining2.len());
        for i in 0..count {
            matches.push(SlideMatch {
                match_type: SlideMatchType::Matched,
                old_index: Some(remaining1[i].index),
                new_index: Some(remaining2[i].index),
                old_slide: Some(remaining1[i].clone()),
                new_slide: Some(remaining2[i].clone()),
                similarity: Self::compute_slide_similarity(remaining1[i], remaining2[i]),
            });
            used1.insert(remaining1[i].index);
            used2.insert(remaining2[i].index);
        }
    }

    fn add_unmatched_as_inserted_deleted(
        sig1: &PresentationSignature,
        sig2: &PresentationSignature,
        matches: &mut Vec<SlideMatch>,
        used1: &HashSet<usize>,
        used2: &HashSet<usize>,
    ) {
        // Deleted slides
        for slide in sig1.slides.iter().filter(|s| !used1.contains(&s.index)) {
            matches.push(SlideMatch {
                match_type: SlideMatchType::Deleted,
                old_index: Some(slide.index),
                new_index: None,
                old_slide: Some(slide.clone()),
                new_slide: None,
                similarity: 0.0,
            });
        }

        // Inserted slides
        for slide in sig2.slides.iter().filter(|s| !used2.contains(&s.index)) {
            matches.push(SlideMatch {
                match_type: SlideMatchType::Inserted,
                old_index: None,
                new_index: Some(slide.index),
                old_slide: None,
                new_slide: Some(slide.clone()),
                similarity: 0.0,
            });
        }
    }

    fn compute_slide_similarity(s1: &SlideSignature, s2: &SlideSignature) -> f64 {
        let mut score = 0.0;
        let mut max_score = 0.0;

        // Title match
        let t1 = s1.title_text.as_deref().unwrap_or("");
        let t2 = s2.title_text.as_deref().unwrap_or("");

        if !t1.is_empty() || !t2.is_empty() {
            max_score += 3.0;
            if !t1.is_empty() && t1 == t2 {
                score += 3.0;
            } else if !t1.is_empty() && !t2.is_empty() {
                score += Self::compute_text_similarity(t1, t2) * 2.0;
            }
        }

        // Content hash match
        max_score += 1.0;
        if s1.content_hash == s2.content_hash {
            score += 1.0;
        }

        // Shape count similarity
        max_score += 1.0;
        let c1 = s1.shapes.len() as i32;
        let c2 = s2.shapes.len() as i32;
        if c1 == c2 {
            score += 1.0;
        } else if (c1 - c2).abs() <= 2 {
            score += 0.5;
        }

        // Shape types overlap
        max_score += 1.0;
        let types1: Vec<_> = s1.shapes.iter().map(|s| &s.type_).collect();
        let types2: Vec<_> = s2.shapes.iter().map(|s| &s.type_).collect();
        let _common_types = types1.iter().filter(|t| types2.contains(t)).count(); // This is rough intersection, C# Intersect does set intersection

        // Exact C# Intersect behavior: Distinct elements that appear in both
        let distinct_types1: HashSet<_> = types1.into_iter().collect();
        let distinct_types2: HashSet<_> = types2.into_iter().collect();
        let common_count = distinct_types1.intersection(&distinct_types2).count();
        let total_count = std::cmp::max(distinct_types1.len(), distinct_types2.len());

        if total_count > 0 {
            score += common_count as f64 / total_count as f64;
        }

        // Shape names overlap
        max_score += 2.0;
        let names1: HashSet<_> = s1
            .shapes
            .iter()
            .map(|s| &s.name)
            .filter(|n| !n.is_empty())
            .collect();
        let names2: HashSet<_> = s2
            .shapes
            .iter()
            .map(|s| &s.name)
            .filter(|n| !n.is_empty())
            .collect();
        let common_names = names1.intersection(&names2).count();
        let total_names = std::cmp::max(names1.len(), names2.len());

        if total_names > 0 {
            score += 2.0 * common_names as f64 / total_names as f64;
        }

        if max_score > 0.0 {
            score / max_score
        } else {
            0.0
        }
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

        // Simple Jaccard similarity on words
        let words1: HashSet<String> = s1
            .to_lowercase()
            .split_whitespace()
            .map(|s| s.to_string())
            .collect();
        let words2: HashSet<String> = s2
            .to_lowercase()
            .split_whitespace()
            .map(|s| s.to_string())
            .collect();

        let intersection = words1.intersection(&words2).count();
        let union = words1.union(&words2).count();

        if union > 0 {
            intersection as f64 / union as f64
        } else {
            0.0
        }
    }
}
