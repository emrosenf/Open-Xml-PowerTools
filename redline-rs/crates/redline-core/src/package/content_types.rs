use std::collections::HashMap;

#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    defaults: HashMap<String, String>,
    overrides: HashMap<String, String>,
}

impl ContentTypes {
    pub fn new() -> Self {
        let mut defaults = HashMap::new();
        defaults.insert("rels".to_string(), "application/vnd.openxmlformats-package.relationships+xml".to_string());
        defaults.insert("xml".to_string(), "application/xml".to_string());
        
        Self {
            defaults,
            overrides: HashMap::new(),
        }
    }

    pub fn get_content_type(&self, path: &str) -> Option<&str> {
        if let Some(ct) = self.overrides.get(path) {
            return Some(ct);
        }
        
        if let Some(ext) = path.rsplit('.').next() {
            if let Some(ct) = self.defaults.get(ext) {
                return Some(ct);
            }
        }
        
        None
    }

    pub fn set_content_type(&mut self, path: &str, content_type: &str) {
        self.overrides.insert(path.to_string(), content_type.to_string());
    }

    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults.insert(extension.to_string(), content_type.to_string());
    }
}

pub mod content_type_values {
    pub const WORD_DOCUMENT: &str = 
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub const WORD_STYLES: &str = 
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    pub const EXCEL_WORKBOOK: &str = 
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    pub const EXCEL_WORKSHEET: &str = 
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    pub const POWERPOINT_PRESENTATION: &str = 
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    pub const POWERPOINT_SLIDE: &str = 
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
}
