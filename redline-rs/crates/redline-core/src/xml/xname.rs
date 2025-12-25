use std::fmt;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct XName {
    pub namespace: Option<String>,
    pub local_name: String,
}

impl XName {
    pub fn new(namespace: &str, local_name: &str) -> Self {
        Self {
            namespace: if namespace.is_empty() { 
                None 
            } else { 
                Some(namespace.to_string()) 
            },
            local_name: local_name.to_string(),
        }
    }

    pub fn local(local_name: &str) -> Self {
        Self {
            namespace: None,
            local_name: local_name.to_string(),
        }
    }
}

impl fmt::Display for XName {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.namespace {
            Some(ns) => write!(f, "{{{}}}{}", ns, self.local_name),
            None => write!(f, "{}", self.local_name),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XAttribute {
    pub name: XName,
    pub value: String,
}

impl XAttribute {
    pub fn new(name: XName, value: &str) -> Self {
        Self {
            name,
            value: value.to_string(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xname_with_namespace_displays_correctly() {
        let name = XName::new("http://example.com", "element");
        assert_eq!(name.to_string(), "{http://example.com}element");
    }

    #[test]
    fn xname_without_namespace_displays_correctly() {
        let name = XName::local("element");
        assert_eq!(name.to_string(), "element");
    }

    #[test]
    fn xattribute_creates_correctly() {
        let attr = XAttribute::new(XName::local("id"), "test123");
        assert_eq!(attr.value, "test123");
        assert_eq!(attr.name.local_name, "id");
    }
}
