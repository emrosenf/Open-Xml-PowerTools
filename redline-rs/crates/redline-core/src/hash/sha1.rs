use sha1::{Sha1, Digest};

pub fn sha1_hash_string(s: &str) -> String {
    sha1_hash_bytes(s.as_bytes())
}

pub fn sha1_hash_bytes(bytes: &[u8]) -> String {
    let mut hasher = Sha1::new();
    hasher.update(bytes);
    hex::encode(hasher.finalize())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sha1_empty_string() {
        assert_eq!(
            sha1_hash_string(""),
            "da39a3ee5e6b4b0d3255bfef95601890afd80709"
        );
    }

    #[test]
    fn sha1_test_string() {
        assert_eq!(
            sha1_hash_string("test"),
            "a94a8fe5ccb19ba61c4c0873d391e987982fbbd3"
        );
    }

    #[test]
    fn sha1_hello_world() {
        assert_eq!(
            sha1_hash_string("Hello, World!"),
            "0a0a9f2a6772942557ab5355d76af442f8f65e01"
        );
    }

    #[test]
    fn sha1_unicode_chinese() {
        let hash = sha1_hash_string("你好世界");
        assert_eq!(hash.len(), 40);
        assert!(hash.chars().all(|c| c.is_ascii_hexdigit()));
    }

    #[test]
    fn sha1_xml_content() {
        let hash = sha1_hash_string("<w:p>test</w:p>");
        assert_eq!(hash.len(), 40);
        assert!(hash.chars().all(|c| c.is_ascii_hexdigit()));
    }

    #[test]
    fn sha1_nbsp() {
        let hash = sha1_hash_string("\u{00A0}");
        assert_eq!(hash.len(), 40);
        assert!(hash.chars().all(|c| c.is_ascii_hexdigit()));
    }

    #[test]
    fn sha1_newlines() {
        let hash1 = sha1_hash_string("line1\nline2");
        let hash2 = sha1_hash_string("line1\r\nline2");
        assert_eq!(hash1.len(), 40);
        assert_eq!(hash2.len(), 40);
        assert_ne!(hash1, hash2);
    }
}
