use sha2::{Sha256, Digest};

pub fn sha256_hash_string(s: &str) -> String {
    sha256_hash_bytes(s.as_bytes())
}

pub fn sha256_hash_bytes(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hex::encode(hasher.finalize())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sha256_empty_string() {
        assert_eq!(
            sha256_hash_string(""),
            "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        );
    }

    #[test]
    fn sha256_test_string() {
        assert_eq!(
            sha256_hash_string("test"),
            "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08"
        );
    }

    #[test]
    fn sha256_hello_world() {
        assert_eq!(
            sha256_hash_string("Hello, World!"),
            "dffd6021bb2bd5b0af676290809ec3a53191dd81c7f70a4b28688a362182986f"
        );
    }

    #[test]
    fn sha256_unicode_chinese() {
        let hash = sha256_hash_string("你好世界");
        assert_eq!(hash.len(), 64);
        assert!(hash.chars().all(|c| c.is_ascii_hexdigit()));
    }

    #[test]
    fn sha256_xml_content() {
        let hash = sha256_hash_string("<x:row><x:c><x:v>123</x:v></x:c></x:row>");
        assert_eq!(hash.len(), 64);
        assert!(hash.chars().all(|c| c.is_ascii_hexdigit()));
    }
}
