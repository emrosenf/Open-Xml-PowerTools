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
        // Verified against C# .NET 10.0.101
        assert_eq!(
            sha1_hash_string("你好世界"),
            "dabaa5fe7c47fb21be902480a13013f16a1ab6eb"
        );
    }

    #[test]
    fn sha1_xml_content() {
        // Verified against C# .NET 10.0.101
        assert_eq!(
            sha1_hash_string("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"),
            "307d14f5780b72d87b6cf5ecaf78c7430200d1ae"
        );
    }

    #[test]
    fn sha1_nbsp() {
        // Verified against C# .NET 10.0.101
        assert_eq!(
            sha1_hash_string("\u{00A0}"),
            "ab90d23f7402359d51e25399fe46dac3401a3352"
        );
    }

    #[test]
    fn sha1_unix_newlines() {
        // Verified against C# .NET 10.0.101
        assert_eq!(
            sha1_hash_string("line1\nline2"),
            "05eed6236c8bda5ecf7af09bae911f9d5f90998b"
        );
    }

    #[test]
    fn sha1_windows_newlines() {
        // Verified against C# .NET 10.0.101
        assert_eq!(
            sha1_hash_string("line1\r\nline2"),
            "2e8b459e11acdf2861942e27e7651513578e8c7d"
        );
    }

    #[test]
    fn sha1_classic_test_phrase() {
        // Verified against C# .NET 10.0.101
        assert_eq!(
            sha1_hash_string("The quick brown fox jumps over the lazy dog"),
            "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"
        );
    }
}
