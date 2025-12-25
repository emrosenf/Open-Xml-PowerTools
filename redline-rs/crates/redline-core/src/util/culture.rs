pub fn to_upper_invariant(s: &str) -> String {
    s.to_uppercase()
}

pub fn to_upper_culture(s: &str, _culture: &str) -> String {
    s.to_uppercase()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn to_upper_invariant_basic() {
        assert_eq!(to_upper_invariant("hello"), "HELLO");
        assert_eq!(to_upper_invariant("Hello World"), "HELLO WORLD");
    }

    #[test]
    fn to_upper_invariant_unicode() {
        assert_eq!(to_upper_invariant("café"), "CAFÉ");
    }
}
