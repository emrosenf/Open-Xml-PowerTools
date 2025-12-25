pub fn make_valid_xml(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    
    for c in s.chars() {
        if is_valid_xml_char(c) {
            result.push(c);
        } else {
            result.push('\u{FFFD}');
        }
    }
    
    result
}

fn is_valid_xml_char(c: char) -> bool {
    matches!(c,
        '\u{0009}' | '\u{000A}' | '\u{000D}' |
        '\u{0020}'..='\u{D7FF}' |
        '\u{E000}'..='\u{FFFD}' |
        '\u{10000}'..='\u{10FFFF}'
    )
}

pub fn string_concatenate<I, S>(items: I, separator: &str) -> String
where
    I: Iterator<Item = S>,
    S: AsRef<str>,
{
    let mut result = String::new();
    let mut first = true;
    
    for item in items {
        if !first {
            result.push_str(separator);
        }
        result.push_str(item.as_ref());
        first = false;
    }
    
    result
}

pub fn normalize_spaces(s: &str, conflate_nbsp: bool) -> String {
    if conflate_nbsp {
        s.replace('\u{00A0}', " ")
    } else {
        s.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn make_valid_xml_replaces_invalid_chars() {
        let input = "hello\u{0000}world";
        let result = make_valid_xml(input);
        assert_eq!(result, "hello\u{FFFD}world");
    }

    #[test]
    fn make_valid_xml_preserves_valid_chars() {
        let input = "hello\tworld\n";
        let result = make_valid_xml(input);
        assert_eq!(result, input);
    }

    #[test]
    fn string_concatenate_joins_with_separator() {
        let items = vec!["a", "b", "c"];
        let result = string_concatenate(items.into_iter(), ", ");
        assert_eq!(result, "a, b, c");
    }

    #[test]
    fn string_concatenate_empty_iterator() {
        let items: Vec<&str> = vec![];
        let result = string_concatenate(items.into_iter(), ", ");
        assert_eq!(result, "");
    }

    #[test]
    fn normalize_spaces_replaces_nbsp() {
        let input = "hello\u{00A0}world";
        let result = normalize_spaces(input, true);
        assert_eq!(result, "hello world");
    }

    #[test]
    fn normalize_spaces_preserves_nbsp_when_disabled() {
        let input = "hello\u{00A0}world";
        let result = normalize_spaces(input, false);
        assert_eq!(result, input);
    }
}
