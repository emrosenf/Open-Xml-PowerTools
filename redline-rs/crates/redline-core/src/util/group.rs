pub fn group_adjacent<T, K, F>(items: impl Iterator<Item = T>, key_selector: F) -> Vec<Vec<T>>
where
    K: Eq,
    F: Fn(&T) -> K,
{
    let mut result: Vec<Vec<T>> = Vec::new();
    let mut current_group: Option<(K, Vec<T>)> = None;

    for item in items {
        let key = key_selector(&item);
        
        match &mut current_group {
            Some((current_key, group)) if *current_key == key => {
                group.push(item);
            }
            _ => {
                if let Some((_, group)) = current_group.take() {
                    result.push(group);
                }
                current_group = Some((key, vec![item]));
            }
        }
    }

    if let Some((_, group)) = current_group {
        result.push(group);
    }

    result
}

pub fn rollup<T, R, F>(items: impl Iterator<Item = T>, seed: R, folder: F) -> Vec<R>
where
    R: Clone,
    F: Fn(R, &T) -> R,
{
    let mut result = Vec::new();
    let mut current = seed;

    for item in items {
        current = folder(current, &item);
        result.push(current.clone());
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn group_adjacent_groups_consecutive_equal_keys() {
        let items = vec![1, 1, 2, 2, 2, 1, 3, 3];
        let groups = group_adjacent(items.into_iter(), |&x| x);
        
        assert_eq!(groups.len(), 4);
        assert_eq!(groups[0], vec![1, 1]);
        assert_eq!(groups[1], vec![2, 2, 2]);
        assert_eq!(groups[2], vec![1]);
        assert_eq!(groups[3], vec![3, 3]);
    }

    #[test]
    fn group_adjacent_with_strings() {
        let items = vec!["aa", "ab", "ba", "bb"];
        let groups = group_adjacent(items.into_iter(), |s| s.chars().next().unwrap());
        
        assert_eq!(groups.len(), 2);
        assert_eq!(groups[0], vec!["aa", "ab"]);
        assert_eq!(groups[1], vec!["ba", "bb"]);
    }

    #[test]
    fn rollup_accumulates_values() {
        let items = vec![1, 2, 3, 4];
        let sums = rollup(items.into_iter(), 0, |acc, &x| acc + x);
        
        assert_eq!(sums, vec![1, 3, 6, 10]);
    }
}
