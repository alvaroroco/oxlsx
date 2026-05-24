pub(crate) fn column_index(reference: &str) -> Option<u32> {
    let mut col = 0u32;
    let mut seen = false;

    for b in reference.bytes() {
        if b.is_ascii_alphabetic() {
            seen = true;
            let upper = b.to_ascii_uppercase();
            col = col
                .checked_mul(26)?
                .checked_add(u32::from(upper - b'A' + 1))?;
        } else {
            break;
        }
    }

    if seen { Some(col) } else { None }
}

pub(crate) fn column_name(mut index: u32) -> String {
    if index == 0 {
        return String::new();
    }

    let mut chars = Vec::new();
    while index > 0 {
        index -= 1;
        chars.push((b'A' + (index % 26) as u8) as char);
        index /= 26;
    }
    chars.iter().rev().collect()
}

pub(crate) fn rc_to_a1(row: u32, col: u32) -> String {
    assert!(row >= 1, "row must be 1-based");
    assert!(col >= 1, "col must be 1-based");
    format!("{}{}", column_name(col), row)
}

#[cfg(test)]
mod tests {
    #[test]
    fn coords_round_trip_a1_conversions() {
        assert_eq!(super::column_index("A1"), Some(1));
        assert_eq!(super::column_index("Z9"), Some(26));
        assert_eq!(super::column_index("AA100"), Some(27));

        assert_eq!(super::column_name(1), "A");
        assert_eq!(super::column_name(26), "Z");
        assert_eq!(super::column_name(27), "AA");

        assert_eq!(super::rc_to_a1(1, 1), "A1");
        assert_eq!(super::rc_to_a1(10, 52), "AZ10");
    }
}
