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

pub(crate) fn parse_a1(reference: &str) -> Option<(u32, u32)> {
    let col = column_index(reference)?;
    let row_start = reference
        .bytes()
        .position(|b| !b.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    if row_start == reference.len() {
        return None;
    }

    let row = reference.get(row_start..)?.parse::<u32>().ok()?;
    if row == 0 {
        return None;
    }

    Some((row, col))
}

pub(crate) fn parse_range(range: &str) -> Option<((u32, u32), (u32, u32))> {
    let mut parts = range.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    if parts.next().is_some() {
        return None;
    }

    let (start_row, start_col) = parse_a1(start)?;
    let (end_row, end_col) = parse_a1(end)?;

    Some((
        (start_row.min(end_row), start_col.min(end_col)),
        (start_row.max(end_row), start_col.max(end_col)),
    ))
}

pub(crate) fn shift_a1(reference: &str, drow: i32, dcol: i32) -> Option<String> {
    let (row, col) = parse_a1(reference)?;
    let shifted_row = i64::from(row) + i64::from(drow);
    let shifted_col = i64::from(col) + i64::from(dcol);
    if shifted_row < 1 || shifted_col < 1 {
        return None;
    }

    let shifted_row = u32::try_from(shifted_row).ok()?;
    let shifted_col = u32::try_from(shifted_col).ok()?;
    Some(rc_to_a1(shifted_row, shifted_col))
}

pub(crate) fn shift_range(range: &str, drow: i32, dcol: i32) -> Option<String> {
    let ((start_row, start_col), (end_row, end_col)) = parse_range(range)?;
    let shifted_start = shift_a1(&rc_to_a1(start_row, start_col), drow, dcol)?;
    let shifted_end = shift_a1(&rc_to_a1(end_row, end_col), drow, dcol)?;
    Some(format!("{shifted_start}:{shifted_end}"))
}

#[cfg(feature = "python")]
pub(crate) enum ColumnKey<'a> {
    Letters(&'a str),
    Index(u32),
}

#[cfg(feature = "python")]
pub(crate) fn column_from_key(key: ColumnKey<'_>) -> Option<u32> {
    match key {
        ColumnKey::Letters(letters) => {
            if letters.is_empty() || !letters.bytes().all(|b| b.is_ascii_alphabetic()) {
                return None;
            }
            column_index(letters)
        }
        ColumnKey::Index(index) => {
            if index == 0 {
                None
            } else {
                Some(index)
            }
        }
    }
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

        assert_eq!(super::parse_a1("A1"), Some((1, 1)));
        assert_eq!(super::parse_a1("AA100"), Some((100, 27)));
    }

    #[test]
    fn coords_parse_range_and_shift() {
        assert_eq!(super::parse_range("A1:C3"), Some(((1, 1), (3, 3))));
        assert_eq!(super::parse_range("C3:A1"), Some(((1, 1), (3, 3))));
        assert_eq!(super::shift_a1("B2", 2, 3), Some("E4".to_string()));
        assert_eq!(super::shift_range("A1:B2", 1, 1), Some("B2:C3".to_string()));

        assert_eq!(super::shift_a1("A1", -1, 0), None);
        assert_eq!(super::shift_range("A1:B2", -1, 0), None);

        #[cfg(feature = "python")]
        {
            assert_eq!(
                super::column_from_key(super::ColumnKey::Letters("AA")),
                Some(27)
            );
            assert_eq!(super::column_from_key(super::ColumnKey::Index(3)), Some(3));
        }
    }
}
