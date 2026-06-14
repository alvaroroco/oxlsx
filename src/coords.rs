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

/// Make a single cell reference absolute (`A1` -> `$A$1`).
///
/// References already starting with `$`, or that are not a clean
/// `<letters><digits>` cell, are returned unchanged.
pub(crate) fn make_cell_absolute(cell: &str) -> String {
    if cell.starts_with('$') {
        return cell.to_string();
    }

    // Find where letters end and numbers begin
    let split_pos = cell
        .bytes()
        .position(|b| b.is_ascii_digit())
        .unwrap_or(cell.len());

    if split_pos == 0 || split_pos == cell.len() {
        return cell.to_string();
    }

    let col = &cell[..split_pos];
    let row = &cell[split_pos..];

    format!("${col}${row}")
}

/// Absolutize a single (non comma-separated) print range, auto-detecting its
/// shape: row (`1:3`), column (`A:C`), area (`A1:D4`), or single cell (`A1`).
///
/// A range already containing `$` is returned unchanged. This is the unified
/// normalizer used by the print-setting code path.
pub(crate) fn make_absolute_range(range: &str) -> String {
    if range.contains('$') {
        return range.to_string();
    }

    if !range.contains(':') {
        return make_cell_absolute(range);
    }

    if let Some((start, end)) = range.split_once(':') {
        // Check if this is rows (digits only) or cols (letters only) or area (mixed)
        if start.bytes().all(|b| b.is_ascii_digit()) {
            // Rows like "1:3"
            format!("${start}:${end}")
        } else if start.bytes().all(|b| b.is_ascii_alphabetic()) {
            // Cols like "A:C"
            format!("${start}:${end}")
        } else {
            // Area like "A1:D4"
            format!("{}:{}", make_cell_absolute(start), make_cell_absolute(end))
        }
    } else {
        range.to_string()
    }
}

/// Absolutize a row (`1:3`) or column (`A:C`) line range by prefixing both
/// endpoints with `$`. Ranges already starting with `$`, or without a `:`,
/// are returned unchanged.
///
/// Callers are expected to pass an already-classified row or column range
/// (see `is_row_segment` / `is_col_segment`).
pub(crate) fn make_absolute_line_range(range: &str) -> String {
    if range.starts_with('$') {
        range.to_string()
    } else if let Some((start, end)) = range.split_once(':') {
        format!("${start}:${end}")
    } else {
        range.to_string()
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
    }

    #[test]
    fn make_cell_absolute_prefixes_letters_and_digits() {
        assert_eq!(super::make_cell_absolute("A1"), "$A$1");
        assert_eq!(super::make_cell_absolute("$A$1"), "$A$1");
        // Not a clean letters+digits cell -> returned unchanged.
        assert_eq!(super::make_cell_absolute("1"), "1");
        assert_eq!(super::make_cell_absolute("A"), "A");
    }

    #[test]
    fn make_absolute_range_detects_row_col_and_area() {
        assert_eq!(super::make_absolute_range("1:3"), "$1:$3");
        assert_eq!(super::make_absolute_range("A:C"), "$A:$C");
        assert_eq!(super::make_absolute_range("A1:D4"), "$A$1:$D$4");
        assert_eq!(super::make_absolute_range("A1"), "$A$1");
        // Already absolute -> unchanged.
        assert_eq!(super::make_absolute_range("$A$1:$C$10"), "$A$1:$C$10");
    }

    #[test]
    fn make_absolute_line_range_prefixes_endpoints() {
        assert_eq!(super::make_absolute_line_range("1:3"), "$1:$3");
        assert_eq!(super::make_absolute_line_range("A:C"), "$A:$C");
        assert_eq!(super::make_absolute_line_range("$1:$3"), "$1:$3");
    }
}
