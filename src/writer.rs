use crate::{Fill, Font, Worksheet};
use std::collections::HashMap;

#[allow(dead_code)]
pub(crate) struct CollectedStyles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub cell_xfs: Vec<(u32, u32)>,
    pub cell_style_ids: HashMap<(usize, String), u32>,
}

// Returns (row, col) for sorting cells in sheet XML.
#[allow(dead_code)]
pub(crate) fn cell_sort_key(reference: &str) -> (u32, u32) {
    let mut col = 0_u32;
    let mut row_text = String::new();

    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() {
            let upper = ch.to_ascii_uppercase() as u32;
            col = col * 26 + (upper - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_text.push(ch);
        }
    }

    let row = row_text.parse::<u32>().unwrap_or(0);
    (row, col)
}

#[allow(dead_code)]
pub(crate) fn collect_styles(sheets: &[Worksheet]) -> CollectedStyles {
    let styled_cells = sheets
        .iter()
        .map(|sheet| sheet.cells().len())
        .sum::<usize>();

    let mut fonts = Vec::with_capacity(1 + styled_cells);
    fonts.push(Font::default());

    let mut fills = Vec::with_capacity(2 + styled_cells);
    fills.push(Fill::none());
    fills.push(Fill::gray125());

    let mut cell_xfs = Vec::with_capacity(1 + styled_cells);
    cell_xfs.push((0_u32, 0_u32));

    let mut cell_style_ids = HashMap::with_capacity(styled_cells);

    for (sheet_idx, sheet) in sheets.iter().enumerate() {
        for (reference, cell) in sheet.cells() {
            if cell.write_font.is_none() && cell.write_fill.is_none() {
                continue;
            }

            let font = cell.write_font.clone().unwrap_or_default();
            let fill = cell.write_fill.clone().unwrap_or_else(Fill::none);

            let font_idx = find_or_push(&mut fonts, font) as u32;
            let fill_idx = find_or_push(&mut fills, fill) as u32;
            let style_id = find_or_push(&mut cell_xfs, (font_idx, fill_idx)) as u32;

            cell_style_ids.insert((sheet_idx, reference.clone()), style_id);
        }
    }

    CollectedStyles {
        fonts,
        fills,
        cell_xfs,
        cell_style_ids,
    }
}

#[allow(dead_code)]
fn find_or_push<T: PartialEq>(items: &mut Vec<T>, value: T) -> usize {
    if let Some(idx) = items.iter().position(|item| item == &value) {
        idx
    } else {
        items.push(value);
        items.len() - 1
    }
}

#[cfg(test)]
mod tests {
    use crate::{Fill, Font, Worksheet};

    #[test]
    fn cell_sort_key_orders_column_and_row_coordinates() {
        assert_eq!(super::cell_sort_key("A1"), (1, 1));
        assert_eq!(super::cell_sort_key("Z1"), (1, 26));
        assert_eq!(super::cell_sort_key("AA1"), (1, 27));
        assert_eq!(super::cell_sort_key("B10"), (10, 2));
    }

    #[test]
    fn collect_styles_defaults_without_styled_cells() {
        let mut ws = Worksheet::new("Sheet");
        ws.set_cell("A1", "Hello");

        let collected = super::collect_styles(&[ws]);

        assert_eq!(collected.fonts.len(), 1);
        assert_eq!(collected.fills.len(), 2);
        assert_eq!(collected.cell_xfs.len(), 1);
        assert!(collected.cell_style_ids.is_empty());
    }

    #[test]
    fn collect_styles_dedupes_repeated_font_styles() {
        let mut ws = Worksheet::new("Sheet");

        ws.set_cell_font(
            "A1",
            Font {
                bold: true,
                ..Default::default()
            },
        );
        ws.set_cell_font(
            "A2",
            Font {
                bold: true,
                ..Default::default()
            },
        );
        ws.set_cell_font(
            "A3",
            Font {
                bold: true,
                ..Default::default()
            },
        );

        let collected = super::collect_styles(&[ws]);

        assert_eq!(collected.fonts.len(), 2);
        assert_eq!(collected.fills.len(), 2);
        assert_eq!(collected.cell_xfs.len(), 2);
        assert_eq!(collected.cell_style_ids.len(), 3);
    }

    #[test]
    fn collect_styles_adds_custom_fill_after_defaults() {
        let mut ws = Worksheet::new("Sheet");

        ws.set_cell_fill(
            "A1",
            Fill {
                pattern_type: "solid".to_string(),
                ..Default::default()
            },
        );

        let collected = super::collect_styles(&[ws]);

        assert_eq!(collected.fills.len(), 3);
    }
}
