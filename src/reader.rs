use crate::comments::{Comment, parse_comments_xml};
use crate::coords::{column_index, make_absolute_range, rc_to_a1};
use crate::defined_names::parse_workbook_defined_names;
use crate::sheet_rels::{
    COMMENTS_REL_TYPE, SheetRelEntry, TABLE_REL_TYPE, parse_sheet_rels, relation_target_for_id,
    relation_target_for_type, relation_targets_for_type, resolve_sheet_target, sheet_rels_path,
};
use crate::table::{RawTable, Table, derive_columns_from_header, parse_table_xml};
use crate::workbook_rels::parse_sheet_refs;
use crate::xml::attr_value;
use crate::{
    Fill, Font, Hyperlink, OxlsxError, ReadonlyWorkbook, SharedStrings, StyleId, StyleSheet,
};
use chrono::{Duration, NaiveDate};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read};
use std::mem;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use zip::result::ZipError;

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    String(String),
    Number(f64),
    Bool(bool),
    Date(NaiveDate),
    Formula(String),
}

impl From<&str> for Value {
    fn from(s: &str) -> Self {
        Value::String(s.to_string())
    }
}

impl From<String> for Value {
    fn from(s: String) -> Self {
        Value::String(s)
    }
}

impl From<f64> for Value {
    fn from(n: f64) -> Self {
        Value::Number(n)
    }
}

impl From<i64> for Value {
    fn from(n: i64) -> Self {
        Value::Number(n as f64)
    }
}

impl From<bool> for Value {
    fn from(b: bool) -> Self {
        Value::Bool(b)
    }
}

impl From<NaiveDate> for Value {
    fn from(d: NaiveDate) -> Self {
        Value::Date(d)
    }
}

#[derive(Debug, Clone)]
pub struct Cell {
    reference: String,
    value: Value,
    style_id: StyleId,
    styles: Arc<StyleSheet>,
    pub(crate) write_font: Option<Font>,
    pub(crate) write_fill: Option<Fill>,
    pub(crate) comment: Option<Comment>,
    pub(crate) write_hyperlink: Option<Hyperlink>,
}

impl Cell {
    pub fn reference(&self) -> &str {
        &self.reference
    }

    pub fn value(&self) -> &Value {
        &self.value
    }

    pub fn style_id(&self) -> StyleId {
        self.style_id
    }

    pub fn font(&self) -> Option<&Font> {
        self.styles.font(self.style_id)
    }

    pub fn fill(&self) -> Option<&Fill> {
        self.styles.fill(self.style_id)
    }

    pub fn text(&self) -> Option<&str> {
        match &self.value {
            Value::String(s) => Some(s.as_str()),
            _ => None,
        }
    }

    pub fn comment(&self) -> Option<&Comment> {
        self.comment.as_ref()
    }

    /// The cell's hyperlink, if any.
    ///
    /// Mirrors openpyxl's `cell.hyperlink`. External links carry a non-empty
    /// `target`; internal `location`-only links and `tooltip`/`display` are
    /// preserved verbatim across read-modify-write.
    pub fn hyperlink(&self) -> Option<&Hyperlink> {
        self.write_hyperlink.as_ref()
    }

    pub(crate) fn parsed(
        reference: String,
        value: Value,
        style_id: StyleId,
        styles: Arc<StyleSheet>,
    ) -> Self {
        Self {
            reference,
            value,
            style_id,
            styles,
            write_font: None,
            write_fill: None,
            comment: None,
            write_hyperlink: None,
        }
    }

    pub(crate) fn empty(reference: String, styles: Arc<StyleSheet>) -> Self {
        Self::parsed(reference, Value::Empty, StyleId::default(), styles)
    }
}

#[derive(Debug, Clone, Default)]
pub struct Worksheet {
    name: String,
    cells: HashMap<String, Cell>,
    pub(crate) merged_cells: Vec<String>,
    pub(crate) column_widths: HashMap<u32, f64>,
    pub(crate) row_heights: HashMap<u32, f64>,
    pub(crate) freeze_panes: Option<String>,
    pub(crate) print_area: Option<String>,
    pub(crate) print_title_rows: Option<String>,
    pub(crate) print_title_cols: Option<String>,
    defined_names: crate::DefinedNames,
    pub(crate) tables: Vec<Table>,
    /// Retained raw `tableN.xml` bytes, index-aligned 1:1 with `tables`.
    ///
    /// `Some` when the table was read from a source workbook (verbatim
    /// candidate); `None` for tables created via [`Worksheet::add_table`].
    /// Kept in lock-step with `tables` by every mutation (`add_table`,
    /// `remove_table`).
    pub(crate) table_retention: Vec<Option<RawTable>>,
}

impl Worksheet {
    pub(crate) fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            cells: HashMap::new(),
            merged_cells: Vec::new(),
            column_widths: HashMap::new(),
            row_heights: HashMap::new(),
            freeze_panes: None,
            print_area: None,
            print_title_rows: None,
            print_title_cols: None,
            defined_names: crate::DefinedNames::default(),
            tables: Vec::new(),
            table_retention: Vec::new(),
        }
    }

    pub fn title(&self) -> &str {
        &self.name
    }

    /// Sheet-local named ranges (`<definedName>` carrying this sheet's `localSheetId`).
    ///
    /// Separate from the workbook-global collection ([`Workbook::defined_names`]);
    /// the same name MAY exist in both scopes (openpyxl parity).
    pub fn defined_names(&self) -> &crate::DefinedNames {
        &self.defined_names
    }

    /// Mutable access to this sheet's local named ranges.
    ///
    /// `localSheetId` is never stored: it is computed from the worksheet's live
    /// position at serialize time.
    pub fn defined_names_mut(&mut self) -> &mut crate::DefinedNames {
        &mut self.defined_names
    }

    /// Excel tables (`<table>` parts) defined on this worksheet.
    ///
    /// Mirrors openpyxl's `ws.tables`; lookup by `display_name` is done by the
    /// caller (the Python layer exposes the dict-like interface).
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Mutable access to a table by `display_name` (case-insensitive lookup,
    /// matching the workbook-wide `displayName` uniqueness rule). Returns
    /// `None` if no table with that name exists on this worksheet.
    ///
    /// Acquiring this mutable reference conservatively marks the table
    /// **dirty**: on save it takes the model serialization path, discarding any
    /// retained verbatim bytes (and the unmodeled features they carried). Use
    /// [`tables`](Self::tables) for read-only access that preserves verbatim
    /// retention.
    pub fn table_mut(&mut self, display_name: &str) -> Option<&mut Table> {
        let idx = self
            .tables
            .iter()
            .position(|table| table.display_name.eq_ignore_ascii_case(display_name))?;
        if let Some(Some(raw)) = self.table_retention.get_mut(idx) {
            raw.dirty = true;
        }
        Some(&mut self.tables[idx])
    }

    /// Remove a table by `display_name` (case-insensitive), returning the
    /// removed [`Table`] or `None` if no such table exists.
    ///
    /// Drops both the table and its retention slot, keeping `tables` and
    /// `table_retention` index-aligned. On save the removed table's part, sheet
    /// rel, `<tablePart>` entry, and content-type override all disappear.
    pub fn remove_table(&mut self, display_name: &str) -> Option<Table> {
        let idx = self
            .tables
            .iter()
            .position(|table| table.display_name.eq_ignore_ascii_case(display_name))?;
        if idx < self.table_retention.len() {
            self.table_retention.remove(idx);
        }
        Some(self.tables.remove(idx))
    }

    /// Add an Excel table to this worksheet, mirroring openpyxl's
    /// `ws.add_table`.
    ///
    /// When `table.columns` is empty, columns are derived from the header row
    /// of `table.ref_range` (ids `1..=N`, names from the header cells). When
    /// `table.name` is empty it defaults to `table.display_name`.
    ///
    /// Add a table to this worksheet.
    ///
    /// Returns [`OxlsxError::DuplicateTableName`] if a table with the same
    /// `display_name` (case-insensitive) already exists on THIS worksheet; no
    /// state is mutated on rejection. Workbook-wide uniqueness is enforced at
    /// save time.
    ///
    /// Returns [`OxlsxError::InvalidFormat`] if the table's `ref_range` does not
    /// parse to at least one column.
    ///
    /// A table added this way has no retained source bytes, so it is always
    /// serialized from the model on save (there is nothing to preserve
    /// verbatim). For tables read from an existing workbook, unmodeled features
    /// survive RMW unless the table is mutated — see the [`crate::table`] module
    /// doc.
    pub fn add_table(&mut self, mut table: Table) -> Result<(), OxlsxError> {
        // FIX 6: Case-insensitive same-sheet check.
        let lower = table.display_name.to_ascii_lowercase();
        if self
            .tables
            .iter()
            .any(|existing| existing.display_name.to_ascii_lowercase() == lower)
        {
            return Err(OxlsxError::DuplicateTableName(table.display_name));
        }

        // FIX 2: Validate that the ref_range has at least one column.
        if !crate::table::validate_table_ref(&table.ref_range) {
            return Err(OxlsxError::InvalidFormat(format!(
                "table ref_range {:?} does not parse to at least one column",
                table.ref_range
            )));
        }

        if table.name.is_empty() {
            table.name = table.display_name.clone();
        }
        if table.columns.is_empty() {
            table.columns = derive_columns_from_header(self, &table.ref_range);
        }

        self.tables.push(table);
        // New tables have no source bytes → model serialization path.
        self.table_retention.push(None);
        Ok(())
    }

    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        self.cells.get(reference)
    }

    pub(crate) fn cells(&self) -> &HashMap<String, Cell> {
        &self.cells
    }

    pub fn set_cell(&mut self, reference: &str, value: impl Into<Value>) {
        let cell = self.ensure_cell_mut(reference);
        cell.value = value.into();
    }

    pub fn set_cell_font(&mut self, reference: &str, font: Font) {
        let cell = self.ensure_cell_mut(reference);
        cell.write_font = Some(font);
    }

    pub fn set_cell_fill(&mut self, reference: &str, fill: Fill) {
        let cell = self.ensure_cell_mut(reference);
        cell.write_fill = Some(fill);
    }

    pub fn set_cell_comment(&mut self, reference: &str, comment: Comment) {
        let cell = self.ensure_cell_mut(reference);
        cell.comment = Some(comment);
    }

    /// Attach an external-URL hyperlink to `reference`, creating the cell if
    /// needed. Mirrors openpyxl's `cell.hyperlink = "https://..."`.
    ///
    /// An empty `target` is a no-op: an empty hyperlink emits nothing, matching
    /// the spec ("An empty target MUST NOT emit a hyperlink").
    pub fn set_cell_hyperlink(&mut self, reference: &str, target: &str) {
        if target.is_empty() {
            return;
        }
        let cell = self.ensure_cell_mut(reference);
        cell.write_hyperlink = Some(Hyperlink::new_external(target));
    }

    pub fn merge_cells(&mut self, range: &str) {
        let range = range.to_string();
        if !self.merged_cells.contains(&range) {
            self.merged_cells.push(range);
        }
    }

    pub fn unmerge_cells(&mut self, range: &str) {
        self.merged_cells.retain(|existing| existing != range);
    }

    pub fn column_width(&mut self, col: u32, width: f64) {
        self.column_widths.insert(col, width);
    }

    pub fn row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
    }

    pub fn freeze_panes(&self) -> Option<&str> {
        self.freeze_panes.as_deref()
    }

    pub fn set_freeze_panes(&mut self, value: Option<&str>) {
        self.freeze_panes = normalize_freeze_panes(value);
    }

    pub fn print_area(&self) -> Option<String> {
        self.print_area.as_ref().map(|area| {
            // Return qualified form like openpyxl: 'Sheet'!$A$1:$D$4
            let quoted_title = format!("'{}'", self.name.replace('\'', "''"));
            if area.contains(',') {
                // Multi-range: qualify each segment
                area.split(',')
                    .map(|segment| format!("{quoted_title}!{}", segment.trim()))
                    .collect::<Vec<_>>()
                    .join(",")
            } else {
                format!("{quoted_title}!{area}")
            }
        })
    }

    pub(crate) fn print_area_bare(&self) -> Option<&str> {
        self.print_area.as_deref()
    }

    pub fn set_print_area(&mut self, value: Option<&str>) {
        self.print_area = normalize_print_setting(value);
    }

    pub fn print_title_rows(&self) -> Option<&str> {
        self.print_title_rows.as_deref()
    }

    pub fn set_print_title_rows(&mut self, value: Option<&str>) {
        self.print_title_rows = normalize_print_setting(value);
    }

    pub fn print_title_cols(&self) -> Option<&str> {
        self.print_title_cols.as_deref()
    }

    pub fn set_print_title_cols(&mut self, value: Option<&str>) {
        self.print_title_cols = normalize_print_setting(value);
    }

    pub fn max_row(&self) -> u32 {
        self.used_bounds().map_or(1, |(_, _, _, max_row)| max_row)
    }

    pub fn max_column(&self) -> u32 {
        self.used_bounds().map_or(1, |(_, max_col, _, _)| max_col)
    }

    pub fn min_row(&self) -> u32 {
        self.used_bounds().map_or(1, |(_, _, min_row, _)| min_row)
    }

    pub fn min_column(&self) -> u32 {
        self.used_bounds().map_or(1, |(min_col, _, _, _)| min_col)
    }

    pub fn iter_rows(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> Vec<Vec<Cell>> {
        let default_min_row = self.min_row();
        let default_max_row = self.max_row();
        let default_min_col = self.min_column();
        let default_max_col = self.max_column();

        let min_row = min_row.unwrap_or(default_min_row);
        let max_row = max_row.unwrap_or(default_max_row);
        let min_col = min_col.unwrap_or(default_min_col);
        let max_col = max_col.unwrap_or(default_max_col);

        if min_row > max_row || min_col > max_col {
            return Vec::new();
        }

        let styles = Arc::new(StyleSheet::default());
        let mut rows = Vec::new();
        for row in min_row..=max_row {
            let mut current = Vec::new();
            for col in min_col..=max_col {
                let reference = rc_to_a1(row, col);
                current.push(
                    self.cells
                        .get(&reference)
                        .cloned()
                        .unwrap_or_else(|| Cell::empty(reference, styles.clone())),
                );
            }
            rows.push(current);
        }

        rows
    }

    pub fn iter_cols(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> Vec<Vec<Cell>> {
        let default_min_row = self.min_row();
        let default_max_row = self.max_row();
        let default_min_col = self.min_column();
        let default_max_col = self.max_column();

        let min_row = min_row.unwrap_or(default_min_row);
        let max_row = max_row.unwrap_or(default_max_row);
        let min_col = min_col.unwrap_or(default_min_col);
        let max_col = max_col.unwrap_or(default_max_col);

        if min_row > max_row || min_col > max_col {
            return Vec::new();
        }

        let styles = Arc::new(StyleSheet::default());
        let mut cols = Vec::new();
        for col in min_col..=max_col {
            let mut current = Vec::new();
            for row in min_row..=max_row {
                let reference = rc_to_a1(row, col);
                current.push(
                    self.cells
                        .get(&reference)
                        .cloned()
                        .unwrap_or_else(|| Cell::empty(reference, styles.clone())),
                );
            }
            cols.push(current);
        }

        cols
    }

    pub fn rows(&self) -> Vec<Vec<Cell>> {
        self.iter_rows(None, None, None, None)
    }

    pub fn columns(&self) -> Vec<Vec<Cell>> {
        self.iter_cols(None, None, None, None)
    }

    pub fn values(&self) -> Vec<Vec<Value>> {
        self.iter_rows(None, None, None, None)
            .into_iter()
            .map(|row| row.into_iter().map(|cell| cell.value().clone()).collect())
            .collect()
    }

    fn used_bounds(&self) -> Option<(u32, u32, u32, u32)> {
        let mut min_col = u32::MAX;
        let mut max_col = 0;
        let mut min_row = u32::MAX;
        let mut max_row = 0;

        for reference in self.cells.keys() {
            let Some((row, col)) = parse_a1(reference) else {
                continue;
            };

            min_col = min_col.min(col);
            max_col = max_col.max(col);
            min_row = min_row.min(row);
            max_row = max_row.max(row);
        }

        if max_row == 0 || max_col == 0 {
            None
        } else {
            Some((min_col, max_col, min_row, max_row))
        }
    }

    fn ensure_cell_mut(&mut self, reference: &str) -> &mut Cell {
        self.cells
            .entry(reference.to_string())
            .or_insert_with(|| Cell {
                reference: reference.to_string(),
                value: Value::Empty,
                style_id: StyleId::default(),
                styles: Arc::new(StyleSheet::default()),
                write_font: None,
                write_fill: None,
                comment: None,
                write_hyperlink: None,
            })
    }
}

fn parse_a1(reference: &str) -> Option<(u32, u32)> {
    let col = column_index(reference)?;
    let row_start = reference
        .bytes()
        .position(|b| !b.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    if row_start == reference.len() {
        return None;
    }

    let row = reference.get(row_start..)?.parse::<u32>().ok()?;
    Some((row, col))
}

#[derive(Debug, Clone)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
    active_index: usize,
    pub(crate) source_path: Option<PathBuf>,
    pub(crate) original_styles: Option<Arc<StyleSheet>>,
    pub(crate) original_sheet_rels: Vec<Vec<SheetRelEntry>>,
    pub(crate) extra_defined_names: Vec<String>,
    defined_names: crate::DefinedNames,
}

impl Workbook {
    pub fn new() -> Self {
        Self {
            sheets: vec![Worksheet::new("Sheet")],
            active_index: 0,
            source_path: None,
            original_styles: None,
            original_sheet_rels: vec![Vec::new()],
            extra_defined_names: Vec::new(),
            defined_names: crate::DefinedNames::default(),
        }
    }

    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OxlsxError> {
        let source_path = path.as_ref().to_path_buf();
        let file = File::open(&source_path)?;
        let mut archive = zip::ZipArchive::new(file).map_err(|err| match err {
            ZipError::InvalidArchive(_) | ZipError::UnsupportedArchive(_) => {
                OxlsxError::InvalidFormat(err.to_string())
            }
            _ => OxlsxError::Zip(err),
        })?;

        let styles = {
            let styles_file = archive
                .by_name("xl/styles.xml")
                .map_err(|_| OxlsxError::MissingPart("xl/styles.xml".to_string()))?;
            Arc::new(StyleSheet::parse(BufReader::new(styles_file))?)
        };

        let shared_strings: Vec<String> = match archive.by_name("xl/sharedStrings.xml") {
            Ok(f) => SharedStrings::parse(BufReader::new(f))?.into_inner(),
            Err(ZipError::FileNotFound) => Vec::new(),
            Err(e) => return Err(OxlsxError::Zip(e)),
        };

        let (
            sheet_refs,
            active_index,
            print_defined_names,
            sheet_local_defined_names,
            defined_names,
            extra_defined_names,
        ) = {
            let workbook_xml = {
                let mut wb_file = archive
                    .by_name("xl/workbook.xml")
                    .map_err(|_| OxlsxError::MissingPart("xl/workbook.xml".to_string()))?;
                let mut bytes = Vec::new();
                wb_file.read_to_end(&mut bytes)?;
                bytes
            };

            let rels_xml = {
                let mut rels_file =
                    archive.by_name("xl/_rels/workbook.xml.rels").map_err(|_| {
                        OxlsxError::MissingPart("xl/_rels/workbook.xml.rels".to_string())
                    })?;
                let mut bytes = Vec::new();
                rels_file.read_to_end(&mut bytes)?;
                bytes
            };

            let (sheet_refs, active_index) = parse_sheet_refs(
                BufReader::new(Cursor::new(workbook_xml.clone())),
                BufReader::new(Cursor::new(rels_xml)),
            )?;
            let (
                print_defined_names,
                sheet_local_defined_names,
                defined_names,
                extra_defined_names,
            ) = parse_workbook_defined_names(
                BufReader::new(Cursor::new(workbook_xml)),
                sheet_refs.len(),
            )?;

            (
                sheet_refs,
                active_index,
                print_defined_names,
                sheet_local_defined_names,
                defined_names,
                extra_defined_names,
            )
        };

        let mut sheets = Vec::with_capacity(sheet_refs.len());
        let mut original_sheet_rels = Vec::with_capacity(sheet_refs.len());
        for sheet_ref in &sheet_refs {
            let sheet_file = archive
                .by_name(&sheet_ref.zip_path)
                .map_err(|_| OxlsxError::MissingPart(sheet_ref.zip_path.clone()))?;
            let (mut ws, raw_hyperlinks) =
                parse_sheet(BufReader::new(sheet_file), styles.clone(), &shared_strings)?;
            ws.name = sheet_ref.name.clone();

            let rels_path = sheet_rels_path(&sheet_ref.zip_path);
            let sheet_rels = match archive.by_name(&rels_path) {
                Ok(rels_file) => parse_sheet_rels(BufReader::new(rels_file))?,
                Err(ZipError::FileNotFound) => Vec::new(),
                Err(err) => return Err(OxlsxError::Zip(err)),
            };

            if let Some(comment_target) = relation_target_for_type(&sheet_rels, COMMENTS_REL_TYPE) {
                let comments_path = resolve_sheet_target(&sheet_ref.zip_path, comment_target);
                let parsed_comments = match archive.by_name(&comments_path) {
                    Ok(comments_file) => parse_comments_xml(BufReader::new(comments_file))?,
                    Err(ZipError::FileNotFound) => HashMap::new(),
                    Err(err) => return Err(OxlsxError::Zip(err)),
                };

                for (reference, comment) in parsed_comments {
                    let cell = ws.cells.entry(reference.clone()).or_insert_with(|| {
                        Cell::parsed(
                            reference.clone(),
                            Value::Empty,
                            StyleId::default(),
                            styles.clone(),
                        )
                    });
                    cell.comment = Some(comment);
                }
            }

            for table_target in relation_targets_for_type(&sheet_rels, TABLE_REL_TYPE) {
                let table_path = resolve_sheet_target(&sheet_ref.zip_path, table_target);
                match archive.by_name(&table_path) {
                    Ok(mut table_file) => {
                        // Retain the exact bytes for verbatim re-emission, then
                        // parse the thin model from the same bytes via a Cursor.
                        let mut bytes = Vec::new();
                        table_file.read_to_end(&mut bytes)?;
                        let table = parse_table_xml(Cursor::new(bytes.as_slice()))?;
                        ws.tables.push(table);
                        ws.table_retention.push(Some(RawTable {
                            bytes,
                            dirty: false,
                        }));
                    }
                    Err(ZipError::FileNotFound) => {}
                    Err(err) => return Err(OxlsxError::Zip(err)),
                }
            }

            // Resolve each raw hyperlink now that the sheet's rels are loaded.
            // External links resolve `r:id` → rel target (the URL); internal
            // `location`-only links keep an empty target and fabricate no rId.
            for raw in raw_hyperlinks {
                let target = raw
                    .r_id
                    .as_deref()
                    .and_then(|rid| relation_target_for_id(&sheet_rels, rid))
                    .map(ToString::to_string)
                    .unwrap_or_default();
                let cell = ws.cells.entry(raw.reference.clone()).or_insert_with(|| {
                    Cell::parsed(
                        raw.reference.clone(),
                        Value::Empty,
                        StyleId::default(),
                        styles.clone(),
                    )
                });
                cell.write_hyperlink = Some(Hyperlink {
                    target,
                    location: raw.location,
                    tooltip: raw.tooltip,
                    display: raw.display,
                });
            }

            original_sheet_rels.push(sheet_rels);
            sheets.push(ws);
        }

        for (sheet_idx, value) in print_defined_names.print_areas {
            if let Some(sheet) = sheets.get_mut(sheet_idx) {
                sheet.print_area = normalize_print_setting(Some(value.as_str()));
            }
        }
        for (sheet_idx, rows, cols) in print_defined_names.print_titles {
            if let Some(sheet) = sheets.get_mut(sheet_idx) {
                sheet.print_title_rows = rows.and_then(|row| normalize_print_setting(Some(&row)));
                sheet.print_title_cols = cols.and_then(|col| normalize_print_setting(Some(&col)));
            }
        }

        // Attach sheet-local named ranges to their owning worksheet by index.
        // The promotion predicate already validated the index is in range.
        for (sheet_idx, defined_name) in sheet_local_defined_names {
            if let Some(sheet) = sheets.get_mut(sheet_idx) {
                sheet.defined_names_mut().insert_parsed(defined_name);
            }
        }

        Ok(Self {
            sheets,
            active_index,
            source_path: Some(source_path),
            original_styles: Some(styles),
            original_sheet_rels,
            extra_defined_names,
            defined_names,
        })
    }

    pub fn open_readonly<P: AsRef<Path>>(path: P) -> Result<ReadonlyWorkbook, OxlsxError> {
        crate::readonly::open_readonly(path)
    }

    pub fn worksheets(&self) -> &[Worksheet] {
        &self.sheets
    }

    /// Workbook-global named ranges (generic defined names).
    pub fn defined_names(&self) -> &crate::DefinedNames {
        &self.defined_names
    }

    /// Mutable access to the workbook-global named ranges.
    pub fn defined_names_mut(&mut self) -> &mut crate::DefinedNames {
        &mut self.defined_names
    }

    pub fn sheetnames(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.title()).collect()
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.title() == name)
    }

    pub fn active(&self) -> Option<&Worksheet> {
        self.sheets.get(self.active_index)
    }

    pub fn active_mut(&mut self) -> Option<&mut Worksheet> {
        self.sheets.get_mut(self.active_index)
    }

    pub fn sheet(&self, index: usize) -> Option<&Worksheet> {
        self.sheets.get(index)
    }

    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.sheets.get_mut(index)
    }

    pub fn create_sheet(&mut self, name: &str) -> &mut Worksheet {
        self.sheets.push(Worksheet::new(name));
        self.original_sheet_rels.push(Vec::new());
        self.sheets
            .last_mut()
            .expect("create_sheet always pushes one worksheet")
    }

    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<(), OxlsxError> {
        crate::writer::write_xlsx(self, path.as_ref(), self.source_path.as_deref())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

/// A raw `<hyperlink>` element captured during SAX parsing, before its external
/// `r:id` is resolved to a target via the sheet's relationships (which are
/// loaded after the sheet body in [`Workbook::open`]).
#[derive(Debug, Clone)]
struct RawHyperlink {
    reference: String,
    r_id: Option<String>,
    location: Option<String>,
    tooltip: Option<String>,
    display: Option<String>,
}

fn parse_sheet<R: BufRead>(
    reader: R,
    styles: Arc<StyleSheet>,
    shared_strings: &[String],
) -> Result<(Worksheet, Vec<RawHyperlink>), OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut decoder = CellDecoder::new(shared_strings);
    let mut cells: HashMap<String, Cell> = HashMap::new();
    let mut merged_cells = Vec::new();
    let mut column_widths = HashMap::new();
    let mut row_heights = HashMap::new();
    let mut freeze_panes = None;
    let mut in_sheet_view = false;
    let mut raw_hyperlinks = Vec::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                decoder.start_cell(&e)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"row" => {
                parse_row_height(&e, &mut row_heights)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"col" => {
                parse_column_widths(&e, &mut column_widths)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"mergeCell" => {
                if let Some(range_ref) = attr_value(&e, b"ref")? {
                    merged_cells.push(range_ref);
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = false;
            }
            Event::Start(e) if in_sheet_view && e.local_name().as_ref() == b"pane" => {
                if is_frozen_pane(&e)? {
                    freeze_panes =
                        normalize_freeze_panes(attr_value(&e, b"topLeftCell")?.as_deref());
                }
            }
            Event::Empty(e) if in_sheet_view && e.local_name().as_ref() == b"pane" => {
                if is_frozen_pane(&e)? {
                    freeze_panes =
                        normalize_freeze_panes(attr_value(&e, b"topLeftCell")?.as_deref());
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"v" => {
                decoder.start_value();
            }
            Event::Start(e) if e.local_name().as_ref() == b"f" => {
                decoder.start_formula();
            }
            Event::Start(e) if decoder.in_inline_str() && e.local_name().as_ref() == b"is" => {
                decoder.start_inline_str();
            }
            Event::Start(e) if decoder.in_inline_str() && e.local_name().as_ref() == b"t" => {
                decoder.start_inline_text();
            }
            Event::Empty(e) if decoder.in_inline_str() && e.local_name().as_ref() == b"t" => {
                decoder.end_inline_text();
            }
            Event::Empty(e) if e.local_name().as_ref() == b"row" => {
                parse_row_height(&e, &mut row_heights)?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"col" => {
                parse_column_widths(&e, &mut column_widths)?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"mergeCell" => {
                if let Some(range_ref) = attr_value(&e, b"ref")? {
                    merged_cells.push(range_ref);
                }
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"hyperlink" => {
                // `<hyperlink ref=".." r:id=".." location=".." tooltip=".."
                // display=".."/>`. The qualified `r:id` is keyed as `r:id`
                // (attr_value matches the full key, not the local name).
                if let Some(reference) = attr_value(&e, b"ref")? {
                    raw_hyperlinks.push(RawHyperlink {
                        reference,
                        r_id: attr_value(&e, b"r:id")?,
                        location: attr_value(&e, b"location")?,
                        tooltip: attr_value(&e, b"tooltip")?,
                        display: attr_value(&e, b"display")?,
                    });
                }
            }
            Event::Text(t) if decoder.in_value() => {
                decoder.push_value_text(&t.unescape()?);
            }
            Event::Text(t) if decoder.in_formula() => {
                decoder.push_formula_text(&t.unescape()?);
            }
            Event::Text(t) if decoder.in_inline_text() => {
                decoder.push_inline_text(&t.unescape()?);
            }
            Event::End(e) if e.local_name().as_ref() == b"v" => {
                decoder.end_value(&styles);
            }
            Event::End(e) if e.local_name().as_ref() == b"f" => {
                decoder.end_formula();
            }
            Event::End(e) if e.local_name().as_ref() == b"t" => {
                decoder.end_inline_text();
            }
            Event::End(e) if e.local_name().as_ref() == b"is" => {
                decoder.end_inline_str();
            }
            Event::End(e) if e.local_name().as_ref() == b"c" => {
                if let Some(cell) = decoder.finish_cell(styles.clone()) {
                    cells.insert(cell.reference().to_string(), cell);
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok((
        Worksheet {
            name: String::new(),
            cells,
            merged_cells,
            column_widths,
            row_heights,
            freeze_panes,
            print_area: None,
            print_title_rows: None,
            print_title_cols: None,
            defined_names: crate::DefinedNames::default(),
            tables: Vec::new(),
            table_retention: Vec::new(),
        },
        raw_hyperlinks,
    ))
}

fn is_frozen_pane(pane: &BytesStart<'_>) -> Result<bool, OxlsxError> {
    let state = attr_value(pane, b"state")?;
    Ok(matches!(
        state.as_deref(),
        Some("frozen") | Some("frozenSplit")
    ))
}

fn normalize_freeze_panes(value: Option<&str>) -> Option<String> {
    let trimmed = value?.trim();
    if trimmed.is_empty() {
        return None;
    }

    let upper = trimmed.to_ascii_uppercase();
    if upper == "A1" || parse_a1(&upper).is_none() {
        return None;
    }

    Some(upper)
}

fn normalize_print_setting(value: Option<&str>) -> Option<String> {
    let trimmed = value?.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Handle multi-range: split by comma first, strip prefix and make absolute per segment
    if trimmed.contains(',') {
        let segments: Vec<String> = trimmed
            .split(',')
            .filter_map(|segment| {
                let segment = segment.trim();
                if segment.is_empty() {
                    return None;
                }

                let bare = segment
                    .rsplit_once('!')
                    .map_or(segment, |(_, rhs)| rhs)
                    .trim();

                if bare.is_empty() {
                    return None;
                }

                Some(make_absolute_range(bare))
            })
            .collect();

        if segments.is_empty() {
            return None;
        }

        return Some(segments.join(","));
    }

    // Single range
    let bare_range = trimmed
        .rsplit_once('!')
        .map_or(trimmed, |(_, rhs)| rhs)
        .trim();

    if bare_range.is_empty() {
        return None;
    }

    let absolute_range = make_absolute_range(bare_range);
    Some(absolute_range)
}

#[derive(Debug, Clone)]
struct CellBuilder {
    reference: String,
    style_id: StyleId,
    cell_type: String,
    value: Value,
    formula_text: String,
}

#[derive(Debug, Clone)]
pub(crate) struct CellDecoder {
    shared_strings: Arc<Vec<String>>,
    value_text: String,
    inline_text: String,
    in_value: bool,
    in_formula: bool,
    in_inline_str: bool,
    in_inline_text: bool,
    current: Option<CellBuilder>,
}

impl CellDecoder {
    pub(crate) fn new(shared_strings: &[String]) -> Self {
        Self {
            shared_strings: Arc::new(shared_strings.to_vec()),
            value_text: String::new(),
            inline_text: String::new(),
            in_value: false,
            in_formula: false,
            in_inline_str: false,
            in_inline_text: false,
            current: None,
        }
    }

    pub(crate) fn new_arc(shared_strings: Arc<Vec<String>>) -> Self {
        Self {
            shared_strings,
            value_text: String::new(),
            inline_text: String::new(),
            in_value: false,
            in_formula: false,
            in_inline_str: false,
            in_inline_text: false,
            current: None,
        }
    }

    pub(crate) fn start_cell(&mut self, cell_start: &BytesStart<'_>) -> Result<(), OxlsxError> {
        let reference = attr_value(cell_start, b"r")?.unwrap_or_default();
        let style_id = attr_value(cell_start, b"s")?
            .and_then(|s| s.parse::<usize>().ok())
            .map(StyleId)
            .unwrap_or_default();
        let cell_type = attr_value(cell_start, b"t")?.unwrap_or_else(|| "n".to_string());

        self.current = Some(CellBuilder {
            reference,
            style_id,
            cell_type: cell_type.clone(),
            value: Value::Empty,
            formula_text: String::new(),
        });
        self.in_inline_str = cell_type == "inlineStr";
        self.in_inline_text = false;

        Ok(())
    }

    pub(crate) fn start_value(&mut self) {
        self.in_value = true;
        self.value_text.clear();
    }

    pub(crate) fn push_value_text(&mut self, text: &str) {
        if self.in_value {
            self.value_text.push_str(text);
        }
    }

    pub(crate) fn end_value(&mut self, styles: &StyleSheet) {
        self.in_value = false;

        if let Some(builder) = self.current.as_mut() {
            if matches!(builder.value, Value::Formula(_)) {
                return;
            }

            let mut value = match builder.cell_type.as_str() {
                "s" => self
                    .value_text
                    .parse::<usize>()
                    .ok()
                    .and_then(|i| self.shared_strings.get(i))
                    .map(|s| Value::String(s.to_string()))
                    .unwrap_or(Value::Empty),
                "b" => Value::Bool(self.value_text == "1"),
                _ => self
                    .value_text
                    .parse::<f64>()
                    .map(Value::Number)
                    .unwrap_or(Value::Empty),
            };

            if let Value::Number(n) = value
                && styles.is_date_format(builder.style_id)
            {
                value = excel_serial_to_date(n)
                    .map(Value::Date)
                    .unwrap_or(Value::Number(n));
            }

            builder.value = value;
        }
    }

    pub(crate) fn start_inline_str(&mut self) {
        if self.in_inline_str {
            self.inline_text.clear();
        }
    }

    pub(crate) fn start_formula(&mut self) {
        if let Some(builder) = self.current.as_mut() {
            builder.formula_text.clear();
            self.in_formula = true;
        }
    }

    pub(crate) fn push_formula_text(&mut self, text: &str) {
        if self.in_formula
            && let Some(builder) = self.current.as_mut()
        {
            builder.formula_text.push_str(text);
        }
    }

    pub(crate) fn end_formula(&mut self) {
        self.in_formula = false;
        if let Some(builder) = self.current.as_mut()
            && !builder.formula_text.is_empty()
        {
            builder.value = Value::Formula(mem::take(&mut builder.formula_text));
        }
    }

    pub(crate) fn start_inline_text(&mut self) {
        if self.in_inline_str {
            self.in_inline_text = true;
        }
    }

    pub(crate) fn push_inline_text(&mut self, text: &str) {
        if self.in_inline_text {
            self.inline_text.push_str(text);
        }
    }

    pub(crate) fn end_inline_text(&mut self) {
        self.in_inline_text = false;
    }

    pub(crate) fn end_inline_str(&mut self) {
        if let Some(builder) = self.current.as_mut() {
            builder.value = Value::String(mem::take(&mut self.inline_text));
        }
    }

    pub(crate) fn finish_cell(&mut self, styles: Arc<StyleSheet>) -> Option<Cell> {
        self.in_inline_str = false;
        self.in_inline_text = false;
        self.in_formula = false;

        let builder = self.current.take()?;
        if builder.reference.is_empty() {
            return None;
        }

        Some(Cell::parsed(
            builder.reference,
            builder.value,
            builder.style_id,
            styles,
        ))
    }

    pub(crate) fn in_value(&self) -> bool {
        self.in_value
    }

    pub(crate) fn in_inline_str(&self) -> bool {
        self.in_inline_str
    }

    pub(crate) fn in_formula(&self) -> bool {
        self.in_formula
    }

    pub(crate) fn in_inline_text(&self) -> bool {
        self.in_inline_text
    }
}

fn parse_column_widths(
    e: &BytesStart<'_>,
    column_widths: &mut HashMap<u32, f64>,
) -> Result<(), OxlsxError> {
    let width = attr_value(e, b"width")?.and_then(|v| v.parse::<f64>().ok());
    let min = attr_value(e, b"min")?.and_then(|v| v.parse::<u32>().ok());
    let max = attr_value(e, b"max")?.and_then(|v| v.parse::<u32>().ok());

    if let (Some(width), Some(min), Some(max)) = (width, min, max) {
        for col in min..=max {
            column_widths.insert(col, width);
        }
    }

    Ok(())
}

fn parse_row_height(
    e: &BytesStart<'_>,
    row_heights: &mut HashMap<u32, f64>,
) -> Result<(), OxlsxError> {
    let row = attr_value(e, b"r")?.and_then(|v| v.parse::<u32>().ok());
    let height = attr_value(e, b"ht")?.and_then(|v| v.parse::<f64>().ok());

    if let (Some(row), Some(height)) = (row, height) {
        row_heights.insert(row, height);
    }

    Ok(())
}

fn excel_serial_to_date(serial: f64) -> Option<NaiveDate> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let days = serial as i64;
    epoch.checked_add_signed(Duration::days(days))
}

#[cfg(test)]
mod tests {
    use super::{CellDecoder, Value, Workbook, Worksheet, parse_sheet};
    use crate::{Fill, Font, OxlsxError, StyleId, StyleSheet};
    use chrono::NaiveDate;
    use quick_xml::events::BytesStart;
    use std::collections::HashMap;

    #[test]
    fn table_capture_stores_raw_bytes() {
        let wb = Workbook::open("samples/tables_totals.xlsx").unwrap();
        let ws = wb.active().unwrap();

        assert_eq!(ws.tables().len(), ws.table_retention.len());
        assert!(!ws.tables().is_empty());

        let raw = ws.table_retention[0]
            .as_ref()
            .expect("a table read from disk retains its raw bytes");
        assert!(!raw.bytes.is_empty());
        assert!(!raw.dirty, "a freshly read table starts clean");
        // The retained bytes are the actual table part (not the parsed model).
        let text = String::from_utf8_lossy(&raw.bytes);
        assert!(text.contains("<table"));
        assert!(text.contains("totalsRowCount"));
    }

    #[test]
    fn dirty_tracking_getters_do_not_mark_dirty() {
        let mut wb = Workbook::open("samples/tables_totals.xlsx").unwrap();
        let ws = wb.active_mut().unwrap();

        // Read-only access leaves the table clean.
        let _ = ws.tables().iter().find(|t| t.display_name == "SalesTable");
        assert!(
            !ws.table_retention[0].as_ref().unwrap().dirty,
            "read-only access must not mark the table dirty"
        );

        // Mutable access flips the dirty flag.
        let _ = ws.table_mut("SalesTable").unwrap();
        assert!(
            ws.table_retention[0].as_ref().unwrap().dirty,
            "table_mut must mark the table dirty"
        );
    }
    use std::fs;
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::sync::Arc;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn tmp_path(name: &str) -> PathBuf {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock must be after unix epoch")
            .as_nanos();
        std::env::temp_dir().join(format!("{stamp}_{name}"))
    }

    const SAMPLE_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="s"><v>0</v></c>
      <c r="B1" s="2" t="s"><v>0</v></c>
      <c r="C1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const INLINE_STR_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>hello</t></is></c>
      <c r="B1" t="b"><v>1</v></c>
      <c r="C1" t="b"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const DATE_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B1" s="2"><v>45678</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const NON_DATE_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="C1" s="0"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    const SAMPLE_STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11.000000"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="11.000000"/><color theme="1"/><name val="Calibri"/></font>
    <font><sz val="11.000000"/><color indexed="4"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor indexed="2"/><bgColor indexed="2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor indexed="5"/><bgColor indexed="5"/></patternFill></fill>
  </fills>
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
  </numFmts>
  <cellXfs count="3">
    <xf fontId="0" fillId="0" numFmtId="0"/>
    <xf fontId="1" fillId="2" numFmtId="14"/>
    <xf fontId="2" fillId="3" numFmtId="164"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn parses_sheet_cells_style_ids_and_values() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let (ws, _) = parse_sheet(
            Cursor::new(SAMPLE_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        let a1 = ws.cell("A1").unwrap();
        assert_eq!(a1.style_id(), StyleId(1));
        assert_eq!(a1.value(), &Value::String("Hola".to_string()));

        let b1 = ws.cell("B1").unwrap();
        assert_eq!(b1.style_id(), StyleId(2));
        assert_eq!(b1.value(), &Value::String("Hola".to_string()));

        let c1 = ws.cell("C1").unwrap();
        assert_eq!(c1.style_id(), StyleId(0));
        assert_eq!(c1.value(), &Value::Number(42.0));

        assert!(ws.cell("D1").is_none());
    }

    #[test]
    fn style_accessors_use_stylesheet_via_style_id() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let (ws, _) = parse_sheet(
            Cursor::new(SAMPLE_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        let a1 = ws.cell("A1").unwrap();
        assert!(a1.font().unwrap().bold);
        assert_eq!(a1.fill().unwrap().pattern_type, "solid");
    }

    #[test]
    fn cell_value_accessor_returns_parsed_value() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let (ws, _) = parse_sheet(
            Cursor::new(SAMPLE_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        assert_eq!(
            ws.cell("A1").unwrap().value(),
            &Value::String("Hola".to_string())
        );
        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Number(42.0));
    }

    #[test]
    fn value_supports_string_and_bool_variants() {
        let text = Value::String("Hola".to_string());
        let boolean = Value::Bool(true);

        assert_eq!(text, Value::String("Hola".to_string()));
        assert_eq!(boolean, Value::Bool(true));
    }

    #[test]
    fn value_from_impls_convert_supported_types() {
        assert_eq!(Value::from("Hola"), Value::String("Hola".to_string()));
        assert_eq!(
            Value::from(String::from("Chau")),
            Value::String("Chau".to_string())
        );
        assert_eq!(Value::from(42.5_f64), Value::Number(42.5));
        assert_eq!(Value::from(7_i64), Value::Number(7.0));
        assert_eq!(Value::from(true), Value::Bool(true));
        assert_eq!(
            Value::from(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap()),
            Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
        );
    }

    #[test]
    fn parses_inline_string_and_boolean_cells() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let (ws, _) = parse_sheet(
            Cursor::new(INLINE_STR_SHEET_XML.as_bytes()),
            styles,
            &["Hola".to_string()],
        )
        .unwrap();

        assert_eq!(
            ws.cell("A1").unwrap().value(),
            &Value::String("hello".into())
        );
        assert_eq!(ws.cell("B1").unwrap().value(), &Value::Bool(true));
        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Bool(false));
    }

    #[test]
    fn cell_text_returns_only_string_values() {
        let styles = Arc::new(StyleSheet::default());

        let string_cell = super::Cell {
            reference: "A1".to_string(),
            value: Value::String("Hola".to_string()),
            style_id: StyleId(0),
            styles: styles.clone(),
            write_font: None,
            write_fill: None,
            comment: None,
            write_hyperlink: None,
        };
        assert_eq!(string_cell.text(), Some("Hola"));

        let number_cell = super::Cell {
            reference: "B1".to_string(),
            value: Value::Number(42.0),
            style_id: StyleId(0),
            styles: styles.clone(),
            write_font: None,
            write_fill: None,
            comment: None,
            write_hyperlink: None,
        };
        assert_eq!(number_cell.text(), None);

        let bool_cell = super::Cell {
            reference: "C1".to_string(),
            value: Value::Bool(true),
            style_id: StyleId(0),
            styles: styles.clone(),
            write_font: None,
            write_fill: None,
            comment: None,
            write_hyperlink: None,
        };
        assert_eq!(bool_cell.text(), None);

        let empty_cell = super::Cell {
            reference: "D1".to_string(),
            value: Value::Empty,
            style_id: StyleId(0),
            styles,
            write_font: None,
            write_fill: None,
            comment: None,
            write_hyperlink: None,
        };
        assert_eq!(empty_cell.text(), None);
    }

    #[test]
    fn workbook_open_missing_file_returns_io_error() {
        let err = Workbook::open("this-file-does-not-exist.xlsx").unwrap_err();
        assert!(matches!(err, OxlsxError::Io(_)));
    }

    #[test]
    fn workbook_open_invalid_zip_returns_invalid_format() {
        let mut path = std::env::temp_dir();
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        path.push(format!("oxlsx-invalid-{nanos}.xlsx"));

        fs::write(&path, b"not-a-zip").unwrap();
        let err = Workbook::open(&path).unwrap_err();
        fs::remove_file(&path).unwrap();

        assert!(matches!(err, OxlsxError::InvalidFormat(_)));
    }

    #[test]
    fn excel_serial_to_date_maps_known_serials() {
        assert_eq!(
            super::excel_serial_to_date(45678.0),
            NaiveDate::from_ymd_opt(2025, 1, 21)
        );
        assert_eq!(
            super::excel_serial_to_date(1.0),
            NaiveDate::from_ymd_opt(1899, 12, 31)
        );
    }

    #[test]
    fn parse_sheet_converts_numeric_date_by_style() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let (ws, _) = parse_sheet(Cursor::new(DATE_SHEET_XML.as_bytes()), styles, &[]).unwrap();

        assert_eq!(
            ws.cell("B1").unwrap().value(),
            &Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
        );
    }

    #[test]
    fn parse_sheet_keeps_non_date_numbers_unchanged() {
        let styles =
            Arc::new(StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap());
        let (ws, _) = parse_sheet(Cursor::new(NON_DATE_SHEET_XML.as_bytes()), styles, &[]).unwrap();

        assert_eq!(ws.cell("C1").unwrap().value(), &Value::Number(42.0));
    }

    #[test]
    fn worksheet_title_and_workbook_multi_sheet_accessors() {
        let wb = Workbook {
            sheets: vec![
                Worksheet {
                    name: "Empleados".to_string(),
                    cells: HashMap::new(),
                    merged_cells: Vec::new(),
                    column_widths: HashMap::new(),
                    row_heights: HashMap::new(),
                    freeze_panes: None,
                    print_area: None,
                    print_title_rows: None,
                    print_title_cols: None,
                    defined_names: crate::DefinedNames::default(),
                    tables: Vec::new(),
                    table_retention: Vec::new(),
                },
                Worksheet {
                    name: "Productos".to_string(),
                    cells: HashMap::new(),
                    merged_cells: Vec::new(),
                    column_widths: HashMap::new(),
                    row_heights: HashMap::new(),
                    freeze_panes: None,
                    print_area: None,
                    print_title_rows: None,
                    print_title_cols: None,
                    defined_names: crate::DefinedNames::default(),
                    tables: Vec::new(),
                    table_retention: Vec::new(),
                },
            ],
            active_index: 1,
            source_path: None,
            original_styles: None,
            original_sheet_rels: vec![Vec::new(), Vec::new()],
            extra_defined_names: Vec::new(),
            defined_names: crate::DefinedNames::default(),
        };

        assert_eq!(wb.sheet(0).unwrap().title(), "Empleados");
        assert_eq!(wb.sheet(1).unwrap().title(), "Productos");
        assert_eq!(wb.sheetnames(), vec!["Empleados", "Productos"]);
        assert!(wb.sheet_by_name("Empleados").is_some());
        assert!(wb.sheet_by_name("NoExiste").is_none());
        assert_eq!(wb.active().unwrap().title(), "Productos");
    }

    #[test]
    fn workbook_new_creates_default_sheet() {
        let wb = Workbook::new();

        assert_eq!(wb.sheet(0).unwrap().title(), "Sheet");
    }

    #[test]
    fn workbook_new_leaves_source_path_empty() {
        let wb = Workbook::new();

        assert!(wb.source_path.is_none());
    }

    #[test]
    fn workbook_open_tracks_source_path() {
        let wb = Workbook::open("samples/test.xlsx").unwrap();

        assert_eq!(
            wb.source_path.as_deref(),
            Some(std::path::Path::new("samples/test.xlsx"))
        );
    }

    #[test]
    fn workbook_create_sheet_appends_new_sheet() {
        let mut wb = Workbook::new();

        wb.create_sheet("Data");

        assert_eq!(wb.sheetnames(), vec!["Sheet", "Data"]);
    }

    #[test]
    fn workbook_active_mut_returns_some_for_default_workbook() {
        let mut wb = Workbook::new();

        assert!(wb.active_mut().is_some());
    }

    #[test]
    fn workbook_save_writes_xlsx_file() {
        let mut wb = Workbook::new();
        wb.active_mut().unwrap().set_cell("A1", "saved");
        let path = tmp_path("workbook_save_writes_xlsx_file.xlsx");

        wb.save(&path).unwrap();

        assert!(path.exists());
        fs::remove_file(path).ok();
    }

    #[test]
    fn worksheet_set_cell_sets_string_number_and_bool_values() {
        let mut ws = Worksheet::new("Data");

        ws.set_cell("A1", "Hello");
        assert_eq!(
            ws.cell("A1").unwrap().value(),
            &Value::String("Hello".to_string())
        );

        ws.set_cell("A1", 42.0_f64);
        assert_eq!(ws.cell("A1").unwrap().value(), &Value::Number(42.0));

        ws.set_cell("A1", true);
        assert_eq!(ws.cell("A1").unwrap().value(), &Value::Bool(true));
    }

    #[test]
    fn worksheet_set_cell_font_stores_write_font_and_creates_missing_cell() {
        let mut ws = Worksheet::new("Data");

        ws.set_cell_font(
            "A1",
            Font {
                bold: true,
                ..Default::default()
            },
        );

        let cell = ws.cell("A1").unwrap();
        assert!(cell.write_font.is_some());
    }

    #[test]
    fn worksheet_set_cell_fill_stores_write_fill() {
        let mut ws = Worksheet::new("Data");

        ws.set_cell_fill(
            "A1",
            Fill {
                pattern_type: "solid".to_string(),
                ..Default::default()
            },
        );

        let cell = ws.cell("A1").unwrap();
        assert!(cell.write_fill.is_some());
    }

    #[test]
    fn worksheet_bounds_scan_from_cells() {
        let mut ws = Worksheet::new("Data");
        ws.set_cell("C3", "three");
        ws.set_cell("A2", "two");
        ws.set_cell("B5", "five");

        assert_eq!(ws.min_row(), 2);
        assert_eq!(ws.max_row(), 5);
        assert_eq!(ws.min_column(), 1);
        assert_eq!(ws.max_column(), 3);
    }

    #[test]
    fn worksheet_iter_rows_rectangular() {
        let mut ws = Worksheet::new("Data");
        ws.set_cell("A1", "left");
        ws.set_cell("C1", "right");
        ws.set_cell("B2", 42.0_f64);

        let rows = ws.iter_rows(Some(1), Some(2), Some(1), Some(3));
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].len(), 3);
        assert_eq!(rows[1].len(), 3);

        assert_eq!(rows[0][0].value(), &Value::String("left".to_string()));
        assert_eq!(rows[0][1].value(), &Value::Empty);
        assert_eq!(rows[0][2].value(), &Value::String("right".to_string()));
        assert_eq!(rows[1][0].value(), &Value::Empty);
        assert_eq!(rows[1][1].value(), &Value::Number(42.0));
        assert_eq!(rows[1][2].value(), &Value::Empty);
    }

    #[test]
    fn worksheet_iter_cols_rectangular() {
        let mut ws = Worksheet::new("Data");
        ws.set_cell("A1", "top");
        ws.set_cell("A3", "bottom");
        ws.set_cell("B2", 7_i64);

        let cols = ws.iter_cols(Some(1), Some(3), Some(1), Some(2));
        assert_eq!(cols.len(), 2);
        assert_eq!(cols[0].len(), 3);
        assert_eq!(cols[1].len(), 3);

        assert_eq!(cols[0][0].value(), &Value::String("top".to_string()));
        assert_eq!(cols[0][1].value(), &Value::Empty);
        assert_eq!(cols[0][2].value(), &Value::String("bottom".to_string()));
        assert_eq!(cols[1][0].value(), &Value::Empty);
        assert_eq!(cols[1][1].value(), &Value::Number(7.0));
        assert_eq!(cols[1][2].value(), &Value::Empty);
    }

    #[test]
    fn worksheet_values_property_returns_values_only() {
        let mut ws = Worksheet::new("Data");
        ws.set_cell("A1", "name");
        ws.set_cell("B1", 42_i64);
        ws.set_cell("A2", true);

        let values = ws.values();
        assert_eq!(values.len(), 2);
        assert_eq!(
            values[0],
            vec![Value::String("name".to_string()), Value::Number(42.0)]
        );
        assert_eq!(values[1], vec![Value::Bool(true), Value::Empty]);
    }

    #[test]
    fn workbook_open_loads_sheet_refs_from_workbook_metadata() {
        let wb = Workbook::open("samples/multi_sheet.xlsx").unwrap();

        assert_eq!(wb.worksheets().len(), 2);
        assert_eq!(wb.sheetnames(), vec!["Empleados", "Productos"]);
        assert_eq!(wb.active().unwrap().title(), "Empleados");
    }

    #[test]
    fn cell_decoder_decodes_shared_string_cells() {
        let shared_strings = ["Hola".to_string()];
        let styles = Arc::new(StyleSheet::default());
        let mut decoder = CellDecoder::new(&shared_strings);

        let mut cell_start = BytesStart::new("c");
        cell_start.push_attribute(("r", "A1"));
        cell_start.push_attribute(("s", "0"));
        cell_start.push_attribute(("t", "s"));

        decoder.start_cell(&cell_start).unwrap();
        decoder.start_value();
        decoder.push_value_text("0");
        decoder.end_value(&styles);

        let cell = decoder.finish_cell(styles).expect("cell should be decoded");
        assert_eq!(cell.reference(), "A1");
        assert_eq!(cell.style_id(), StyleId(0));
        assert_eq!(cell.value(), &Value::String("Hola".to_string()));
    }
}
