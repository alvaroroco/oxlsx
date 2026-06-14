use crate::coords::{make_absolute_line_range, make_absolute_range};
use crate::xml::attr_value;
use crate::{OxlsxError, Worksheet};
use quick_xml::Reader;
use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use std::io::BufRead;

pub(crate) const PRINT_AREA_NAME: &str = "_xlnm.Print_Area";
pub(crate) const PRINT_TITLES_NAME: &str = "_xlnm.Print_Titles";

const RESERVED_NAME_PREFIX: &str = "_xlnm.";

/// A workbook-global named range (generic `<definedName>` entry).
///
/// Mirrors openpyxl's `DefinedName`: `value` is the cell reference, range, or
/// formula text (openpyxl's `attr_text`). Only the conservatively modeled
/// attributes are represented; entries carrying any other attribute are kept
/// verbatim as raw passthrough and never surfaced here.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    pub comment: Option<String>,
    pub hidden: bool,
}

impl DefinedName {
    /// Creates a named range from a name and its reference/formula value.
    pub fn new(name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            value: value.into(),
            comment: None,
            hidden: false,
        }
    }
}

/// Ordered collection of workbook-global named ranges with dict-like access.
///
/// Names are unique; adding a name that already exists replaces it
/// (openpyxl semantics). Reserved `_xlnm.*` names are rejected by [`add`](Self::add)
/// because they are owned by the print-settings API.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DefinedNames {
    // A Vec preserves deterministic (insertion/read) order for serialization.
    // Named ranges are typically few, so linear lookup is adequate.
    names: Vec<DefinedName>,
}

impl DefinedNames {
    /// Inserts a named range, replacing any existing entry with the same name.
    ///
    /// # Errors
    /// Returns [`OxlsxError::InvalidFormat`] when `defined_name.name` starts with
    /// the reserved `_xlnm.` prefix.
    pub fn add(&mut self, defined_name: DefinedName) -> Result<(), OxlsxError> {
        if defined_name.name.starts_with(RESERVED_NAME_PREFIX) {
            return Err(OxlsxError::InvalidFormat(
                "'_xlnm.*' names are reserved for print settings; use ws.print_area / ws.print_title_rows / ws.print_title_cols instead".to_string(),
            ));
        }

        if let Some(existing) = self
            .names
            .iter_mut()
            .find(|dn| dn.name == defined_name.name)
        {
            *existing = defined_name;
        } else {
            self.names.push(defined_name);
        }
        Ok(())
    }

    /// Returns the named range matching `name`, if present.
    pub fn get(&self, name: &str) -> Option<&DefinedName> {
        self.names.iter().find(|dn| dn.name == name)
    }

    /// Removes and returns the named range matching `name`, if present.
    pub fn remove(&mut self, name: &str) -> Option<DefinedName> {
        let position = self.names.iter().position(|dn| dn.name == name)?;
        Some(self.names.remove(position))
    }

    /// Reports whether a named range with `name` exists.
    pub fn contains(&self, name: &str) -> bool {
        self.names.iter().any(|dn| dn.name == name)
    }

    /// Iterates over the names in deterministic order.
    pub fn names(&self) -> impl Iterator<Item = &str> {
        self.names.iter().map(|dn| dn.name.as_str())
    }

    /// Iterates over the named ranges in deterministic order.
    pub fn iter(&self) -> impl Iterator<Item = &DefinedName> {
        self.names.iter()
    }

    /// Returns the number of named ranges.
    pub fn len(&self) -> usize {
        self.names.len()
    }

    /// Reports whether the collection has no named ranges.
    pub fn is_empty(&self) -> bool {
        self.names.is_empty()
    }

    /// Inserts an already-validated entry parsed from the workbook, replacing
    /// any duplicate. Used on the read path where the reserved-name filter has
    /// already been applied by the promotion predicate.
    pub(crate) fn insert_parsed(&mut self, defined_name: DefinedName) {
        if let Some(existing) = self
            .names
            .iter_mut()
            .find(|dn| dn.name == defined_name.name)
        {
            *existing = defined_name;
        } else {
            self.names.push(defined_name);
        }
    }
}

impl<'a> IntoIterator for &'a DefinedNames {
    type Item = &'a DefinedName;
    type IntoIter = std::slice::Iter<'a, DefinedName>;

    fn into_iter(self) -> Self::IntoIter {
        self.names.iter()
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PrintDefinedNames {
    pub(crate) print_areas: Vec<(usize, String)>,
    pub(crate) print_titles: Vec<(usize, Option<String>, Option<String>)>,
}

/// The four buckets produced when parsing `<definedNames>`: reserved print
/// settings, promotable sheet-local names (keyed by sheet index), promotable
/// workbook-global names, and raw verbatim passthrough.
type ParsedDefinedNames = (
    PrintDefinedNames,
    Vec<(usize, DefinedName)>,
    DefinedNames,
    Vec<String>,
);

pub(crate) fn parse_workbook_defined_names<R: BufRead>(
    reader: R,
    sheet_count: usize,
) -> Result<ParsedDefinedNames, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(false);

    let mut parsed = PrintDefinedNames::default();
    let mut sheet_local_names: Vec<(usize, DefinedName)> = Vec::new();
    let mut defined_names = DefinedNames::default();
    let mut extras = Vec::new();
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                let name = attr_value(&e, b"name")?;
                let local_sheet_id =
                    attr_value(&e, b"localSheetId")?.and_then(|v| v.parse::<usize>().ok());

                match (name.as_deref(), local_sheet_id) {
                    (Some(PRINT_AREA_NAME), Some(sheet_idx)) => {
                        let text = read_defined_name_text(&mut xml)?;
                        let bare_range = strip_sheet_prefix_multi(text.trim());
                        if !bare_range.is_empty() {
                            let absolute_range = make_absolute_area_range(&bare_range);
                            parsed.print_areas.push((sheet_idx, absolute_range));
                        }
                    }
                    (Some(PRINT_TITLES_NAME), Some(sheet_idx)) => {
                        let text = read_defined_name_text(&mut xml)?;
                        let (rows, cols) = parse_print_titles(text.trim());
                        if rows.is_some() || cols.is_some() {
                            parsed.print_titles.push((sheet_idx, rows, cols));
                        }
                    }
                    _ => match classify_defined_name_scope(&e, sheet_count)? {
                        Scope::RawExtra => {
                            extras.push(capture_defined_name_raw(&mut xml, e.into_owned())?);
                        }
                        scope => {
                            let comment = attr_value(&e, b"comment")?;
                            let hidden =
                                attr_value(&e, b"hidden")?.is_some_and(|v| parse_xml_bool(&v));
                            let name = name.expect("promotable predicate implies a name attribute");
                            let value = read_defined_name_text(&mut xml)?;
                            let defined_name = DefinedName {
                                name,
                                value,
                                comment,
                                hidden,
                            };
                            route_promoted(
                                scope,
                                defined_name,
                                &mut sheet_local_names,
                                &mut defined_names,
                            );
                        }
                    },
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"definedName" => {
                let name = attr_value(&e, b"name")?;
                let local_sheet_id =
                    attr_value(&e, b"localSheetId")?.and_then(|v| v.parse::<usize>().ok());

                if matches!(
                    (name.as_deref(), local_sheet_id),
                    (Some(PRINT_AREA_NAME), Some(_)) | (Some(PRINT_TITLES_NAME), Some(_))
                ) {
                    // Reserved self-closing entry carries no range; print code owns it.
                } else {
                    match classify_defined_name_scope(&e, sheet_count)? {
                        Scope::RawExtra => {
                            extras.push(event_to_xml(Event::Empty(e.into_owned()))?);
                        }
                        scope => {
                            let comment = attr_value(&e, b"comment")?;
                            let hidden =
                                attr_value(&e, b"hidden")?.is_some_and(|v| parse_xml_bool(&v));
                            let name = name.expect("promotable predicate implies a name attribute");
                            let defined_name = DefinedName {
                                name,
                                value: String::new(),
                                comment,
                                hidden,
                            };
                            route_promoted(
                                scope,
                                defined_name,
                                &mut sheet_local_names,
                                &mut defined_names,
                            );
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((parsed, sheet_local_names, defined_names, extras))
}

/// Routes a promoted name into its scoped collection. Reserved names and raw
/// passthrough never reach here (handled by the caller).
fn route_promoted(
    scope: Scope,
    defined_name: DefinedName,
    sheet_local_names: &mut Vec<(usize, DefinedName)>,
    defined_names: &mut DefinedNames,
) {
    match scope {
        Scope::SheetLocal(sheet_idx) => sheet_local_names.push((sheet_idx, defined_name)),
        Scope::Global => defined_names.insert_parsed(defined_name),
        // RawExtra is filtered out before this call.
        Scope::RawExtra => {}
    }
}

/// Promotion scope for a `<definedName>`. Reserved `_xlnm.*` print names are
/// intercepted before classification and never produce a [`Scope`].
enum Scope {
    /// Promotable workbook-global name (no `localSheetId`).
    Global,
    /// Promotable sheet-local name bound to the worksheet at this index.
    SheetLocal(usize),
    /// Not modeled — kept verbatim as raw passthrough (zero loss).
    RawExtra,
}

/// Classifies a `<definedName>` (conservative/lossless). A name is promoted only
/// when it is not reserved and carries no attribute outside
/// `{name, comment, hidden, localSheetId}`. A present `localSheetId` MUST parse
/// and map to an existing sheet index; otherwise the entry stays raw passthrough
/// so malformed or out-of-range references round-trip verbatim.
fn classify_defined_name_scope(
    e: &BytesStart<'_>,
    sheet_count: usize,
) -> Result<Scope, OxlsxError> {
    let mut has_name = false;
    let mut has_local_sheet_id = false;
    let mut local_sheet_id: Option<usize> = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        match attr.key.as_ref() {
            b"name" => {
                let value = attr.unescape_value()?;
                if value.starts_with(RESERVED_NAME_PREFIX) {
                    return Ok(Scope::RawExtra);
                }
                has_name = true;
            }
            b"localSheetId" => {
                has_local_sheet_id = true;
                local_sheet_id = attr.unescape_value()?.parse::<usize>().ok();
            }
            b"comment" | b"hidden" => {}
            // Any other attribute (function, customMenu, ...) blocks promotion.
            _ => return Ok(Scope::RawExtra),
        }
    }

    if !has_name {
        return Ok(Scope::RawExtra);
    }

    if has_local_sheet_id {
        // Present but unparsable or out-of-range => keep verbatim (lossless).
        match local_sheet_id {
            Some(sheet_idx) if sheet_idx < sheet_count => Ok(Scope::SheetLocal(sheet_idx)),
            _ => Ok(Scope::RawExtra),
        }
    } else {
        Ok(Scope::Global)
    }
}

fn parse_xml_bool(value: &str) -> bool {
    matches!(value, "1" | "true")
}

pub(crate) fn write_defined_names_block(
    xml: &mut Writer<Vec<u8>>,
    sheets: &[Worksheet],
    defined_names: &DefinedNames,
    extras: &[String],
) -> Result<(), OxlsxError> {
    let has_reserved = sheets.iter().any(|sheet| {
        sheet.print_area_bare().is_some()
            || sheet.print_title_rows().is_some()
            || sheet.print_title_cols().is_some()
    });

    if !has_reserved
        && defined_names.is_empty()
        && sheets.iter().all(|sheet| sheet.defined_names().is_empty())
        && extras.is_empty()
    {
        return Ok(());
    }

    xml.write_event(Event::Start(BytesStart::new("definedNames")))?;

    for (sheet_idx, sheet) in sheets.iter().enumerate() {
        if let Some(value) = compose_print_area(sheet.title(), sheet.print_area_bare()) {
            write_reserved_defined_name(xml, PRINT_AREA_NAME, sheet_idx, value.as_str())?;
        }

        if let Some(value) = compose_print_titles(
            sheet.title(),
            sheet.print_title_rows(),
            sheet.print_title_cols(),
        ) {
            write_reserved_defined_name(xml, PRINT_TITLES_NAME, sheet_idx, value.as_str())?;
        }
    }

    for defined_name in defined_names.iter() {
        write_modeled_defined_name(xml, defined_name, None)?;
    }

    // Sheet-local names: localSheetId is COMPUTED from the worksheet's live
    // position, never stored — so it stays correct across the workbook lifecycle.
    for (sheet_idx, sheet) in sheets.iter().enumerate() {
        for defined_name in sheet.defined_names().iter() {
            write_modeled_defined_name(xml, defined_name, Some(sheet_idx))?;
        }
    }

    for extra in extras {
        xml.get_mut().extend_from_slice(extra.as_bytes());
    }

    xml.write_event(Event::End(BytesEnd::new("definedNames")))?;
    Ok(())
}

/// Emits a modeled `<definedName>`. When `local_sheet_id` is `Some`, the
/// computed sheet index is written, producing a sheet-local entry; otherwise a
/// workbook-global entry is written.
fn write_modeled_defined_name(
    xml: &mut Writer<Vec<u8>>,
    defined_name: &DefinedName,
    local_sheet_id: Option<usize>,
) -> Result<(), OxlsxError> {
    let mut start = BytesStart::new("definedName");
    start.push_attribute(("name", defined_name.name.as_str()));
    let local_sheet_id = local_sheet_id.map(|idx| idx.to_string());
    if let Some(local_sheet_id) = &local_sheet_id {
        start.push_attribute(("localSheetId", local_sheet_id.as_str()));
    }
    if defined_name.hidden {
        start.push_attribute(("hidden", "1"));
    }
    if let Some(comment) = &defined_name.comment {
        start.push_attribute(("comment", comment.as_str()));
    }

    xml.write_event(Event::Start(start))?;
    xml.write_event(Event::Text(BytesText::new(&defined_name.value)))?;
    xml.write_event(Event::End(BytesEnd::new("definedName")))?;
    Ok(())
}

fn compose_print_titles(title: &str, rows: Option<&str>, cols: Option<&str>) -> Option<String> {
    let mut parts = Vec::with_capacity(2);
    let quoted_title = quote_sheet_title(title);
    if let Some(rows) = rows {
        parts.push(format!("{quoted_title}!{rows}"));
    }
    if let Some(cols) = cols {
        parts.push(format!("{quoted_title}!{cols}"));
    }

    if parts.is_empty() {
        None
    } else {
        Some(parts.join(","))
    }
}

fn compose_print_area(title: &str, area: Option<&str>) -> Option<String> {
    let area = area?;
    let quoted_title = quote_sheet_title(title);

    // Handle multi-range: split by comma, qualify each segment
    if area.contains(',') {
        let qualified = split_top_level_csv(area)
            .iter()
            .map(|segment| format!("{}!{}", quoted_title, segment.trim()))
            .collect::<Vec<_>>()
            .join(",");
        return Some(qualified);
    }

    Some(format!("{quoted_title}!{area}"))
}

#[cfg(test)]
mod compose_tests {
    use super::{compose_print_area, compose_print_titles};

    #[test]
    fn compose_print_titles_qualifies_with_sheet() {
        let result = compose_print_titles("Sheet", Some("$1:$3"), Some("$A:$C"));
        assert_eq!(result, Some("'Sheet'!$1:$3,'Sheet'!$A:$C".to_string()));
    }

    #[test]
    fn compose_print_area_qualifies_with_sheet() {
        let result = compose_print_area("Sheet", Some("$A$1:$D$4"));
        assert_eq!(result, Some("'Sheet'!$A$1:$D$4".to_string()));
    }
}

fn quote_sheet_title(title: &str) -> String {
    let escaped = title.replace('\'', "''");
    format!("'{escaped}'")
}

fn write_reserved_defined_name(
    xml: &mut Writer<Vec<u8>>,
    name: &str,
    sheet_idx: usize,
    value: &str,
) -> Result<(), OxlsxError> {
    let mut start = BytesStart::new("definedName");
    let local_sheet_id = sheet_idx.to_string();
    start.push_attribute(("name", name));
    start.push_attribute(("localSheetId", local_sheet_id.as_str()));

    xml.write_event(Event::Start(start))?;
    xml.write_event(Event::Text(BytesText::new(value)))?;
    xml.write_event(Event::End(BytesEnd::new("definedName")))?;
    Ok(())
}

fn capture_defined_name_raw<R: BufRead>(
    xml: &mut Reader<R>,
    start: BytesStart<'static>,
) -> Result<String, OxlsxError> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(start))?;

    let mut depth = 1_usize;
    let mut buf = Vec::new();
    while depth > 0 {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                depth += 1;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::End(e) if e.local_name().as_ref() == b"definedName" => {
                depth -= 1;
                writer.write_event(Event::End(e.into_owned()))?;
            }
            Event::Eof => {
                return Err(OxlsxError::InvalidFormat(
                    "unterminated <definedName> element".to_string(),
                ));
            }
            event => {
                writer.write_event(event.into_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner()).map_err(|err| OxlsxError::InvalidFormat(err.to_string()))
}

fn event_to_xml(event: Event<'static>) -> Result<String, OxlsxError> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event)?;
    String::from_utf8(writer.into_inner()).map_err(|err| OxlsxError::InvalidFormat(err.to_string()))
}

fn read_defined_name_text<R: BufRead>(xml: &mut Reader<R>) -> Result<String, OxlsxError> {
    let mut value = String::new();
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Text(text) => value.push_str(text.unescape()?.as_ref()),
            Event::CData(text) => {
                let cdata = std::str::from_utf8(text.as_ref())
                    .map_err(|err| OxlsxError::InvalidFormat(err.to_string()))?;
                value.push_str(cdata);
            }
            Event::End(e) if e.local_name().as_ref() == b"definedName" => break,
            Event::Eof => {
                return Err(OxlsxError::InvalidFormat(
                    "unterminated <definedName> text".to_string(),
                ));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(value)
}

fn parse_print_titles(value: &str) -> (Option<String>, Option<String>) {
    let mut rows = None;
    let mut cols = None;

    for segment in split_top_level_csv(value) {
        let segment = segment.trim();
        if segment.is_empty() {
            continue;
        }

        let range = segment
            .rsplit_once('!')
            .map_or(segment, |(_, rhs)| rhs)
            .trim();
        if is_row_segment(range) {
            rows = Some(make_absolute_line_range(range));
        } else if is_col_segment(range) {
            cols = Some(make_absolute_line_range(range));
        }
    }

    (rows, cols)
}

fn make_absolute_area_range(range: &str) -> String {
    // Handle multi-range (comma-separated): "A1:B2,D1:E2"
    if range.contains(',') {
        return split_top_level_csv(range)
            .iter()
            .map(|segment| make_absolute_range(segment.trim()))
            .collect::<Vec<_>>()
            .join(",");
    }

    make_absolute_range(range)
}

fn strip_sheet_prefix(value: &str) -> &str {
    value.rsplit_once('!').map_or(value, |(_, rhs)| rhs).trim()
}

fn strip_sheet_prefix_multi(value: &str) -> String {
    // Handle multi-range: "Sheet!$A$1:$B$2,Sheet!$D$1:$E$2"
    // Split by comma first, strip prefix from each segment
    if value.contains(',') {
        return split_top_level_csv(value)
            .iter()
            .map(|segment| strip_sheet_prefix(segment.trim()))
            .collect::<Vec<_>>()
            .join(",");
    }

    strip_sheet_prefix(value).to_string()
}

fn split_top_level_csv(value: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut chars = value.chars().peekable();
    let mut in_quotes = false;

    while let Some(ch) = chars.next() {
        if ch == '\'' {
            if in_quotes && chars.peek().is_some_and(|next| *next == '\'') {
                current.push(ch);
                current.push(chars.next().expect("peeked a quote"));
                continue;
            }
            in_quotes = !in_quotes;
            current.push(ch);
            continue;
        }

        if ch == ',' && !in_quotes {
            parts.push(current.trim().to_string());
            current.clear();
            continue;
        }

        current.push(ch);
    }

    if !current.trim().is_empty() {
        parts.push(current.trim().to_string());
    }

    parts
}

fn is_row_segment(range: &str) -> bool {
    let Some((start, end)) = range.split_once(':') else {
        return false;
    };
    let start = start.strip_prefix('$').unwrap_or(start);
    let end = end.strip_prefix('$').unwrap_or(end);
    !start.is_empty()
        && !end.is_empty()
        && start.bytes().all(|byte| byte.is_ascii_digit())
        && end.bytes().all(|byte| byte.is_ascii_digit())
}

fn is_col_segment(range: &str) -> bool {
    let Some((start, end)) = range.split_once(':') else {
        return false;
    };
    let start = start.strip_prefix('$').unwrap_or(start);
    let end = end.strip_prefix('$').unwrap_or(end);
    !start.is_empty()
        && !end.is_empty()
        && start.bytes().all(|byte| byte.is_ascii_alphabetic())
        && end.bytes().all(|byte| byte.is_ascii_alphabetic())
}

#[cfg(test)]
mod collection_tests {
    use super::{DefinedName, DefinedNames};

    #[test]
    fn add_then_get_returns_value() {
        let mut names = DefinedNames::default();
        names
            .add(DefinedName::new("Sales", "'Sheet'!$A$1"))
            .unwrap();
        assert_eq!(names.get("Sales").unwrap().value, "'Sheet'!$A$1");
        assert!(names.contains("Sales"));
    }

    #[test]
    fn add_duplicate_name_replaces_value() {
        let mut names = DefinedNames::default();
        names
            .add(DefinedName::new("Sales", "'Sheet'!$A$1"))
            .unwrap();
        names
            .add(DefinedName::new("Sales", "'Sheet'!$B$2"))
            .unwrap();
        assert_eq!(names.get("Sales").unwrap().value, "'Sheet'!$B$2");
        assert_eq!(names.names().count(), 1);
    }

    #[test]
    fn add_rejects_reserved_xlnm_names() {
        let mut names = DefinedNames::default();
        let err = names.add(DefinedName::new("_xlnm.Print_Area", "'Sheet'!$A$1"));
        assert!(err.is_err());
        assert!(!names.contains("_xlnm.Print_Area"));
    }

    #[test]
    fn remove_returns_entry_and_drops_it() {
        let mut names = DefinedNames::default();
        names
            .add(DefinedName::new("Region", "'Sheet'!$B$1"))
            .unwrap();
        assert_eq!(
            names.remove("Region").map(|dn| dn.value),
            Some("'Sheet'!$B$1".to_string())
        );
        assert!(!names.contains("Region"));
    }

    #[test]
    fn remove_missing_name_returns_none() {
        let mut names = DefinedNames::default();
        names
            .add(DefinedName::new("Region", "'Sheet'!$B$1"))
            .unwrap();
        assert!(names.remove("Ghost").is_none());
        assert!(names.contains("Region"));
    }
}

#[cfg(test)]
mod tests {
    use super::{
        PRINT_AREA_NAME, PRINT_TITLES_NAME, parse_print_titles, parse_workbook_defined_names,
        split_top_level_csv,
    };
    use std::io::Cursor;

    #[test]
    fn parse_workbook_defined_names_extracts_reserved_and_extras() {
        let workbook_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="{PRINT_AREA_NAME}" localSheetId="0">'Sheet'!$A$1:$C$10</definedName>
    <definedName name="{PRINT_TITLES_NAME}" localSheetId="0">'Sheet'!$1:$3,'Sheet'!$A:$C</definedName>
    <definedName name="KeepMe">Sheet!$D$1</definedName>
  </definedNames>
</workbook>"#
        );

        let (parsed, sheet_local, defined_names, extras) =
            parse_workbook_defined_names(Cursor::new(workbook_xml), 1).unwrap();

        assert_eq!(parsed.print_areas, vec![(0, "$A$1:$C$10".to_string())]);
        assert_eq!(
            parsed.print_titles,
            vec![(0, Some("$1:$3".to_string()), Some("$A:$C".to_string()))]
        );
        // "KeepMe" is a simple global name => promoted into the typed collection.
        assert_eq!(
            defined_names.get("KeepMe").map(|dn| dn.value.as_str()),
            Some("Sheet!$D$1")
        );
        assert!(sheet_local.is_empty());
        assert!(extras.is_empty());
    }

    #[test]
    fn parse_workbook_defined_names_splits_attr_global_and_sheet_local() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="Macro" function="1">SUM(1,2)</definedName>
    <definedName name="Local" localSheetId="0">Sheet!$D$1</definedName>
  </definedNames>
</workbook>"#;

        let (_, sheet_local, defined_names, extras) =
            parse_workbook_defined_names(Cursor::new(workbook_xml), 1).unwrap();

        // The function-attr global is unmodeled => stays raw verbatim.
        assert!(defined_names.get("Macro").is_none());
        assert_eq!(extras.len(), 1);
        assert!(extras[0].contains("function=\"1\""));

        // The simple sheet-local name is promoted onto sheet index 0.
        assert_eq!(sheet_local.len(), 1);
        assert_eq!(sheet_local[0].0, 0);
        assert_eq!(sheet_local[0].1.name, "Local");
        assert_eq!(sheet_local[0].1.value, "Sheet!$D$1");
    }

    #[test]
    fn parse_workbook_defined_names_keeps_out_of_range_sheet_local_raw() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="Ghost" localSheetId="9">Sheet!$D$1</definedName>
  </definedNames>
</workbook>"#;

        // Only one sheet exists, so localSheetId="9" is out of range => raw passthrough.
        let (_, sheet_local, defined_names, extras) =
            parse_workbook_defined_names(Cursor::new(workbook_xml), 1).unwrap();

        assert!(sheet_local.is_empty());
        assert!(defined_names.get("Ghost").is_none());
        assert_eq!(extras.len(), 1);
        assert!(extras[0].contains("localSheetId=\"9\""));
    }

    #[test]
    fn split_top_level_csv_ignores_comma_inside_sheet_name() {
        let parts = split_top_level_csv("'Sales,West'!$A:$C,'Sales,West'!$1:$2");
        assert_eq!(parts.len(), 2);
        assert_eq!(parts[0], "'Sales,West'!$A:$C");
        assert_eq!(parts[1], "'Sales,West'!$1:$2");
    }

    #[test]
    fn parse_print_titles_splits_rows_and_cols() {
        let (rows, cols) = parse_print_titles("'Sheet'!$A:$C,'Sheet'!$1:$3");
        assert_eq!(rows, Some("$1:$3".to_string()));
        assert_eq!(cols, Some("$A:$C".to_string()));
    }

    #[test]
    fn parse_print_titles_accepts_bare_ranges() {
        let (rows, cols) = parse_print_titles("'Sheet'!1:3,'Sheet'!A:C");
        assert_eq!(rows, Some("$1:$3".to_string()));
        assert_eq!(cols, Some("$A:$C".to_string()));
    }
}
