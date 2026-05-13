use crate::OxlsxError;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use std::io::BufRead;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetRef {
    pub name: String,
    pub zip_path: String,
}

pub fn parse_sheet_refs<R1: BufRead, R2: BufRead>(
    workbook_xml: R1,
    rels_xml: R2,
) -> Result<(Vec<SheetRef>, usize), OxlsxError> {
    let (sheet_entries, active_index) = parse_workbook(workbook_xml)?;
    let rels = parse_rels(rels_xml)?;

    let mut sheets = Vec::with_capacity(sheet_entries.len());
    for (name, rel_id) in sheet_entries {
        let zip_path = rels
            .get(&rel_id)
            .cloned()
            .ok_or_else(|| OxlsxError::MissingPart(format!("workbook relationship {rel_id}")))?;
        sheets.push(SheetRef { name, zip_path });
    }

    Ok((sheets, active_index))
}

fn parse_workbook<R: BufRead>(reader: R) -> Result<(Vec<(String, String)>, usize), OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheet_entries = Vec::new();
    let mut active_index = 0usize;

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => match e.local_name().as_ref() {
                b"sheet" => {
                    let rel_id = attr_value(&e, b"r:id")?.or(attr_value(&e, b"id")?);
                    if let (Some(name), Some(rel_id)) = (attr_value(&e, b"name")?, rel_id) {
                        sheet_entries.push((name, rel_id));
                    }
                }
                b"workbookView" => {
                    if let Some(idx) =
                        attr_value(&e, b"activeTab")?.and_then(|s| s.parse::<usize>().ok())
                    {
                        active_index = idx;
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((sheet_entries, active_index))
}

fn parse_rels<R: BufRead>(reader: R) -> Result<HashMap<String, String>, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = HashMap::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let rel_type = attr_value(&e, b"Type")?;
                if rel_type.as_deref().is_some_and(|t| t.contains("worksheet"))
                    && let (Some(id), Some(target)) =
                        (attr_value(&e, b"Id")?, attr_value(&e, b"Target")?)
                {
                    rels.insert(id, normalize_target(&target));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(rels)
}

fn normalize_target(target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }

    if target.starts_with("xl/") {
        return target.to_string();
    }

    format!("xl/{target}")
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, OxlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::{normalize_target, parse_sheet_refs};
    use std::io::Cursor;

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView activeTab="1"/>
  </bookViews>
  <sheets>
    <sheet name="Empleados" sheetId="1" r:id="rId1"/>
    <sheet name="Productos" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;

    const RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>
</Relationships>"#;

    #[test]
    fn parse_sheet_refs_resolves_order_names_paths_and_active_tab() {
        let (sheets, active_index) = parse_sheet_refs(
            Cursor::new(WORKBOOK_XML.as_bytes()),
            Cursor::new(RELS_XML.as_bytes()),
        )
        .unwrap();

        assert_eq!(active_index, 1);
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Empleados");
        assert_eq!(sheets[1].name, "Productos");
        assert_eq!(sheets[0].zip_path, "xl/worksheets/sheet1.xml");
        assert_eq!(sheets[1].zip_path, "xl/worksheets/sheet2.xml");
    }

    #[test]
    fn normalize_target_prepends_xl_for_relative_paths() {
        assert_eq!(
            normalize_target("worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn normalize_target_strips_leading_slash_for_absolute_paths() {
        assert_eq!(
            normalize_target("/xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }
}
