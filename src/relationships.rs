use crate::OxlsxError;
use quick_xml::Reader;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use std::io::BufRead;

pub const HYPERLINK_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetRelationship {
    Hyperlink {
        id: String,
        target: String,
        target_mode: Option<String>,
    },
    Unknown {
        id: String,
        rel_type: String,
        target: String,
        target_mode: Option<String>,
    },
}

impl SheetRelationship {
    pub fn hyperlink(id: impl Into<String>, target: impl Into<String>) -> Self {
        Self::Hyperlink {
            id: id.into(),
            target: target.into(),
            target_mode: Some("External".to_string()),
        }
    }

    pub fn id(&self) -> &str {
        match self {
            SheetRelationship::Hyperlink { id, .. } | SheetRelationship::Unknown { id, .. } => id,
        }
    }

    pub fn rel_type(&self) -> &str {
        match self {
            SheetRelationship::Hyperlink { .. } => HYPERLINK_REL_TYPE,
            SheetRelationship::Unknown { rel_type, .. } => rel_type,
        }
    }

    pub fn target(&self) -> &str {
        match self {
            SheetRelationship::Hyperlink { target, .. }
            | SheetRelationship::Unknown { target, .. } => target,
        }
    }

    pub fn target_mode(&self) -> Option<&str> {
        match self {
            SheetRelationship::Hyperlink { target_mode, .. }
            | SheetRelationship::Unknown { target_mode, .. } => target_mode.as_deref(),
        }
    }

    pub fn hyperlink_target_by_id<'a>(rels: &'a [SheetRelationship], id: &str) -> Option<&'a str> {
        rels.iter().find_map(|rel| match rel {
            SheetRelationship::Hyperlink { id: rel_id, target, .. } if rel_id == id => {
                Some(target.as_str())
            }
            _ => None,
        })
    }
}

pub fn sheet_rels_path(sheet_xml_path: &str) -> String {
    let filename = sheet_xml_path.rsplit('/').next().unwrap_or(sheet_xml_path);
    format!("xl/worksheets/_rels/{filename}.rels")
}

pub fn parse_sheet_rels<R: BufRead>(reader: R) -> Result<Vec<SheetRelationship>, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = Vec::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let id = attr_value(&e, b"Id")?;
                let rel_type = attr_value(&e, b"Type")?;
                let target = attr_value(&e, b"Target")?;
                let target_mode = attr_value(&e, b"TargetMode")?;

                if let (Some(id), Some(rel_type), Some(target)) = (id, rel_type, target) {
                    if rel_type == HYPERLINK_REL_TYPE {
                        rels.push(SheetRelationship::Hyperlink {
                            id,
                            target,
                            target_mode,
                        });
                    } else {
                        rels.push(SheetRelationship::Unknown {
                            id,
                            rel_type,
                            target,
                            target_mode,
                        });
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(rels)
}

pub fn serialize_sheet_rels(rels: &[SheetRelationship]) -> Result<Vec<u8>, OxlsxError> {
    let mut xml = Writer::new(Vec::new());
    xml.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    write_start(
        &mut xml,
        "Relationships",
        &[(
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        )],
    )?;

    for rel in rels {
        match rel {
            SheetRelationship::Hyperlink {
                id,
                target,
                target_mode,
                ..
            } => {
                let mut attrs = vec![
                    ("Id", id.as_str()),
                    ("Type", HYPERLINK_REL_TYPE),
                    ("Target", target.as_str()),
                ];

                if let Some(target_mode) = target_mode {
                    attrs.push(("TargetMode", target_mode.as_str()));
                }

                write_empty(&mut xml, "Relationship", &attrs)?;
            }
            SheetRelationship::Unknown {
                id,
                rel_type,
                target,
                target_mode,
            } => {
                let mut attrs = vec![
                    ("Id", id.as_str()),
                    ("Type", rel_type.as_str()),
                    ("Target", target.as_str()),
                ];

                if let Some(target_mode) = target_mode {
                    attrs.push(("TargetMode", target_mode.as_str()));
                }

                write_empty(&mut xml, "Relationship", &attrs)?;
            }
        }
    }

    write_end(&mut xml, "Relationships")?;

    Ok(xml.into_inner())
}

pub fn allocate_next_rid(existing: &[SheetRelationship]) -> String {
    let next = existing
        .iter()
        .filter_map(|rel| rel.id().strip_prefix("rId")?.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        + 1;

    format!("rId{next}")
}

fn write_start(
    xml: &mut Writer<Vec<u8>>,
    name: &str,
    attrs: &[(&str, &str)],
) -> Result<(), OxlsxError> {
    let mut elem = BytesStart::new(name);
    for (k, v) in attrs {
        elem.push_attribute((*k, *v));
    }
    xml.write_event(Event::Start(elem))?;
    Ok(())
}

fn write_empty(
    xml: &mut Writer<Vec<u8>>,
    name: &str,
    attrs: &[(&str, &str)],
) -> Result<(), OxlsxError> {
    let mut elem = BytesStart::new(name);
    for (k, v) in attrs {
        elem.push_attribute((*k, *v));
    }
    xml.write_event(Event::Empty(elem))?;
    Ok(())
}

fn write_end(xml: &mut Writer<Vec<u8>>, name: &str) -> Result<(), OxlsxError> {
    xml.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(())
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
    use super::{
        SheetRelationship, allocate_next_rid, parse_sheet_rels, serialize_sheet_rels,
    };
    use std::io::Cursor;

    const SHEET_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    #[test]
    fn rels_parse_and_serialize_round_trip() {
        let rels = parse_sheet_rels(Cursor::new(SHEET_RELS_XML.as_bytes())).unwrap();
        let bytes = serialize_sheet_rels(&rels).unwrap();
        let reparsed = parse_sheet_rels(Cursor::new(bytes)).unwrap();

        assert_eq!(reparsed, rels);
    }

    #[test]
    fn rels_allocate_next_rid_after_preserved_unknowns() {
        let rels = vec![
            SheetRelationship::Unknown {
                id: "rId2".to_string(),
                rel_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
                        .to_string(),
                target: "../comments1.xml".to_string(),
                target_mode: None,
            },
            SheetRelationship::Unknown {
                id: "rId9".to_string(),
                rel_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
                        .to_string(),
                target: "../drawings/vmlDrawing1.vml".to_string(),
                target_mode: None,
            },
        ];

        assert_eq!(allocate_next_rid(&rels), "rId10");
    }
}
