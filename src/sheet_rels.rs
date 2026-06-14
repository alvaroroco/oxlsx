use crate::OxlsxError;
use crate::xml::attr_value;
use quick_xml::Reader;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use std::io::BufRead;

pub const COMMENTS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
pub const VML_DRAWING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
pub const TABLE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetRelEntry {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

pub fn parse_sheet_rels<R: BufRead>(reader: R) -> Result<Vec<SheetRelEntry>, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = Vec::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                if let (Some(id), Some(rel_type), Some(target)) = (
                    attr_value(&e, b"Id")?,
                    attr_value(&e, b"Type")?,
                    attr_value(&e, b"Target")?,
                ) {
                    rels.push(SheetRelEntry {
                        id,
                        rel_type,
                        target,
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(rels)
}

pub fn write_sheet_rels(entries: &[SheetRelEntry]) -> Result<Vec<u8>, OxlsxError> {
    let mut xml = Writer::new(Vec::new());
    xml.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("Relationships");
    root.push_attribute((
        "xmlns",
        "http://schemas.openxmlformats.org/package/2006/relationships",
    ));
    xml.write_event(Event::Start(root))?;

    for entry in entries {
        let mut relationship = BytesStart::new("Relationship");
        relationship.push_attribute(("Id", entry.id.as_str()));
        relationship.push_attribute(("Type", entry.rel_type.as_str()));
        relationship.push_attribute(("Target", entry.target.as_str()));
        xml.write_event(Event::Empty(relationship))?;
    }

    xml.write_event(Event::End(BytesEnd::new("Relationships")))?;
    Ok(xml.into_inner())
}

pub fn upsert_comment_rels(
    existing: &[SheetRelEntry],
    comment_target: &str,
    vml_target: &str,
) -> Vec<SheetRelEntry> {
    let mut merged = existing
        .iter()
        .filter(|entry| {
            entry.rel_type != COMMENTS_REL_TYPE && entry.rel_type != VML_DRAWING_REL_TYPE
        })
        .cloned()
        .collect::<Vec<_>>();

    let comment_id = id_for_type_or_next(existing, COMMENTS_REL_TYPE, &merged);
    merged.push(SheetRelEntry {
        id: comment_id,
        rel_type: COMMENTS_REL_TYPE.to_string(),
        target: comment_target.to_string(),
    });

    // Allocate the vml id AFTER pushing the comment entry, so on a from-scratch
    // sheet the "next id" fallback sees the comment's id and never reuses it
    // (bug #274: both fell through to the same `next` → duplicate `@Id`). The
    // existing-id branch is preserved for RMW.
    let vml_id = id_for_type_or_next(existing, VML_DRAWING_REL_TYPE, &merged);
    merged.push(SheetRelEntry {
        id: vml_id,
        rel_type: VML_DRAWING_REL_TYPE.to_string(),
        target: vml_target.to_string(),
    });

    merged
}

/// Replace this sheet's table relationships with one entry per `table_targets`.
///
/// Mirrors [`upsert_comment_rels`]: unrelated relationships (comments,
/// hyperlinks, printer settings, ...) are preserved untouched. Existing
/// `table` relationships are dropped and re-created so re-numbered table parts
/// never leave a stale entry behind. New rIds are allocated above the highest
/// existing id, guaranteeing no collision with comment/hyperlink rIds.
pub fn upsert_table_rels(existing: &[SheetRelEntry], table_targets: &[&str]) -> Vec<SheetRelEntry> {
    let mut merged = existing
        .iter()
        .filter(|entry| entry.rel_type != TABLE_REL_TYPE)
        .cloned()
        .collect::<Vec<_>>();

    for (next, target) in (next_rel_id(&merged)..).zip(table_targets.iter()) {
        merged.push(SheetRelEntry {
            id: format!("rId{next}"),
            rel_type: TABLE_REL_TYPE.to_string(),
            target: (*target).to_string(),
        });
    }

    merged
}

/// Find the relationship id whose target matches `target`, if any.
pub fn relation_id_for_target<'a>(entries: &'a [SheetRelEntry], target: &str) -> Option<&'a str> {
    entries
        .iter()
        .find(|entry| entry.target == target)
        .map(|entry| entry.id.as_str())
}

/// Smallest unused `rIdN` numeric suffix for the given entries.
fn next_rel_id(entries: &[SheetRelEntry]) -> u32 {
    entries
        .iter()
        .filter_map(|entry| entry.id.strip_prefix("rId"))
        .filter_map(|suffix| suffix.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        + 1
}

pub fn relation_target_for_type<'a>(
    entries: &'a [SheetRelEntry],
    rel_type: &str,
) -> Option<&'a str> {
    entries
        .iter()
        .find(|entry| entry.rel_type == rel_type)
        .map(|entry| entry.target.as_str())
}

pub fn relation_id_for_type<'a>(entries: &'a [SheetRelEntry], rel_type: &str) -> Option<&'a str> {
    entries
        .iter()
        .find(|entry| entry.rel_type == rel_type)
        .map(|entry| entry.id.as_str())
}

/// All relationship targets matching `rel_type`, in document order.
///
/// Unlike [`relation_target_for_type`] (first match only), a sheet may
/// reference several parts of the same type — e.g. multiple `<table>` parts.
pub fn relation_targets_for_type<'a>(entries: &'a [SheetRelEntry], rel_type: &str) -> Vec<&'a str> {
    entries
        .iter()
        .filter(|entry| entry.rel_type == rel_type)
        .map(|entry| entry.target.as_str())
        .collect()
}

pub fn sheet_rels_path(sheet_path: &str) -> String {
    let (parent, file) = sheet_path.rsplit_once('/').unwrap_or(("", sheet_path));

    if parent.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{parent}/_rels/{file}.rels")
    }
}

pub fn resolve_sheet_target(sheet_path: &str, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    if target.starts_with("xl/") {
        return target.to_string();
    }

    let parent = sheet_path
        .rsplit_once('/')
        .map(|(head, _)| head)
        .unwrap_or_default();
    normalize_path(parent, target)
}

fn normalize_path(base: &str, relative: &str) -> String {
    let mut parts = base
        .split('/')
        .filter(|segment| !segment.is_empty())
        .map(ToString::to_string)
        .collect::<Vec<_>>();

    for segment in relative.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value.to_string()),
        }
    }

    parts.join("/")
}

fn id_for_type_or_next(
    existing: &[SheetRelEntry],
    rel_type: &str,
    merged: &[SheetRelEntry],
) -> String {
    if let Some(existing_id) = existing
        .iter()
        .find(|entry| entry.rel_type == rel_type)
        .map(|entry| entry.id.clone())
    {
        return existing_id;
    }

    let next = existing
        .iter()
        .chain(merged.iter())
        .filter_map(|entry| entry.id.strip_prefix("rId"))
        .filter_map(|suffix| suffix.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        + 1;
    format!("rId{next}")
}

#[cfg(test)]
mod tests {
    use super::{
        COMMENTS_REL_TYPE, SheetRelEntry, TABLE_REL_TYPE, VML_DRAWING_REL_TYPE, parse_sheet_rels,
        relation_id_for_target, relation_id_for_type, relation_target_for_type,
        relation_targets_for_type, resolve_sheet_target, sheet_rels_path, upsert_comment_rels,
        upsert_table_rels, write_sheet_rels,
    };
    use std::io::Cursor;

    const SHEET_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#;

    #[test]
    fn parse_sheet_rels_reads_relationship_entries() {
        let rels = parse_sheet_rels(Cursor::new(SHEET_RELS_XML.as_bytes())).unwrap();
        assert_eq!(rels.len(), 2);
        assert_eq!(rels[0].id, "rId1");
        assert!(rels[0].rel_type.ends_with("printerSettings"));
    }

    #[test]
    fn upsert_comment_rels_preserves_unrelated_entries() {
        let existing = vec![
            SheetRelEntry {
                id: "rId1".to_string(),
                rel_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
                        .to_string(),
                target: "../printerSettings/printerSettings1.bin".to_string(),
            },
            SheetRelEntry {
                id: "rId2".to_string(),
                rel_type: COMMENTS_REL_TYPE.to_string(),
                target: "../comments-old.xml".to_string(),
            },
        ];

        let merged =
            upsert_comment_rels(&existing, "../comments1.xml", "../drawings/vmlDrawing1.vml");

        assert!(merged.iter().any(|entry| {
            entry.rel_type.ends_with("printerSettings")
                && entry.target == "../printerSettings/printerSettings1.bin"
        }));
        assert_eq!(
            relation_target_for_type(&merged, COMMENTS_REL_TYPE),
            Some("../comments1.xml")
        );
        assert_eq!(
            relation_target_for_type(&merged, VML_DRAWING_REL_TYPE),
            Some("../drawings/vmlDrawing1.vml")
        );
    }

    #[test]
    fn upsert_comment_rels_from_scratch_yields_distinct_ids() {
        // Bug #274: on a from-scratch sheet both ids were computed against the
        // same `merged` slice (which held neither entry yet), so both fell
        // through to the same "next id" → duplicate `rId1`. They MUST be
        // sequential and distinct.
        let merged = upsert_comment_rels(&[], "../comments1.xml", "../drawings/vmlDrawing1.vml");

        assert_eq!(merged.len(), 2);
        let comment_id = relation_id_for_type(&merged, COMMENTS_REL_TYPE).unwrap();
        let vml_id = relation_id_for_type(&merged, VML_DRAWING_REL_TYPE).unwrap();
        assert_ne!(comment_id, vml_id, "comment and vml rIds must be distinct");
        assert_eq!(comment_id, "rId1");
        assert_eq!(vml_id, "rId2");
    }

    #[test]
    fn upsert_comment_rels_preserves_existing_comment_id() {
        // RMW: an existing comment rId3 must be kept; the freshly allocated vml
        // id must be distinct from it.
        let existing = vec![SheetRelEntry {
            id: "rId3".to_string(),
            rel_type: COMMENTS_REL_TYPE.to_string(),
            target: "../comments-old.xml".to_string(),
        }];

        let merged =
            upsert_comment_rels(&existing, "../comments1.xml", "../drawings/vmlDrawing1.vml");

        let comment_id = relation_id_for_type(&merged, COMMENTS_REL_TYPE).unwrap();
        let vml_id = relation_id_for_type(&merged, VML_DRAWING_REL_TYPE).unwrap();
        assert_eq!(comment_id, "rId3", "existing comment id must be preserved");
        assert_ne!(comment_id, vml_id, "ids must stay distinct");
    }

    #[test]
    fn upsert_comment_rels_preserves_existing_vml_id() {
        // RMW: an existing vml rId2 must be kept; the freshly allocated comment
        // id must be distinct from it.
        let existing = vec![SheetRelEntry {
            id: "rId2".to_string(),
            rel_type: VML_DRAWING_REL_TYPE.to_string(),
            target: "../drawings/vmlDrawing-old.vml".to_string(),
        }];

        let merged =
            upsert_comment_rels(&existing, "../comments1.xml", "../drawings/vmlDrawing1.vml");

        let comment_id = relation_id_for_type(&merged, COMMENTS_REL_TYPE).unwrap();
        let vml_id = relation_id_for_type(&merged, VML_DRAWING_REL_TYPE).unwrap();
        assert_eq!(vml_id, "rId2", "existing vml id must be preserved");
        assert_ne!(comment_id, vml_id, "ids must stay distinct");
    }

    #[test]
    fn write_sheet_rels_serializes_relationships_xml() {
        let rels = vec![SheetRelEntry {
            id: "rId1".to_string(),
            rel_type: COMMENTS_REL_TYPE.to_string(),
            target: "../comments1.xml".to_string(),
        }];
        let xml = String::from_utf8(write_sheet_rels(&rels).unwrap()).unwrap();
        assert!(xml.contains("Relationship"));
        assert!(xml.contains("../comments1.xml"));
    }

    #[test]
    fn upsert_table_rels_preserves_unrelated_rels_and_avoids_id_collision() {
        let existing = vec![
            SheetRelEntry {
                id: "rId1".to_string(),
                rel_type: COMMENTS_REL_TYPE.to_string(),
                target: "../comments1.xml".to_string(),
            },
            SheetRelEntry {
                id: "rId2".to_string(),
                rel_type: VML_DRAWING_REL_TYPE.to_string(),
                target: "../drawings/vmlDrawing1.vml".to_string(),
            },
            SheetRelEntry {
                id: "rId3".to_string(),
                rel_type: TABLE_REL_TYPE.to_string(),
                target: "../tables/stale.xml".to_string(),
            },
        ];

        let merged =
            upsert_table_rels(&existing, &["../tables/table1.xml", "../tables/table2.xml"]);

        assert_eq!(
            relation_target_for_type(&merged, COMMENTS_REL_TYPE),
            Some("../comments1.xml")
        );
        assert_eq!(
            relation_target_for_type(&merged, VML_DRAWING_REL_TYPE),
            Some("../drawings/vmlDrawing1.vml")
        );

        let table_targets = relation_targets_for_type(&merged, TABLE_REL_TYPE);
        assert_eq!(
            table_targets,
            vec!["../tables/table1.xml", "../tables/table2.xml"]
        );

        let rid1 = relation_id_for_target(&merged, "../tables/table1.xml").unwrap();
        let rid2 = relation_id_for_target(&merged, "../tables/table2.xml").unwrap();
        assert_ne!(rid1, "rId1");
        assert_ne!(rid1, "rId2");
        assert_ne!(rid1, rid2);
    }

    #[test]
    fn relation_targets_for_type_returns_all_matches_in_order() {
        let entries = vec![
            SheetRelEntry {
                id: "rId1".to_string(),
                rel_type: COMMENTS_REL_TYPE.to_string(),
                target: "../comments1.xml".to_string(),
            },
            SheetRelEntry {
                id: "rId2".to_string(),
                rel_type: TABLE_REL_TYPE.to_string(),
                target: "../tables/table1.xml".to_string(),
            },
            SheetRelEntry {
                id: "rId3".to_string(),
                rel_type: TABLE_REL_TYPE.to_string(),
                target: "/xl/tables/table2.xml".to_string(),
            },
        ];

        let targets = relation_targets_for_type(&entries, TABLE_REL_TYPE);
        assert_eq!(
            targets,
            vec!["../tables/table1.xml", "/xl/tables/table2.xml"]
        );
    }

    #[test]
    fn resolve_sheet_target_normalizes_relative_paths() {
        assert_eq!(
            resolve_sheet_target("xl/worksheets/sheet1.xml", "../comments1.xml"),
            "xl/comments1.xml"
        );
        assert_eq!(
            sheet_rels_path("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
    }
}
