use crate::OxlsxError;
use crate::xml::attr_value;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::BufRead;
use std::mem;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub text: String,
    pub author: String,
}

pub fn parse_comments_xml<R: BufRead>(reader: R) -> Result<HashMap<String, Comment>, OxlsxError> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut comments = HashMap::new();
    let mut authors = Vec::<String>::new();

    let mut in_author = false;
    let mut in_text_node = false;

    let mut current_ref: Option<String> = None;
    let mut current_author_id = 0usize;
    let mut current_text = String::new();

    loop {
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"comment" => {
                current_ref = attr_value(&e, b"ref")?;
                current_author_id = attr_value(&e, b"authorId")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .unwrap_or(0);
                current_text.clear();
            }
            Event::End(e) if e.local_name().as_ref() == b"comment" => {
                if let Some(reference) = current_ref.take() {
                    let author = authors
                        .get(current_author_id)
                        .cloned()
                        .unwrap_or_else(String::new);
                    comments.insert(
                        reference,
                        Comment {
                            text: mem::take(&mut current_text),
                            author,
                        },
                    );
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"author" => {
                in_author = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"author" => {
                in_author = false;
            }
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                in_text_node = current_ref.is_some();
            }
            Event::End(e) if e.local_name().as_ref() == b"t" => {
                in_text_node = false;
            }
            Event::Text(text) if in_author => {
                authors.push(text.unescape()?.into_owned());
            }
            Event::Text(text) if in_text_node => {
                current_text.push_str(&text.unescape()?);
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(comments)
}

#[cfg(test)]
mod tests {
    use super::parse_comments_xml;
    use std::io::Cursor;

    const COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Alice</author>
    <author>Bob</author>
  </authors>
  <commentList>
    <comment ref="B2" authorId="0">
      <text>
        <t>Hello</t>
      </text>
    </comment>
    <comment ref="C3" authorId="1">
      <text>
        <r><t>Part1 </t></r>
        <r><t>Part2</t></r>
      </text>
    </comment>
  </commentList>
</comments>"#;

    #[test]
    fn parse_comments_xml_reads_plain_text_and_author() {
        let comments = parse_comments_xml(Cursor::new(COMMENTS_XML.as_bytes())).unwrap();

        let b2 = comments.get("B2").unwrap();
        assert_eq!(b2.text, "Hello");
        assert_eq!(b2.author, "Alice");

        let c3 = comments.get("C3").unwrap();
        assert_eq!(c3.text, "Part1Part2");
        assert_eq!(c3.author, "Bob");
    }
}
