use crate::OxlsxError;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::io::BufRead;
use std::mem;

#[derive(Debug, Clone, Default)]
pub struct SharedStrings(Vec<String>);

impl SharedStrings {
    pub fn parse<R: BufRead>(reader: R) -> Result<Self, OxlsxError> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut values = Vec::new();
        let mut in_si = false;
        let mut in_t = false;
        let mut current_text = String::new();

        loop {
            match xml.read_event_into(&mut buf)? {
                Event::Start(e) if e.local_name().as_ref() == b"sst" => {
                    let capacity = attr_usize(&e, b"uniqueCount")?.unwrap_or(0);
                    values = Vec::with_capacity(capacity);
                }
                Event::Start(e) if e.local_name().as_ref() == b"si" => {
                    in_si = true;
                    current_text.clear();
                }
                Event::Start(e) if in_si && e.local_name().as_ref() == b"t" => {
                    in_t = true;
                }
                Event::Empty(e) if in_si && e.local_name().as_ref() == b"t" => {
                    in_t = false;
                }
                Event::Text(t) if in_t => {
                    current_text.push_str(&t.unescape()?);
                }
                Event::End(e) if e.local_name().as_ref() == b"t" => {
                    in_t = false;
                }
                Event::End(e) if e.local_name().as_ref() == b"si" => {
                    in_si = false;
                    values.push(mem::take(&mut current_text));
                }
                Event::Eof => break,
                _ => {}
            }

            buf.clear();
        }

        Ok(Self(values))
    }

    pub fn get(&self, index: usize) -> Option<&str> {
        self.0.get(index).map(String::as_str)
    }

    pub fn as_slice(&self) -> &[String] {
        self.0.as_slice()
    }

    pub fn into_inner(self) -> Vec<String> {
        self.0
    }
}

fn attr_usize(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<usize>, OxlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        if attr.key.as_ref() == key {
            return Ok(attr.unescape_value()?.parse::<usize>().ok());
        }
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::SharedStrings;
    use std::io::Cursor;

    #[test]
    fn parses_plain_string_item() {
        let xml = r#"<sst uniqueCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Hola</t></si></sst>"#;
        let shared = SharedStrings::parse(Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(shared.as_slice(), &["Hola"]);
    }

    #[test]
    fn concatenates_rich_text_runs() {
        let xml = r#"<sst uniqueCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>foo</t></r><r><t>bar</t></r></si></sst>"#;
        let shared = SharedStrings::parse(Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(shared.as_slice(), &["foobar"]);
    }

    #[test]
    fn parses_empty_sst() {
        let xml = r#"<sst uniqueCount="0" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>"#;
        let shared = SharedStrings::parse(Cursor::new(xml.as_bytes())).unwrap();
        assert!(shared.as_slice().is_empty());
    }

    #[test]
    fn get_returns_some_for_existing_index() {
        let xml = r#"<sst uniqueCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Hola</t></si></sst>"#;
        let shared = SharedStrings::parse(Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(shared.get(0), Some("Hola"));
    }

    #[test]
    fn get_returns_none_for_out_of_bounds_index() {
        let xml = r#"<sst uniqueCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Hola</t></si></sst>"#;
        let shared = SharedStrings::parse(Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(shared.get(99), None);
    }

    #[test]
    fn malformed_xml_returns_xml_error() {
        use crate::OxlsxError;
        // Invalid XML entity reference — quick-xml returns an error on unescape
        let xml = b"<sst><si><t>&invalid_entity;</t></si></sst>";
        let result = SharedStrings::parse(Cursor::new(xml.as_slice()));
        assert!(
            matches!(result, Err(OxlsxError::Xml(_))),
            "expected OxlsxError::Xml for malformed XML, got: {:?}",
            result
        );
    }
}
