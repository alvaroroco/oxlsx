use crate::OxlsxError;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::Reader;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Color {
    Argb(String),
    Indexed(u32),
    Theme(u32),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<Color>,
    pub bg_color: Option<Color>,
}

impl Default for Fill {
    fn default() -> Self {
        Self {
            pattern_type: "none".to_string(),
            fg_color: None,
            bg_color: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Font {
    pub name: String,
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<Color>,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            name: String::new(),
            size: 0.0,
            bold: false,
            italic: false,
            color: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellXf {
    pub font_id: u32,
    pub fill_id: u32,
}

#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub cell_xfs: Vec<CellXf>,
}

impl StyleSheet {
    pub fn parse<R: std::io::BufRead>(reader: R) -> Result<Self, OxlsxError> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut buf = Vec::new();

        let mut in_fonts = false;
        let mut in_fills = false;
        let mut in_cell_xfs = false;
        let mut in_font = false;
        let mut in_fill = false;
        let mut in_pattern_fill = false;

        let mut cur_font = Font::default();
        let mut cur_fill = Fill::default();

        let mut sheet = StyleSheet::default();

        loop {
            match xml.read_event_into(&mut buf)? {
                Event::Start(e) => match e.local_name().as_ref() {
                    b"fonts" => in_fonts = true,
                    b"fills" => in_fills = true,
                    b"cellXfs" => in_cell_xfs = true,
                    b"font" if in_fonts => {
                        in_font = true;
                        cur_font = Font::default();
                    }
                    b"fill" if in_fills => {
                        in_fill = true;
                        cur_fill = Fill::default();
                    }
                    b"patternFill" if in_fill => {
                        in_pattern_fill = true;
                        cur_fill.pattern_type =
                            attr_value(&e, b"patternType")?.unwrap_or_else(|| "none".to_string());
                    }
                    b"sz" if in_font => {
                        if let Some(val) = attr_value(&e, b"val")? {
                            cur_font.size = val.parse().unwrap_or(0.0);
                        }
                    }
                    b"name" if in_font => {
                        if let Some(val) = attr_value(&e, b"val")? {
                            cur_font.name = val;
                        }
                    }
                    b"color" if in_font => {
                        cur_font.color = parse_color(&e)?;
                    }
                    b"fgColor" if in_pattern_fill => {
                        cur_fill.fg_color = parse_color(&e)?;
                    }
                    b"bgColor" if in_pattern_fill => {
                        cur_fill.bg_color = parse_color(&e)?;
                    }
                    b"xf" if in_cell_xfs => {
                        sheet.cell_xfs.push(CellXf {
                            font_id: attr_value(&e, b"fontId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            fill_id: attr_value(&e, b"fillId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                        });
                    }
                    _ => {}
                },
                Event::Empty(e) => match e.local_name().as_ref() {
                    b"b" if in_font => cur_font.bold = true,
                    b"i" if in_font => cur_font.italic = true,
                    b"sz" if in_font => {
                        if let Some(val) = attr_value(&e, b"val")? {
                            cur_font.size = val.parse().unwrap_or(0.0);
                        }
                    }
                    b"name" if in_font => {
                        if let Some(val) = attr_value(&e, b"val")? {
                            cur_font.name = val;
                        }
                    }
                    b"color" if in_font => {
                        cur_font.color = parse_color(&e)?;
                    }
                    b"patternFill" if in_fill => {
                        cur_fill.pattern_type =
                            attr_value(&e, b"patternType")?.unwrap_or_else(|| "none".to_string());
                    }
                    b"fgColor" if in_fill => cur_fill.fg_color = parse_color(&e)?,
                    b"bgColor" if in_fill => cur_fill.bg_color = parse_color(&e)?,
                    b"xf" if in_cell_xfs => {
                        sheet.cell_xfs.push(CellXf {
                            font_id: attr_value(&e, b"fontId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            fill_id: attr_value(&e, b"fillId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                        });
                    }
                    _ => {}
                },
                Event::End(e) => match e.local_name().as_ref() {
                    b"font" if in_fonts => {
                        in_font = false;
                        sheet.fonts.push(cur_font.clone());
                    }
                    b"fill" if in_fills => {
                        in_fill = false;
                        in_pattern_fill = false;
                        sheet.fills.push(cur_fill.clone());
                    }
                    b"patternFill" => in_pattern_fill = false,
                    b"fonts" => in_fonts = false,
                    b"fills" => in_fills = false,
                    b"cellXfs" => in_cell_xfs = false,
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }

            buf.clear();
        }

        Ok(sheet)
    }

    pub fn font(&self, style_id: u32) -> Option<&Font> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.fonts.get(xf.font_id as usize)
    }

    pub fn fill(&self, style_id: u32) -> Option<&Fill> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.fills.get(xf.fill_id as usize)
    }
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, OxlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        if attr.key.as_ref() == key {
            let val = attr.unescape_value()?.into_owned();
            return Ok(Some(val));
        }
    }

    Ok(None)
}

fn parse_color(e: &BytesStart<'_>) -> Result<Option<Color>, OxlsxError> {
    if let Some(rgb) = attr_value(e, b"rgb")? {
        return Ok(Some(Color::Argb(rgb)));
    }
    if let Some(indexed) = attr_value(e, b"indexed")? {
        if let Ok(value) = indexed.parse::<u32>() {
            return Ok(Some(Color::Indexed(value)));
        }
    }
    if let Some(theme) = attr_value(e, b"theme")? {
        if let Ok(value) = theme.parse::<u32>() {
            return Ok(Some(Color::Theme(value)));
        }
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::{Color, StyleSheet};
    use std::io::Cursor;

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
  <cellXfs count="3">
    <xf fontId="0" fillId="0"/>
    <xf fontId="1" fillId="2"/>
    <xf fontId="2" fillId="3"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn parses_styles_collections_and_attributes() {
        let sheet = StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap();

        assert_eq!(sheet.fonts.len(), 3);
        assert_eq!(sheet.fills.len(), 4);
        assert_eq!(sheet.cell_xfs.len(), 3);
        assert!(sheet.fonts[1].bold);
        assert_eq!(sheet.fonts[2].color, Some(Color::Indexed(4)));
        assert_eq!(sheet.fills[2].pattern_type, "solid");
        assert_eq!(sheet.fills[2].fg_color, Some(Color::Indexed(2)));
        assert_eq!(sheet.fills[3].fg_color, Some(Color::Indexed(5)));
    }

    #[test]
    fn resolves_fonts_and_fills_by_style_id() {
        let sheet = StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap();

        assert!(sheet.font(1).unwrap().bold);
        assert_eq!(sheet.font(2).unwrap().color, Some(Color::Indexed(4)));
    }
}
