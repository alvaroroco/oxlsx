use crate::OxlsxError;
use crate::xml::attr_value;
use quick_xml::Reader;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use std::collections::HashMap;

/// Newtype for a cell's `s` attribute — the index into `cellXfs`.
/// Prevents confusing a `StyleId` with a raw `font_id` or `fill_id`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct StyleId(pub usize);

impl StyleId {
    pub fn as_usize(self) -> usize {
        self.0
    }
}

impl From<u32> for StyleId {
    fn from(v: u32) -> Self {
        StyleId(v as usize)
    }
}

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

impl Fill {
    #[allow(dead_code)]
    pub(crate) fn none() -> Self {
        Self {
            pattern_type: "none".to_string(),
            fg_color: None,
            bg_color: None,
        }
    }

    #[allow(dead_code)]
    pub(crate) fn gray125() -> Self {
        Self {
            pattern_type: "gray125".to_string(),
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
    pub num_fmt_id: u32,
}

#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub cell_xfs: Vec<CellXf>,
    num_fmts: HashMap<u32, String>,
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
        let mut in_num_fmts = false;

        let mut cur_font = Font::default();
        let mut cur_fill = Fill::default();

        let mut sheet = StyleSheet::default();

        loop {
            match xml.read_event_into(&mut buf)? {
                Event::Start(e) => match e.local_name().as_ref() {
                    b"fonts" => in_fonts = true,
                    b"fills" => in_fills = true,
                    b"cellXfs" => in_cell_xfs = true,
                    b"numFmts" => {
                        in_num_fmts = true;
                        if let Some(count) =
                            attr_value(&e, b"count")?.and_then(|v| v.parse::<usize>().ok())
                        {
                            sheet.num_fmts = HashMap::with_capacity(count);
                        }
                    }
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
                            num_fmt_id: attr_value(&e, b"numFmtId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                        });
                    }
                    b"numFmt" if in_num_fmts => {
                        if let (Some(num_fmt_id), Some(format_code)) = (
                            attr_value(&e, b"numFmtId")?.and_then(|v| v.parse::<u32>().ok()),
                            attr_value(&e, b"formatCode")?,
                        ) {
                            sheet.num_fmts.insert(num_fmt_id, format_code);
                        }
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
                            num_fmt_id: attr_value(&e, b"numFmtId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                        });
                    }
                    b"numFmt" if in_num_fmts => {
                        if let (Some(num_fmt_id), Some(format_code)) = (
                            attr_value(&e, b"numFmtId")?.and_then(|v| v.parse::<u32>().ok()),
                            attr_value(&e, b"formatCode")?,
                        ) {
                            sheet.num_fmts.insert(num_fmt_id, format_code);
                        }
                    }
                    _ => {}
                },
                Event::End(e) => match e.local_name().as_ref() {
                    b"font" if in_fonts => {
                        in_font = false;
                        sheet.fonts.push(std::mem::take(&mut cur_font));
                    }
                    b"fill" if in_fills => {
                        in_fill = false;
                        in_pattern_fill = false;
                        sheet.fills.push(std::mem::take(&mut cur_fill));
                    }
                    b"patternFill" => in_pattern_fill = false,
                    b"fonts" => in_fonts = false,
                    b"fills" => in_fills = false,
                    b"cellXfs" => in_cell_xfs = false,
                    b"numFmts" => in_num_fmts = false,
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }

            buf.clear();
        }

        Ok(sheet)
    }

    pub fn font(&self, style_id: StyleId) -> Option<&Font> {
        let xf = self.cell_xfs.get(style_id.as_usize())?;
        self.fonts.get(xf.font_id as usize)
    }

    pub fn fill(&self, style_id: StyleId) -> Option<&Fill> {
        let xf = self.cell_xfs.get(style_id.as_usize())?;
        self.fills.get(xf.fill_id as usize)
    }

    pub fn num_fmt_id(&self, style_id: StyleId) -> Option<u32> {
        self.cell_xfs
            .get(style_id.as_usize())
            .map(|xf| xf.num_fmt_id)
    }

    pub fn is_date_format(&self, style_id: StyleId) -> bool {
        match self.num_fmt_id(style_id) {
            Some(id) if (14..=17).contains(&id) || id == 22 => true,
            Some(id) if id >= 164 => self
                .num_fmts
                .get(&id)
                .map(|fmt| fmt.contains('y') || fmt.contains('d'))
                .unwrap_or(false),
            _ => false,
        }
    }

    pub(crate) fn num_fmts_vec(&self) -> Vec<(u32, String)> {
        self.num_fmts
            .iter()
            .map(|(num_fmt_id, format_code)| (*num_fmt_id, format_code.clone()))
            .collect()
    }
}

fn parse_color(e: &BytesStart<'_>) -> Result<Option<Color>, OxlsxError> {
    if let Some(rgb) = attr_value(e, b"rgb")? {
        return Ok(Some(Color::Argb(rgb)));
    }
    if let Some(indexed) = attr_value(e, b"indexed")?
        && let Ok(value) = indexed.parse::<u32>()
    {
        return Ok(Some(Color::Indexed(value)));
    }
    if let Some(theme) = attr_value(e, b"theme")?
        && let Ok(value) = theme.parse::<u32>()
    {
        return Ok(Some(Color::Theme(value)));
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::{Color, StyleId, StyleSheet};
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
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
  </numFmts>
  <cellXfs count="4">
    <xf fontId="0" fillId="0" numFmtId="0"/>
    <xf fontId="1" fillId="2" numFmtId="14"/>
    <xf fontId="2" fillId="3" numFmtId="164"/>
    <xf fontId="0" fillId="0" numFmtId="2"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn parses_styles_collections_and_attributes() {
        let sheet = StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap();

        assert_eq!(sheet.fonts.len(), 3);
        assert_eq!(sheet.fills.len(), 4);
        assert_eq!(sheet.cell_xfs.len(), 4);
        assert!(sheet.fonts[1].bold);
        assert_eq!(sheet.fonts[2].color, Some(Color::Indexed(4)));
        assert_eq!(sheet.fills[2].pattern_type, "solid");
        assert_eq!(sheet.fills[2].fg_color, Some(Color::Indexed(2)));
        assert_eq!(sheet.fills[3].fg_color, Some(Color::Indexed(5)));
    }

    #[test]
    fn resolves_fonts_and_fills_by_style_id() {
        let sheet = StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap();

        assert!(sheet.font(StyleId(1)).unwrap().bold);
        assert_eq!(
            sheet.font(StyleId(2)).unwrap().color,
            Some(Color::Indexed(4))
        );
    }

    #[test]
    fn resolves_num_fmt_id_and_date_formats() {
        let sheet = StyleSheet::parse(Cursor::new(SAMPLE_STYLES_XML.as_bytes())).unwrap();

        assert_eq!(sheet.num_fmt_id(StyleId(2)), Some(164));
        assert!(sheet.is_date_format(StyleId(1)));
        assert!(sheet.is_date_format(StyleId(2)));
        assert!(!sheet.is_date_format(StyleId(0)));
        assert!(!sheet.is_date_format(StyleId(3)));
    }
}
