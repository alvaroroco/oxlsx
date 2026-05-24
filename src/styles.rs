use crate::OxlsxError;
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
pub struct Alignment {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub indent: u32,
}

impl Default for Alignment {
    fn default() -> Self {
        Self {
            horizontal: None,
            vertical: None,
            wrap_text: false,
            indent: 0,
        }
    }
}

impl Alignment {
    pub(crate) fn to_xml_attributes(&self) -> HashMap<String, String> {
        let mut attrs = HashMap::new();
        if let Some(horizontal) = &self.horizontal {
            attrs.insert("horizontal".to_string(), horizontal.clone());
        }
        if let Some(vertical) = &self.vertical {
            attrs.insert("vertical".to_string(), vertical.clone());
        }
        if self.wrap_text {
            attrs.insert("wrapText".to_string(), "1".to_string());
        }
        if self.indent > 0 {
            attrs.insert("indent".to_string(), self.indent.to_string());
        }
        attrs
    }

    pub(crate) fn from_xml_attributes(attrs: &HashMap<String, String>) -> Self {
        Self {
            horizontal: attrs.get("horizontal").cloned(),
            vertical: attrs.get("vertical").cloned(),
            wrap_text: parse_bool_like(attrs.get("wrapText").map(String::as_str)),
            indent: attrs
                .get("indent")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or(0),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Side {
    pub style: Option<String>,
    pub color: Option<Color>,
}

impl Side {
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn to_xml_parts(&self) -> (Option<String>, Option<Color>) {
        (self.style.clone(), self.color.clone())
    }

    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn from_xml_parts(style: Option<&str>, color: Option<&Color>) -> Self {
        Self {
            style: style.map(str::to_owned),
            color: color.cloned(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Border {
    pub left: Side,
    pub right: Side,
    pub top: Side,
    pub bottom: Side,
    pub diagonal: Side,
}

impl Border {
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn to_xml_sides(&self) -> HashMap<String, Side> {
        HashMap::from([
            ("left".to_string(), self.left.clone()),
            ("right".to_string(), self.right.clone()),
            ("top".to_string(), self.top.clone()),
            ("bottom".to_string(), self.bottom.clone()),
            ("diagonal".to_string(), self.diagonal.clone()),
        ])
    }

    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn from_xml_sides(sides: &HashMap<String, Side>) -> Self {
        Self {
            left: sides.get("left").cloned().unwrap_or_default(),
            right: sides.get("right").cloned().unwrap_or_default(),
            top: sides.get("top").cloned().unwrap_or_default(),
            bottom: sides.get("bottom").cloned().unwrap_or_default(),
            diagonal: sides.get("diagonal").cloned().unwrap_or_default(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Protection {
    pub locked: bool,
    pub hidden: bool,
}

impl Default for Protection {
    fn default() -> Self {
        Self {
            locked: true,
            hidden: false,
        }
    }
}

impl Protection {
    pub(crate) fn to_xml_attributes(&self) -> HashMap<String, String> {
        let mut attrs = HashMap::new();
        attrs.insert(
            "locked".to_string(),
            if self.locked { "1" } else { "0" }.to_string(),
        );
        attrs.insert(
            "hidden".to_string(),
            if self.hidden { "1" } else { "0" }.to_string(),
        );
        attrs
    }

    pub(crate) fn from_xml_attributes(attrs: &HashMap<String, String>) -> Self {
        let mut protection = Self::default();
        if let Some(locked) = attrs.get("locked") {
            protection.locked = parse_bool_like(Some(locked.as_str()));
        }
        if let Some(hidden) = attrs.get("hidden") {
            protection.hidden = parse_bool_like(Some(hidden.as_str()));
        }
        protection
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellXf {
    pub font_id: u32,
    pub fill_id: u32,
    pub num_fmt_id: u32,
    pub border_id: u32,
    pub alignment: Option<Alignment>,
    pub protection: Option<Protection>,
    pub apply_font: bool,
    pub apply_fill: bool,
    pub apply_border: bool,
    pub apply_number_format: bool,
    pub apply_alignment: bool,
    pub apply_protection: bool,
}

impl CellXf {
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn to_xml_attributes(&self) -> HashMap<String, String> {
        HashMap::from([
            ("fontId".to_string(), self.font_id.to_string()),
            ("fillId".to_string(), self.fill_id.to_string()),
            ("numFmtId".to_string(), self.num_fmt_id.to_string()),
            ("borderId".to_string(), self.border_id.to_string()),
            (
                "applyFont".to_string(),
                if self.apply_font { "1" } else { "0" }.to_string(),
            ),
            (
                "applyFill".to_string(),
                if self.apply_fill { "1" } else { "0" }.to_string(),
            ),
            (
                "applyBorder".to_string(),
                if self.apply_border { "1" } else { "0" }.to_string(),
            ),
            (
                "applyNumberFormat".to_string(),
                if self.apply_number_format { "1" } else { "0" }.to_string(),
            ),
            (
                "applyAlignment".to_string(),
                if self.apply_alignment { "1" } else { "0" }.to_string(),
            ),
            (
                "applyProtection".to_string(),
                if self.apply_protection { "1" } else { "0" }.to_string(),
            ),
        ])
    }

    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn from_xml_attributes(
        attrs: &HashMap<String, String>,
        alignment: Option<Alignment>,
        protection: Option<Protection>,
    ) -> Self {
        Self {
            font_id: attrs
                .get("fontId")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or(0),
            fill_id: attrs
                .get("fillId")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or(0),
            num_fmt_id: attrs
                .get("numFmtId")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or(0),
            border_id: attrs
                .get("borderId")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or(0),
            apply_font: parse_bool_like(attrs.get("applyFont").map(String::as_str)),
            apply_fill: parse_bool_like(attrs.get("applyFill").map(String::as_str)),
            apply_border: parse_bool_like(attrs.get("applyBorder").map(String::as_str)),
            apply_number_format: parse_bool_like(
                attrs.get("applyNumberFormat").map(String::as_str),
            ),
            apply_alignment: parse_bool_like(attrs.get("applyAlignment").map(String::as_str)),
            apply_protection: parse_bool_like(attrs.get("applyProtection").map(String::as_str)),
            alignment,
            protection,
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
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
        let mut in_borders = false;
        let mut in_border = false;
        let mut in_font = false;
        let mut in_fill = false;
        let mut in_pattern_fill = false;
        let mut in_num_fmts = false;
        let mut in_xf = false;
        let mut in_side = false;

        let mut cur_font = Font::default();
        let mut cur_fill = Fill::default();
        let mut cur_border = Border::default();
        let mut cur_side = Side::default();
        let mut cur_side_name: Option<String> = None;
        let mut cur_xf: Option<CellXf> = None;

        let mut sheet = StyleSheet::default();

        loop {
            match xml.read_event_into(&mut buf)? {
                Event::Start(e) => match e.local_name().as_ref() {
                    b"fonts" => in_fonts = true,
                    b"fills" => in_fills = true,
                    b"borders" => in_borders = true,
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
                    b"border" if in_borders => {
                        in_border = true;
                        cur_border = Border::default();
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
                        in_xf = true;
                        cur_xf = Some(CellXf {
                            font_id: attr_value(&e, b"fontId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            fill_id: attr_value(&e, b"fillId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            num_fmt_id: attr_value(&e, b"numFmtId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            border_id: attr_value(&e, b"borderId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            alignment: None,
                            protection: None,
                            apply_font: attr_value(&e, b"applyFont")
                                .ok()
                                .flatten()
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_fill: attr_value(&e, b"applyFill")
                                .ok()
                                .flatten()
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_border: attr_value(&e, b"applyBorder")
                                .ok()
                                .flatten()
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_number_format: attr_value(&e, b"applyNumberFormat")
                                .ok()
                                .flatten()
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_alignment: attr_value(&e, b"applyAlignment")
                                .ok()
                                .flatten()
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_protection: attr_value(&e, b"applyProtection")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                        });
                    }
                    b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_border => {
                        in_side = true;
                        cur_side_name = Some(String::from_utf8_lossy(e.local_name().as_ref()).to_string());
                        cur_side = Side {
                            style: attr_value(&e, b"style")?,
                            color: None,
                        };
                    }
                    b"color" if in_side => {
                        cur_side.color = parse_color(&e)?;
                    }
                    b"alignment" if in_xf => {
                        if let Some(xf) = cur_xf.as_mut() {
                            let attrs = parse_attrs_map(&e)?;
                            xf.alignment = Some(Alignment::from_xml_attributes(&attrs));
                        }
                    }
                    b"protection" if in_xf => {
                        if let Some(xf) = cur_xf.as_mut() {
                            let attrs = parse_attrs_map(&e)?;
                            xf.protection = Some(Protection::from_xml_attributes(&attrs));
                        }
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
                            border_id: attr_value(&e, b"borderId")?
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(0),
                            alignment: None,
                            protection: None,
                            apply_font: attr_value(&e, b"applyFont")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_fill: attr_value(&e, b"applyFill")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_border: attr_value(&e, b"applyBorder")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_number_format: attr_value(&e, b"applyNumberFormat")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_alignment: attr_value(&e, b"applyAlignment")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                            apply_protection: attr_value(&e, b"applyProtection")?
                                .is_some_and(|v| parse_bool_like(Some(v.as_str()))),
                        });
                    }
                    b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_border => {
                        let side = Side {
                            style: attr_value(&e, b"style")?,
                            color: None,
                        };
                        assign_side(&mut cur_border, e.local_name().as_ref(), side);
                    }
                    b"color" if in_side => {
                        cur_side.color = parse_color(&e)?;
                    }
                    b"alignment" if in_xf => {
                        if let Some(xf) = cur_xf.as_mut() {
                            let attrs = parse_attrs_map(&e)?;
                            xf.alignment = Some(Alignment::from_xml_attributes(&attrs));
                        }
                    }
                    b"protection" if in_xf => {
                        if let Some(xf) = cur_xf.as_mut() {
                            let attrs = parse_attrs_map(&e)?;
                            xf.protection = Some(Protection::from_xml_attributes(&attrs));
                        }
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
                    b"borders" => in_borders = false,
                    b"cellXfs" => in_cell_xfs = false,
                    b"numFmts" => in_num_fmts = false,
                    b"border" if in_borders => {
                        in_border = false;
                        sheet.borders.push(std::mem::take(&mut cur_border));
                    }
                    b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_side => {
                        if let Some(name) = cur_side_name.take() {
                            assign_side(&mut cur_border, name.as_bytes(), std::mem::take(&mut cur_side));
                        }
                        in_side = false;
                    }
                    b"xf" if in_xf => {
                        in_xf = false;
                        if let Some(xf) = cur_xf.take() {
                            sheet.cell_xfs.push(xf);
                        }
                    }
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

    pub fn border(&self, style_id: StyleId) -> Option<&Border> {
        let xf = self.cell_xfs.get(style_id.as_usize())?;
        self.borders.get(xf.border_id as usize)
    }

    pub fn alignment(&self, style_id: StyleId) -> Option<&Alignment> {
        let xf = self.cell_xfs.get(style_id.as_usize())?;
        xf.alignment.as_ref()
    }

    pub fn protection(&self, style_id: StyleId) -> Option<&Protection> {
        let xf = self.cell_xfs.get(style_id.as_usize())?;
        xf.protection.as_ref()
    }

    pub fn number_format(&self, style_id: StyleId) -> Option<&str> {
        let num_fmt_id = self.num_fmt_id(style_id)?;
        builtin_num_fmt(num_fmt_id).or_else(|| self.num_fmts.get(&num_fmt_id).map(String::as_str))
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

fn parse_attrs_map(e: &BytesStart<'_>) -> Result<HashMap<String, String>, OxlsxError> {
    let mut attrs = HashMap::new();
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        attrs.insert(
            String::from_utf8_lossy(attr.key.as_ref()).to_string(),
            attr.unescape_value()?.into_owned(),
        );
    }
    Ok(attrs)
}

fn assign_side(border: &mut Border, side_name: &[u8], side: Side) {
    match side_name {
        b"left" => border.left = side,
        b"right" => border.right = side,
        b"top" => border.top = side,
        b"bottom" => border.bottom = side,
        b"diagonal" => border.diagonal = side,
        _ => {}
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

fn parse_bool_like(value: Option<&str>) -> bool {
    matches!(value, Some("1" | "true" | "TRUE" | "True"))
}

fn builtin_num_fmt(num_fmt_id: u32) -> Option<&'static str> {
    match num_fmt_id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        49 => Some("@"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::{Alignment, Border, CellXf, Color, Protection, Side, StyleId, StyleSheet};
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

    #[test]
    fn alignment_default_and_round_trip() {
        let alignment = Alignment::default();
        assert_eq!(alignment.horizontal, None);
        assert_eq!(alignment.vertical, None);
        assert!(!alignment.wrap_text);
        assert_eq!(alignment.indent, 0);

        let attrs = alignment.to_xml_attributes();
        let reparsed = Alignment::from_xml_attributes(&attrs);
        assert_eq!(reparsed, alignment);
    }

    #[test]
    fn side_color_round_trip() {
        let side = Side {
            style: Some("thin".to_string()),
            color: Some(Color::Theme(3)),
        };

        let (style, color) = side.to_xml_parts();
        let reparsed = Side::from_xml_parts(style.as_deref(), color.as_ref());
        assert_eq!(reparsed, side);
    }

    #[test]
    fn border_round_trip_preserves_sides() {
        let border = Border {
            left: Side {
                style: Some("thin".to_string()),
                color: Some(Color::Indexed(4)),
            },
            right: Side {
                style: Some("double".to_string()),
                color: Some(Color::Argb("FFFF0000".to_string())),
            },
            top: Side::default(),
            bottom: Side {
                style: Some("dashed".to_string()),
                color: Some(Color::Theme(2)),
            },
            diagonal: Side::default(),
        };

        let serialized = border.to_xml_sides();
        let reparsed = Border::from_xml_sides(&serialized);
        assert_eq!(reparsed, border);
    }

    #[test]
    fn protection_round_trip_preserves_flags() {
        let protection = Protection {
            locked: false,
            hidden: true,
        };

        let attrs = protection.to_xml_attributes();
        let reparsed = Protection::from_xml_attributes(&attrs);
        assert_eq!(reparsed, protection);
    }

    #[test]
    fn cellxfs_apply_flags_round_trip() {
        let xf = CellXf {
            font_id: 1,
            fill_id: 2,
            num_fmt_id: 164,
            border_id: 3,
            alignment: Some(Alignment {
                horizontal: Some("center".to_string()),
                vertical: Some("top".to_string()),
                wrap_text: true,
                indent: 1,
            }),
            protection: Some(Protection {
                locked: false,
                hidden: true,
            }),
            apply_font: true,
            apply_fill: false,
            apply_border: true,
            apply_number_format: true,
            apply_alignment: true,
            apply_protection: true,
        };

        let attrs = xf.to_xml_attributes();
        let reparsed = CellXf::from_xml_attributes(&attrs, xf.alignment.clone(), xf.protection.clone());
        assert_eq!(reparsed, xf);
    }

    #[test]
    fn parses_border_alignment_and_protection_from_cell_xf() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color theme="4"/></left>
      <right style="double"><color indexed="3"/></right>
      <top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellXfs count="2">
    <xf fontId="0" fillId="0" borderId="0" numFmtId="0" xfId="0"/>
    <xf fontId="0" fillId="0" borderId="1" numFmtId="165" xfId="0" applyBorder="1" applyAlignment="1" applyProtection="1">
      <alignment horizontal="center" vertical="top" wrapText="1" indent="2"/>
      <protection locked="0" hidden="1"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

        let sheet = StyleSheet::parse(Cursor::new(xml.as_bytes())).unwrap();
        let xf = &sheet.cell_xfs[1];
        assert_eq!(xf.border_id, 1);
        assert!(xf.apply_border);
        assert!(xf.apply_alignment);
        assert!(xf.apply_protection);

        let alignment = xf.alignment.as_ref().unwrap();
        assert_eq!(alignment.horizontal.as_deref(), Some("center"));
        assert_eq!(alignment.vertical.as_deref(), Some("top"));
        assert!(alignment.wrap_text);
        assert_eq!(alignment.indent, 2);

        let protection = xf.protection.as_ref().unwrap();
        assert!(!protection.locked);
        assert!(protection.hidden);

        let border = sheet.border(StyleId(1)).unwrap();
        assert_eq!(border.left.style.as_deref(), Some("thin"));
        assert_eq!(border.left.color, Some(Color::Theme(4)));
        assert_eq!(border.right.style.as_deref(), Some("double"));
        assert_eq!(border.right.color, Some(Color::Indexed(3)));
    }
}
