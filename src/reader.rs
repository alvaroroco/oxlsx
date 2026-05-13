use crate::{Fill, Font, OxlsxError, StyleSheet};
use std::collections::HashMap;
use std::path::Path;
use std::sync::Arc;

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
}

#[derive(Debug, Clone)]
pub struct Cell {
    style_id: u32,
    styles: Arc<StyleSheet>,
}

impl Cell {
    pub fn style_id(&self) -> u32 {
        self.style_id
    }

    pub fn font(&self) -> Option<&Font> {
        self.styles.font(self.style_id)
    }

    pub fn fill(&self) -> Option<&Fill> {
        self.styles.fill(self.style_id)
    }
}

#[derive(Debug, Clone, Default)]
pub struct Worksheet {
    cells: HashMap<String, Cell>,
}

impl Worksheet {
    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        self.cells.get(reference)
    }
}

#[derive(Debug, Clone, Default)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
}

impl Workbook {
    pub fn open<P: AsRef<Path>>(_path: P) -> Result<Self, OxlsxError> {
        unimplemented!("implemented in slice 2")
    }

    pub fn sheet(&self, index: usize) -> Option<&Worksheet> {
        self.sheets.get(index)
    }
}
