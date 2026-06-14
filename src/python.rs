#![allow(unsafe_op_in_unsafe_fn)]
#![allow(clippy::useless_conversion)]

use crate::coords::rc_to_a1;
use crate::{
    Cell, Color, Comment, DefinedName, DefinedNames, Fill, Font, Hyperlink, OxlsxError, StyleSheet,
    Table, TableColumn, TableStyleInfo, Value, Workbook,
};
use chrono::Datelike;
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyModule, PyTuple};
use std::sync::{Arc, Mutex, MutexGuard};

#[pymodule]
fn oxlsx(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCell>()?;
    m.add_class::<PyComment>()?;
    m.add_class::<PyHyperlink>()?;
    m.add_class::<PyFont>()?;
    m.add_class::<PyFill>()?;
    m.add_class::<PyColor>()?;
    m.add_class::<PyDefinedName>()?;
    m.add_class::<PyDefinedNames>()?;
    m.add_class::<PyTable>()?;
    m.add_class::<PyTableColumn>()?;
    m.add_class::<PyTableStyleInfo>()?;

    Ok(())
}

#[pyfunction]
fn load_workbook(path: &str) -> PyResult<PyWorkbook> {
    let inner = match Workbook::open(path) {
        Ok(workbook) => workbook,
        Err(err) => return Err(map_oxlsx_error(err)),
    };
    Ok(PyWorkbook::from_workbook(inner))
}

#[pyclass(name = "Workbook")]
struct PyWorkbook {
    inner: Arc<Mutex<Workbook>>,
}

impl PyWorkbook {
    fn from_workbook(workbook: Workbook) -> Self {
        Self {
            inner: Arc::new(Mutex::new(workbook)),
        }
    }

    fn lock(&self) -> PyResult<MutexGuard<'_, Workbook>> {
        self.inner
            .lock()
            .map_err(|_| pyo3::exceptions::PyRuntimeError::new_err("Workbook lock poisoned"))
    }
}

#[pymethods]
impl PyWorkbook {
    #[new]
    fn new() -> Self {
        Self::from_workbook(Workbook::new())
    }

    #[getter]
    fn worksheets(&self) -> PyResult<Vec<PyWorksheet>> {
        let workbook = self.lock()?;
        Ok((0..workbook.worksheets().len())
            .map(|sheet_index| PyWorksheet {
                workbook: Arc::clone(&self.inner),
                sheet_index,
            })
            .collect())
    }

    #[getter]
    fn sheetnames(&self) -> PyResult<Vec<String>> {
        let workbook = self.lock()?;
        Ok(workbook
            .sheetnames()
            .into_iter()
            .map(str::to_owned)
            .collect())
    }

    #[getter]
    fn active(&self) -> PyResult<Option<PyWorksheet>> {
        let workbook = self.lock()?;
        let Some(active) = workbook.active() else {
            return Ok(None);
        };

        let active_title = active.title();
        let index = workbook
            .worksheets()
            .iter()
            .position(|sheet| sheet.title() == active_title);

        Ok(index.map(|sheet_index| PyWorksheet {
            workbook: Arc::clone(&self.inner),
            sheet_index,
        }))
    }

    fn __getitem__(&self, title: &str) -> PyResult<PyWorksheet> {
        let workbook = self.lock()?;
        let sheet_index = workbook
            .worksheets()
            .iter()
            .position(|sheet| sheet.title() == title)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(title.to_string()))?;

        Ok(PyWorksheet {
            workbook: Arc::clone(&self.inner),
            sheet_index,
        })
    }

    fn save(&self, path: &str) -> PyResult<()> {
        let workbook = self.lock()?;
        match workbook.save(path) {
            Ok(()) => Ok(()),
            Err(err) => Err(map_oxlsx_error(err)),
        }
    }

    #[getter]
    fn defined_names(&self) -> PyDefinedNames {
        PyDefinedNames {
            workbook: Arc::clone(&self.inner),
            sheet_index: None,
        }
    }
}

/// openpyxl-parity `DefinedName`: a workbook-global named range.
#[pyclass(name = "DefinedName")]
#[derive(Clone)]
struct PyDefinedName {
    inner: DefinedName,
}

#[pymethods]
impl PyDefinedName {
    #[new]
    #[pyo3(signature = (name, value, *, comment=None, hidden=false))]
    fn new(name: &str, value: &str, comment: Option<String>, hidden: bool) -> Self {
        let mut inner = DefinedName::new(name, value);
        inner.comment = comment;
        inner.hidden = hidden;
        Self { inner }
    }

    #[getter]
    fn name(&self) -> &str {
        &self.inner.name
    }

    #[getter]
    fn value(&self) -> &str {
        &self.inner.value
    }

    #[getter]
    fn comment(&self) -> Option<&str> {
        self.inner.comment.as_deref()
    }

    #[getter]
    fn hidden(&self) -> bool {
        self.inner.hidden
    }

    /// Global names are never reserved; mirrors openpyxl returning `None`.
    #[getter]
    fn is_reserved(&self) -> Option<bool> {
        None
    }

    /// Global names have no sheet scope; mirrors openpyxl returning `None`.
    #[getter]
    #[allow(non_snake_case)]
    fn localSheetId(&self) -> Option<usize> {
        None
    }

    fn __repr__(&self) -> String {
        format!(
            "DefinedName(name={:?}, value={:?})",
            self.inner.name, self.inner.value
        )
    }
}

/// Dict-like collection of named ranges, mirroring openpyxl's `DefinedNameDict`.
/// Backed live by the owning workbook. `sheet_index` selects the scope:
/// `None` => workbook-global (`wb.defined_names`); `Some(i)` => the worksheet at
/// index `i` (`ws.defined_names`).
#[pyclass(name = "DefinedNames")]
struct PyDefinedNames {
    workbook: Arc<Mutex<Workbook>>,
    sheet_index: Option<usize>,
}

impl PyDefinedNames {
    fn lock(&self) -> PyResult<MutexGuard<'_, Workbook>> {
        self.workbook
            .lock()
            .map_err(|_| pyo3::exceptions::PyRuntimeError::new_err("Workbook lock poisoned"))
    }

    /// Resolves the scoped collection for read access.
    fn collection<'wb>(&self, workbook: &'wb Workbook) -> PyResult<&'wb DefinedNames> {
        match self.sheet_index {
            None => Ok(workbook.defined_names()),
            Some(index) => workbook
                .sheet(index)
                .map(|sheet| sheet.defined_names())
                .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Worksheet not found")),
        }
    }

    /// Resolves the scoped collection for mutable access.
    fn collection_mut<'wb>(&self, workbook: &'wb mut Workbook) -> PyResult<&'wb mut DefinedNames> {
        match self.sheet_index {
            None => Ok(workbook.defined_names_mut()),
            Some(index) => workbook
                .sheet_mut(index)
                .map(|sheet| sheet.defined_names_mut())
                .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Worksheet not found")),
        }
    }
}

#[pymethods]
impl PyDefinedNames {
    fn add(&self, defined_name: &PyDefinedName) -> PyResult<()> {
        let mut workbook = self.lock()?;
        self.collection_mut(&mut workbook)?
            .add(defined_name.inner.clone())
            .map_err(map_oxlsx_error)
    }

    fn __getitem__(&self, name: &str) -> PyResult<PyDefinedName> {
        let workbook = self.lock()?;
        self.collection(&workbook)?
            .get(name)
            .map(|dn| PyDefinedName { inner: dn.clone() })
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(name.to_string()))
    }

    fn __setitem__(&self, name: &str, defined_name: &PyDefinedName) -> PyResult<()> {
        // openpyxl keys must match the DefinedName.name; enforce the same.
        if name != defined_name.inner.name {
            return Err(pyo3::exceptions::PyValueError::new_err(format!(
                "key {name:?} does not match DefinedName.name {:?}",
                defined_name.inner.name
            )));
        }
        self.add(defined_name)
    }

    fn __delitem__(&self, name: &str) -> PyResult<()> {
        let mut workbook = self.lock()?;
        self.collection_mut(&mut workbook)?
            .remove(name)
            .map(|_| ())
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(name.to_string()))
    }

    fn __contains__(&self, name: &str) -> PyResult<bool> {
        let workbook = self.lock()?;
        Ok(self.collection(&workbook)?.contains(name))
    }

    fn __len__(&self) -> PyResult<usize> {
        let workbook = self.lock()?;
        Ok(self.collection(&workbook)?.len())
    }

    fn __iter__(&self) -> PyResult<PyDefinedNamesIter> {
        let workbook = self.lock()?;
        let names = self
            .collection(&workbook)?
            .names()
            .map(str::to_owned)
            .collect();
        Ok(PyDefinedNamesIter { names, position: 0 })
    }

    /// Returns the list of names, mirroring openpyxl's `DefinedNameDict.keys()`.
    fn keys(&self) -> PyResult<Vec<String>> {
        let workbook = self.lock()?;
        Ok(self
            .collection(&workbook)?
            .names()
            .map(str::to_owned)
            .collect())
    }
}

/// Iterator over named-range keys (snapshot taken at `iter()` time).
#[pyclass]
struct PyDefinedNamesIter {
    names: Vec<String>,
    position: usize,
}

#[pymethods]
impl PyDefinedNamesIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<String> {
        let name = self.names.get(self.position).cloned()?;
        self.position += 1;
        Some(name)
    }
}

#[pyclass(name = "Worksheet")]
struct PyWorksheet {
    workbook: Arc<Mutex<Workbook>>,
    sheet_index: usize,
}

impl PyWorksheet {
    fn lock(&self) -> PyResult<MutexGuard<'_, Workbook>> {
        self.workbook
            .lock()
            .map_err(|_| pyo3::exceptions::PyRuntimeError::new_err("Workbook lock poisoned"))
    }

    fn cell_from_coord(&self, coord: &str) -> PyResult<PyCell> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        let cell = sheet
            .cell(coord)
            .cloned()
            .unwrap_or_else(|| Cell::empty(coord.to_string(), Arc::new(StyleSheet::default())));

        Ok(PyCell { inner: cell })
    }
}

#[pymethods]
impl PyWorksheet {
    #[getter]
    fn title(&self) -> PyResult<String> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        Ok(sheet.title().to_string())
    }

    fn __getitem__(&self, coord: &str) -> PyResult<PyCell> {
        self.cell_from_coord(coord)
    }

    fn __setitem__(&self, coord: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let value = py_to_value(value)?;
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.set_cell(coord, value);

        Ok(())
    }

    #[pyo3(signature = (*, row, column))]
    fn cell(&self, row: u32, column: u32) -> PyResult<PyCell> {
        let coord = rc_to_a1(row, column);
        self.cell_from_coord(&coord)
    }

    #[getter]
    fn max_row(&self) -> PyResult<u32> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        Ok(sheet.max_row())
    }

    #[getter]
    fn max_column(&self) -> PyResult<u32> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        Ok(sheet.max_column())
    }

    #[getter]
    fn min_row(&self) -> PyResult<u32> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        Ok(sheet.min_row())
    }

    #[getter]
    fn min_column(&self) -> PyResult<u32> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        Ok(sheet.min_column())
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    fn iter_rows(
        &self,
        py: Python<'_>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<Vec<Py<PyTuple>>> {
        let rows = {
            let workbook = self.lock()?;
            let sheet = workbook
                .sheet(self.sheet_index)
                .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
            sheet.iter_rows(min_row, max_row, min_col, max_col)
        };

        rows_to_py_tuples(py, rows, values_only)
    }

    #[pyo3(signature = (min_col=None, max_col=None, min_row=None, max_row=None, values_only=false))]
    fn iter_cols(
        &self,
        py: Python<'_>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        values_only: bool,
    ) -> PyResult<Vec<Py<PyTuple>>> {
        let cols = {
            let workbook = self.lock()?;
            let sheet = workbook
                .sheet(self.sheet_index)
                .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
            sheet.iter_cols(min_row, max_row, min_col, max_col)
        };

        rows_to_py_tuples(py, cols, values_only)
    }

    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Vec<Py<PyTuple>>> {
        self.iter_rows(py, None, None, None, None, false)
    }

    #[getter]
    fn columns(&self, py: Python<'_>) -> PyResult<Vec<Py<PyTuple>>> {
        self.iter_cols(py, None, None, None, None, false)
    }

    #[getter]
    fn values(&self, py: Python<'_>) -> PyResult<Vec<Py<PyTuple>>> {
        self.iter_rows(py, None, None, None, None, true)
    }

    fn merge_cells(&self, range: &str) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.merge_cells(range);
        Ok(())
    }

    fn unmerge_cells(&self, range: &str) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.unmerge_cells(range);
        Ok(())
    }

    fn column_width(&self, col: u32, width: f64) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.column_width(col, width);
        Ok(())
    }

    fn row_height(&self, row: u32, height: f64) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.row_height(row, height);
        Ok(())
    }

    fn set_cell_comment(&self, coord: &str, text: &str, author: &str) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.set_cell_comment(
            coord,
            Comment {
                text: text.to_string(),
                author: author.to_string(),
            },
        );
        Ok(())
    }

    /// Attach an external-URL hyperlink to `coord`, mirroring openpyxl's
    /// `cell.hyperlink = "https://..."`. Routed as a worksheet method (like
    /// `set_cell_comment`) because `PyCell` is a value snapshot and cannot
    /// persist a setter back to the workbook. An empty target is a no-op.
    fn set_cell_hyperlink(&self, coord: &str, target: &str) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.set_cell_hyperlink(coord, target);
        Ok(())
    }

    #[getter]
    fn freeze_panes(&self) -> PyResult<Option<String>> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        Ok(sheet.freeze_panes().map(str::to_owned))
    }

    #[setter]
    fn set_freeze_panes(&self, value: Option<String>) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        sheet.set_freeze_panes(value.as_deref());
        Ok(())
    }

    #[getter]
    fn print_area(&self) -> PyResult<Option<String>> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        Ok(sheet.print_area())
    }

    #[setter]
    fn set_print_area(&self, value: Option<String>) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        sheet.set_print_area(value.as_deref());
        Ok(())
    }

    #[getter]
    fn print_title_rows(&self) -> PyResult<Option<String>> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        Ok(sheet.print_title_rows().map(str::to_owned))
    }

    #[setter]
    fn set_print_title_rows(&self, value: Option<String>) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        sheet.set_print_title_rows(value.as_deref());
        Ok(())
    }

    #[getter]
    fn print_title_cols(&self) -> PyResult<Option<String>> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        Ok(sheet.print_title_cols().map(str::to_owned))
    }

    #[setter]
    fn set_print_title_cols(&self, value: Option<String>) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;

        sheet.set_print_title_cols(value.as_deref());
        Ok(())
    }

    /// Sheet-local named ranges, mirroring openpyxl's `ws.defined_names`.
    /// Independent from `wb.defined_names`; the same name may exist in both.
    #[getter]
    fn defined_names(&self) -> PyDefinedNames {
        PyDefinedNames {
            workbook: Arc::clone(&self.workbook),
            sheet_index: Some(self.sheet_index),
        }
    }

    /// Dict-like collection of this sheet's tables, mirroring openpyxl's
    /// `ws.tables` (keyed by `displayName`).
    #[getter]
    fn tables(&self) -> PyTableList {
        PyTableList {
            workbook: Arc::clone(&self.workbook),
            sheet_index: self.sheet_index,
        }
    }

    /// Add an Excel table, mirroring openpyxl's `ws.add_table`.
    ///
    /// Raises `ValueError` if a table with the same `displayName` (case-insensitive)
    /// already exists anywhere in the workbook (checked BEFORE any mutation),
    /// matching openpyxl. Columns are derived from the header row when none are given.
    fn add_table(&self, table: &PyTable) -> PyResult<()> {
        // Extract the table from the DETACHED PyTable.
        // (Attached tables are already in the workbook and cannot be re-added.)
        let table_to_add = match table.backing.as_ref() {
            PyTableBacking::Detached(inner_table) => inner_table.clone(),
            PyTableBacking::Attached { .. } => {
                return Err(pyo3::exceptions::PyValueError::new_err(
                    "Cannot add an attached table; use detached Table(...) constructor",
                ));
            }
        };

        let mut workbook = self.lock()?;

        // FIX 6: Case-insensitive workbook-wide duplicate check.
        let lower = table_to_add.display_name.to_ascii_lowercase();
        let duplicate = workbook
            .worksheets()
            .iter()
            .flat_map(|sheet| sheet.tables())
            .any(|existing| existing.display_name.to_ascii_lowercase() == lower);
        if duplicate {
            return Err(pyo3::exceptions::PyValueError::new_err(format!(
                "Table with name {} already exists",
                table_to_add.display_name
            )));
        }

        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet.add_table(table_to_add).map_err(map_oxlsx_error)
    }
}

#[pyclass(name = "Cell")]
struct PyCell {
    inner: Cell,
}

#[pymethods]
impl PyCell {
    #[getter]
    fn value(&self, py: Python<'_>) -> PyObject {
        value_to_py(self.inner.value(), py)
    }

    #[getter]
    fn font(&self) -> Option<PyFont> {
        self.inner
            .font()
            .cloned()
            .or_else(|| self.inner.write_font.clone())
            .map(|inner| PyFont { inner })
    }

    #[getter]
    fn fill(&self) -> Option<PyFill> {
        self.inner
            .fill()
            .cloned()
            .or_else(|| self.inner.write_fill.clone())
            .map(|inner| PyFill { inner })
    }

    #[getter]
    fn comment(&self) -> Option<PyComment> {
        self.inner
            .comment()
            .cloned()
            .map(|inner| PyComment { inner })
    }

    #[getter]
    fn hyperlink(&self) -> Option<PyHyperlink> {
        self.inner.hyperlink().map(PyHyperlink::from)
    }
}

#[pyclass(name = "Comment")]
struct PyComment {
    inner: Comment,
}

#[pymethods]
impl PyComment {
    #[getter]
    fn text(&self) -> String {
        self.inner.text.clone()
    }

    #[getter]
    fn author(&self) -> String {
        self.inner.author.clone()
    }
}

/// openpyxl-parity `Hyperlink`: a value snapshot of a cell's hyperlink,
/// exposing `target`/`location`/`tooltip`/`display`.
///
/// Mirrors `openpyxl.worksheet.hyperlink.Hyperlink`: `target` is `None` for an
/// internal `location`-only link (Rust stores it as an empty string), so the
/// getter normalizes empty to `None` to match openpyxl exactly.
#[pyclass(name = "Hyperlink")]
struct PyHyperlink {
    inner: Hyperlink,
}

impl From<&Hyperlink> for PyHyperlink {
    fn from(hyperlink: &Hyperlink) -> Self {
        Self {
            inner: hyperlink.clone(),
        }
    }
}

#[pymethods]
impl PyHyperlink {
    #[getter]
    fn target(&self) -> Option<String> {
        if self.inner.target.is_empty() {
            None
        } else {
            Some(self.inner.target.clone())
        }
    }

    #[getter]
    fn location(&self) -> Option<String> {
        self.inner.location.clone()
    }

    #[getter]
    fn tooltip(&self) -> Option<String> {
        self.inner.tooltip.clone()
    }

    #[getter]
    fn display(&self) -> Option<String> {
        self.inner.display.clone()
    }
}

/// openpyxl-parity `TableStyleInfo`: built-in style name + four banding flags.
#[pyclass(name = "TableStyleInfo")]
#[derive(Clone)]
struct PyTableStyleInfo {
    inner: TableStyleInfo,
}

#[pymethods]
impl PyTableStyleInfo {
    #[new]
    #[pyo3(signature = (
        name,
        *,
        showFirstColumn = false,
        showLastColumn = false,
        showRowStripes = false,
        showColumnStripes = false
    ))]
    #[allow(non_snake_case)]
    fn new(
        name: &str,
        showFirstColumn: bool,
        showLastColumn: bool,
        showRowStripes: bool,
        showColumnStripes: bool,
    ) -> Self {
        Self {
            inner: TableStyleInfo {
                name: name.to_string(),
                show_first_column: showFirstColumn,
                show_last_column: showLastColumn,
                show_row_stripes: showRowStripes,
                show_column_stripes: showColumnStripes,
            },
        }
    }

    #[getter]
    fn name(&self) -> &str {
        &self.inner.name
    }

    #[getter]
    #[allow(non_snake_case)]
    fn showFirstColumn(&self) -> bool {
        self.inner.show_first_column
    }

    #[getter]
    #[allow(non_snake_case)]
    fn showLastColumn(&self) -> bool {
        self.inner.show_last_column
    }

    #[getter]
    #[allow(non_snake_case)]
    fn showRowStripes(&self) -> bool {
        self.inner.show_row_stripes
    }

    #[getter]
    #[allow(non_snake_case)]
    fn showColumnStripes(&self) -> bool {
        self.inner.show_column_stripes
    }
}

/// openpyxl-parity `TableColumn`: read-only id + name.
#[pyclass(name = "TableColumn")]
#[derive(Clone)]
struct PyTableColumn {
    inner: TableColumn,
}

#[pymethods]
impl PyTableColumn {
    #[getter]
    fn id(&self) -> u32 {
        self.inner.id
    }

    #[getter]
    fn name(&self) -> &str {
        &self.inner.name
    }
}

/// Backing storage mode for PyTable: either detached (owns a Table) or
/// attached (backed by a live workbook).
#[derive(Clone)]
enum PyTableBacking {
    /// Detached mode: owns the table (used for the standalone `Table(...)` constructor
    /// before `add_table`).
    Detached(Table),
    /// Attached mode: backed by the workbook + sheet + display_name identity
    /// (returned by `ws.tables["X"]`).
    Attached {
        workbook: Arc<Mutex<Workbook>>,
        sheet_index: usize,
        display_name: String,
    },
}

/// openpyxl-parity `Table`: construct as `Table(displayName=..., ref=...)`.
///
/// Supports dual mode: detached (constructed via `Table(...)`) or attached
/// (returned by `ws.tables["X"]`). Attached tables mutate the live workbook.
#[pyclass(name = "Table")]
struct PyTable {
    backing: Box<PyTableBacking>,
}

impl PyTable {
    /// Helper to access the table (read-only) regardless of backing mode.
    fn with_table<F, R>(&self, f: F) -> PyResult<R>
    where
        F: FnOnce(&Table) -> R,
    {
        match self.backing.as_ref() {
            PyTableBacking::Detached(table) => Ok(f(table)),
            PyTableBacking::Attached {
                workbook,
                sheet_index,
                display_name,
            } => {
                let wb = workbook.lock().map_err(|_| {
                    pyo3::exceptions::PyRuntimeError::new_err("Workbook lock poisoned")
                })?;
                let sheet = wb
                    .sheet(*sheet_index)
                    .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
                let table = sheet
                    .tables()
                    .iter()
                    .find(|t| t.display_name == *display_name)
                    .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(display_name.clone()))?;
                Ok(f(table))
            }
        }
    }

    /// Helper to mutate the table regardless of backing mode.
    fn with_table_mut<F, R>(&mut self, f: F) -> PyResult<R>
    where
        F: FnOnce(&mut Table) -> R,
    {
        match self.backing.as_mut() {
            PyTableBacking::Detached(table) => Ok(f(table)),
            PyTableBacking::Attached {
                workbook,
                sheet_index,
                display_name,
            } => {
                let mut wb = workbook.lock().map_err(|_| {
                    pyo3::exceptions::PyRuntimeError::new_err("Workbook lock poisoned")
                })?;
                let sheet = wb
                    .sheet_mut(*sheet_index)
                    .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
                let table = sheet
                    .table_mut(display_name)
                    .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(display_name.clone()))?;
                Ok(f(table))
            }
        }
    }
}

#[pymethods]
impl PyTable {
    #[new]
    #[allow(non_snake_case)]
    fn new(displayName: &str, r#ref: &str) -> Self {
        Self {
            backing: Box::new(PyTableBacking::Detached(Table::new(displayName, r#ref))),
        }
    }

    #[getter]
    #[allow(non_snake_case)]
    fn displayName(&self) -> PyResult<String> {
        self.with_table(|table| table.display_name.clone())
    }

    #[getter]
    fn name(&self) -> PyResult<String> {
        self.with_table(|table| table.name.clone())
    }

    #[getter]
    #[pyo3(name = "ref")]
    fn ref_range(&self) -> PyResult<String> {
        self.with_table(|table| table.ref_range.clone())
    }

    #[getter]
    #[allow(non_snake_case)]
    fn tableColumns(&self) -> PyResult<Vec<PyTableColumn>> {
        self.with_table(|table| {
            table
                .columns
                .iter()
                .cloned()
                .map(|inner| PyTableColumn { inner })
                .collect()
        })
    }

    #[getter]
    #[allow(non_snake_case)]
    fn tableStyleInfo(&self) -> PyResult<Option<PyTableStyleInfo>> {
        self.with_table(|table| table.style.clone().map(|inner| PyTableStyleInfo { inner }))
    }

    #[setter]
    #[allow(non_snake_case)]
    fn set_tableStyleInfo(&mut self, value: Option<PyTableStyleInfo>) -> PyResult<()> {
        self.with_table_mut(|table| {
            table.style = value.map(|style| style.inner);
        })
    }
}

/// Dict-like collection of a worksheet's tables, mirroring openpyxl's
/// `TableList` (keyed by `displayName`). Backed live by the owning workbook.
#[pyclass(name = "TableList")]
struct PyTableList {
    workbook: Arc<Mutex<Workbook>>,
    sheet_index: usize,
}

impl PyTableList {
    fn lock(&self) -> PyResult<MutexGuard<'_, Workbook>> {
        self.workbook
            .lock()
            .map_err(|_| pyo3::exceptions::PyRuntimeError::new_err("Workbook lock poisoned"))
    }

    fn display_names(&self) -> PyResult<Vec<String>> {
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        Ok(sheet
            .tables()
            .iter()
            .map(|table| table.display_name.clone())
            .collect())
    }
}

#[pymethods]
impl PyTableList {
    fn __getitem__(&self, name: &str) -> PyResult<PyTable> {
        // Verify the table exists before returning an attached PyTable.
        let workbook = self.lock()?;
        let sheet = workbook
            .sheet(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        let table_exists = sheet
            .tables()
            .iter()
            .any(|table| table.display_name == name);

        if !table_exists {
            return Err(pyo3::exceptions::PyKeyError::new_err(name.to_string()));
        }

        // Return an ATTACHED PyTable backed by the live workbook.
        Ok(PyTable {
            backing: Box::new(PyTableBacking::Attached {
                workbook: Arc::clone(&self.workbook),
                sheet_index: self.sheet_index,
                display_name: name.to_string(),
            }),
        })
    }

    /// Removes a table by `displayName`, mirroring openpyxl's
    /// `del ws.tables["X"]` (its `TableList` subclasses `dict`). Routes to the
    /// Rust `Worksheet::remove_table`, dropping the part, its rel, its
    /// `<tablePart>`, and its retained verbatim bytes on the next save. Raises
    /// `KeyError` when no table with that name exists.
    fn __delitem__(&self, name: &str) -> PyResult<()> {
        let mut workbook = self.lock()?;
        let sheet = workbook
            .sheet_mut(self.sheet_index)
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err("Worksheet not found"))?;
        sheet
            .remove_table(name)
            .map(|_| ())
            .ok_or_else(|| pyo3::exceptions::PyKeyError::new_err(name.to_string()))
    }

    fn __contains__(&self, name: &str) -> PyResult<bool> {
        Ok(self.display_names()?.iter().any(|item| item == name))
    }

    fn __len__(&self) -> PyResult<usize> {
        Ok(self.display_names()?.len())
    }

    fn __iter__(&self) -> PyResult<PyTableNamesIter> {
        Ok(PyTableNamesIter {
            names: self.display_names()?,
            position: 0,
        })
    }

    /// Mirrors openpyxl's `TableList.keys()` (the `displayName`s).
    fn keys(&self) -> PyResult<Vec<String>> {
        self.display_names()
    }
}

/// Iterator over table `displayName`s (snapshot taken at `iter()` time).
#[pyclass]
struct PyTableNamesIter {
    names: Vec<String>,
    position: usize,
}

#[pymethods]
impl PyTableNamesIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<String> {
        let name = self.names.get(self.position).cloned()?;
        self.position += 1;
        Some(name)
    }
}

#[pyclass(name = "Font")]
struct PyFont {
    inner: Font,
}

#[pymethods]
impl PyFont {
    #[getter]
    fn bold(&self) -> bool {
        self.inner.bold
    }

    #[getter]
    fn italic(&self) -> bool {
        self.inner.italic
    }
}

#[pyclass(name = "Fill")]
struct PyFill {
    inner: Fill,
}

#[pymethods]
impl PyFill {
    #[getter]
    fn fg_color(&self) -> Option<PyColor> {
        self.inner.fg_color.clone().map(|inner| PyColor { inner })
    }
}

#[pyclass(name = "Color")]
struct PyColor {
    inner: Color,
}

#[pymethods]
impl PyColor {
    #[getter]
    fn rgb(&self) -> Option<String> {
        match &self.inner {
            Color::Argb(value) => Some(value.clone()),
            _ => None,
        }
    }

    #[getter]
    fn theme(&self) -> Option<u32> {
        match self.inner {
            Color::Theme(value) => Some(value),
            _ => None,
        }
    }

    #[getter]
    fn indexed(&self) -> Option<u32> {
        match self.inner {
            Color::Indexed(value) => Some(value),
            _ => None,
        }
    }
}

fn rows_to_py_tuples(
    py: Python<'_>,
    rows: Vec<Vec<Cell>>,
    values_only: bool,
) -> PyResult<Vec<Py<PyTuple>>> {
    let mut tuples = Vec::with_capacity(rows.len());

    for row in rows {
        let mut items: Vec<PyObject> = Vec::with_capacity(row.len());
        if values_only {
            for cell in row {
                items.push(value_to_py(cell.value(), py));
            }
        } else {
            for cell in row {
                let py_cell = Py::new(py, PyCell { inner: cell })?;
                items.push(py_cell.into_py(py));
            }
        }
        let tuple = PyTuple::new_bound(py, items).unbind();
        tuples.push(tuple);
    }

    Ok(tuples)
}

fn value_to_py(value: &Value, py: Python<'_>) -> PyObject {
    match value {
        Value::Empty => py.None(),
        Value::String(s) => s.into_py(py),
        Value::Number(n)
            if n.is_finite()
                && n.fract() == 0.0
                && *n >= i64::MIN as f64
                && *n <= i64::MAX as f64 =>
        {
            (*n as i64).into_py(py)
        }
        Value::Number(n) => n.into_py(py),
        Value::Bool(b) => b.into_py(py),
        Value::Date(date) => {
            // Build the date through the `datetime` module's object protocol
            // rather than the `PyDate` C-API. The C-API datetime types are not
            // exposed under the limited API (abi3), and this path compiles
            // identically for both the interpreter-linked and abi3 builds while
            // still producing a genuine `datetime.date` instance.
            let datetime = PyModule::import_bound(py, "datetime")
                .expect("datetime module is always available");
            datetime
                .getattr("date")
                .expect("datetime.date always exists")
                .call1((date.year(), date.month(), date.day()))
                .expect("NaiveDate is always a valid Python date")
                .into_py(py)
        }
        Value::Formula(formula) => format!("={formula}").into_py(py),
    }
}

fn py_to_value(obj: &Bound<'_, PyAny>) -> PyResult<Value> {
    if obj.is_none() {
        return Ok(Value::Empty);
    }

    if let Ok(value) = obj.extract::<bool>() {
        return Ok(Value::Bool(value));
    }

    if let Ok(value) = obj.extract::<i64>() {
        return Ok(Value::Number(value as f64));
    }

    if let Ok(value) = obj.extract::<f64>() {
        return Ok(Value::Number(value));
    }

    if let Ok(value) = obj.extract::<String>() {
        if let Some(formula) = value.strip_prefix('=') {
            return Ok(Value::Formula(formula.to_string()));
        }

        return Ok(Value::String(value));
    }

    // Match `datetime.date` (and its `datetime.datetime` subclass) via the
    // object protocol instead of downcasting to the `PyDate` C-API type, which
    // is unavailable under the limited API (abi3). `datetime.datetime` is a
    // subclass of `datetime.date`, so `is_instance` preserves the prior
    // `downcast::<PyDate>()` behaviour for both types.
    let datetime = PyModule::import_bound(obj.py(), "datetime")?;
    let date_cls = datetime.getattr("date")?;
    if obj.is_instance(&date_cls)? {
        let year: i32 = obj.getattr("year")?.extract()?;
        let month: u32 = obj.getattr("month")?.extract()?;
        let day: u32 = obj.getattr("day")?.extract()?;
        return Ok(Value::Date(
            chrono::NaiveDate::from_ymd_opt(year, month, day)
                .expect("datetime.date always represents a valid date"),
        ));
    }

    Err(pyo3::exceptions::PyTypeError::new_err(
        "Unsupported Python value type for oxlsx.Value",
    ))
}

fn map_oxlsx_error(err: OxlsxError) -> PyErr {
    match err {
        OxlsxError::InvalidFormat(message) | OxlsxError::MissingPart(message) => {
            pyo3::exceptions::PyValueError::new_err(message)
        }
        OxlsxError::DuplicateTableName(_) => {
            pyo3::exceptions::PyValueError::new_err(err.to_string())
        }
        other => pyo3::exceptions::PyIOError::new_err(other.to_string()),
    }
}

// These tests exercise the PyO3 bindings through a real, interpreter-linked
// build (the `python` feature enables `auto-initialize`). They use the `PyDate`
// C-API directly, which is unavailable under the limited API (abi3), so they
// are gated to the `python` feature and never compiled in the abi3 wheel build.
#[cfg(all(test, feature = "python"))]
mod tests {
    use super::{PyCell, PyWorkbook, py_to_value, value_to_py};
    use crate::Value;
    use chrono::NaiveDate;
    use pyo3::prelude::*;
    use pyo3::types::{PyDate, PyDateAccess};

    #[test]
    fn value_to_py_maps_supported_variants() {
        Python::with_gil(|py| {
            let empty = value_to_py(&Value::Empty, py);
            assert!(empty.bind(py).is_none());

            let string = value_to_py(&Value::String("Hola".to_string()), py);
            assert_eq!(string.bind(py).extract::<String>().unwrap(), "Hola");

            let int_like = value_to_py(&Value::Number(42.0), py);
            assert_eq!(int_like.bind(py).extract::<i64>().unwrap(), 42);

            let float_like = value_to_py(&Value::Number(42.5), py);
            assert_eq!(float_like.bind(py).extract::<f64>().unwrap(), 42.5);

            let bool_value = value_to_py(&Value::Bool(true), py);
            assert!(bool_value.bind(py).extract::<bool>().unwrap());

            let date = NaiveDate::from_ymd_opt(2025, 1, 21).unwrap();
            let py_date = value_to_py(&Value::Date(date), py);
            let py_date = py_date.bind(py).downcast::<PyDate>().unwrap();
            assert_eq!(py_date.get_year(), 2025);
            assert_eq!(py_date.get_month(), 1);
            assert_eq!(py_date.get_day(), 21);

            let formula = value_to_py(&Value::Formula("SUM(B1:B3)".to_string()), py);
            assert_eq!(formula.bind(py).extract::<String>().unwrap(), "=SUM(B1:B3)");
        });
    }

    #[test]
    fn py_to_value_maps_supported_python_types() {
        Python::with_gil(|py| {
            let none_obj = py.None();
            assert_eq!(py_to_value(none_obj.bind(py)).unwrap(), Value::Empty);

            let bool_obj = true.into_py(py);
            assert_eq!(py_to_value(bool_obj.bind(py)).unwrap(), Value::Bool(true));

            let int_obj = 42i64.into_py(py);
            assert_eq!(py_to_value(int_obj.bind(py)).unwrap(), Value::Number(42.0));

            let float_obj = 42.5f64.into_py(py);
            assert_eq!(
                py_to_value(float_obj.bind(py)).unwrap(),
                Value::Number(42.5)
            );

            let str_obj: PyObject = "Hola".into_py(py);
            assert_eq!(
                py_to_value(str_obj.bind(py)).unwrap(),
                Value::String("Hola".to_string())
            );

            let formula_obj: PyObject = "=SUM(B1:B3)".into_py(py);
            assert_eq!(
                py_to_value(formula_obj.bind(py)).unwrap(),
                Value::Formula("SUM(B1:B3)".to_string())
            );

            let date_obj = PyDate::new_bound(py, 2025, 1, 21).unwrap();
            assert_eq!(
                py_to_value(&date_obj.into_any()).unwrap(),
                Value::Date(NaiveDate::from_ymd_opt(2025, 1, 21).unwrap())
            );
        });
    }

    #[test]
    fn py_cell_comment_getter_exposes_text_and_author() {
        let mut workbook = crate::Workbook::new();
        workbook.active_mut().unwrap().set_cell_comment(
            "A1",
            crate::Comment {
                text: "Hello".to_string(),
                author: "Alice".to_string(),
            },
        );

        let cell = workbook.active().unwrap().cell("A1").unwrap().clone();
        let py_cell = PyCell { inner: cell };
        let comment = py_cell.comment().expect("comment should exist");

        assert_eq!(comment.text(), "Hello");
        assert_eq!(comment.author(), "Alice");
    }

    #[test]
    fn py_worksheet_set_cell_comment_round_trips_through_workbook_view() {
        let workbook = PyWorkbook::new();
        let worksheet = workbook
            .active()
            .unwrap()
            .expect("workbook should have active sheet");

        worksheet
            .set_cell_comment("C3", "Created from rust test", "Py Tester")
            .unwrap();
        let cell = worksheet.__getitem__("C3").unwrap();
        let comment = cell.comment().expect("comment should exist");

        assert_eq!(comment.text(), "Created from rust test");
        assert_eq!(comment.author(), "Py Tester");
    }

    #[test]
    fn py_worksheet_set_cell_hyperlink_round_trips_through_workbook_view() {
        let workbook = PyWorkbook::new();
        let worksheet = workbook
            .active()
            .unwrap()
            .expect("workbook should have active sheet");

        worksheet
            .set_cell_hyperlink("A1", "https://example.com")
            .unwrap();
        let cell = worksheet.__getitem__("A1").unwrap();
        let link = cell.hyperlink().expect("hyperlink should exist");

        assert_eq!(link.target(), Some("https://example.com".to_string()));
    }

    #[test]
    fn py_hyperlink_normalizes_empty_target_to_none() {
        // An internal location-only link stores an empty target; openpyxl
        // surfaces that as `target is None`.
        let py_link = super::PyHyperlink::from(&crate::Hyperlink {
            target: String::new(),
            location: Some("Sheet!A1".to_string()),
            tooltip: None,
            display: Some("Go to A1".to_string()),
        });
        assert_eq!(py_link.target(), None);
        assert_eq!(py_link.location(), Some("Sheet!A1".to_string()));
        assert_eq!(py_link.display(), Some("Go to A1".to_string()));
    }

    #[test]
    fn test_freeze_panes_round_trip() {
        let path = std::env::temp_dir().join("oxlsx_py_test_freeze_round_trip.xlsx");
        let mut workbook = crate::Workbook::new();
        workbook.active_mut().unwrap().set_freeze_panes(Some("B2"));
        workbook.save(&path).unwrap();

        let py_workbook = PyWorkbook::from_workbook(crate::Workbook::open(&path).unwrap());
        let worksheet = py_workbook
            .active()
            .unwrap()
            .expect("workbook should have active sheet");

        assert_eq!(worksheet.freeze_panes().unwrap(), Some("B2".to_string()));
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_freeze_panes_write_new_workbook() {
        let path = std::env::temp_dir().join("oxlsx_py_test_freeze_write_new.xlsx");
        let workbook = PyWorkbook::new();
        let worksheet = workbook
            .active()
            .unwrap()
            .expect("workbook should have active sheet");

        worksheet.set_freeze_panes(Some("B2".to_string())).unwrap();
        workbook.save(path.to_str().unwrap()).unwrap();

        let reopened = crate::Workbook::open(&path).unwrap();
        assert_eq!(reopened.active().unwrap().freeze_panes(), Some("B2"));
        std::fs::remove_file(path).ok();
    }

    #[test]
    fn test_freeze_panes_clear_removes_pane() {
        let path = std::env::temp_dir().join("oxlsx_py_test_freeze_clear.xlsx");
        let mut workbook = crate::Workbook::new();
        workbook.active_mut().unwrap().set_freeze_panes(Some("B2"));
        let py_workbook = PyWorkbook::from_workbook(workbook);
        let worksheet = py_workbook
            .active()
            .unwrap()
            .expect("workbook should have active sheet");

        worksheet.set_freeze_panes(None).unwrap();
        py_workbook.save(path.to_str().unwrap()).unwrap();

        let reopened = crate::Workbook::open(&path).unwrap();
        assert_eq!(reopened.active().unwrap().freeze_panes(), None);
        std::fs::remove_file(path).ok();
    }
}
