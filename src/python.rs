#![allow(unsafe_op_in_unsafe_fn)]
#![allow(clippy::useless_conversion)]

use crate::coords::rc_to_a1;
use crate::{Cell, Color, Fill, Font, OxlsxError, StyleSheet, Value, Workbook};
use chrono::Datelike;
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDate, PyDateAccess, PyModule, PyTuple};
use std::sync::{Arc, Mutex, MutexGuard};

#[pymodule]
fn oxlsx(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCell>()?;
    m.add_class::<PyFont>()?;
    m.add_class::<PyFill>()?;
    m.add_class::<PyColor>()?;

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
            PyDate::new_bound(py, date.year(), date.month() as u8, date.day() as u8)
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

    if let Ok(date) = obj.downcast::<PyDate>() {
        return Ok(Value::Date(
            chrono::NaiveDate::from_ymd_opt(
                date.get_year(),
                date.get_month() as u32,
                date.get_day() as u32,
            )
            .expect("PyDate always represents a valid date"),
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
        other => pyo3::exceptions::PyIOError::new_err(other.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::{py_to_value, value_to_py};
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
}
