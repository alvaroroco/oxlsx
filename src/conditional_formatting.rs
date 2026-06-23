//! Worksheet conditional-formatting rule model (openpyxl parity).
//!
//! oxlsx models the two openpyxl rule families used by migrations —
//! `FormulaRule` (`cfRule type="expression"`) and `CellIsRule`
//! (`cfRule type="cellIs"` + `operator`) — while round-tripping every other
//! rule type (`colorScale`, `dataBar`, `iconSet`, …) verbatim. Modeled rules
//! carry a [`Dxf`] differential format that is registered into the
//! [`StyleSheet`](crate::StyleSheet) dxf registry at write time; unmodeled
//! rules retain their inner XML as a raw string so they survive read-modify-write.

use crate::styles::Dxf;

/// A single `<cfRule>` inside a `<conditionalFormatting>` block.
///
/// Modeled rules (`expression`, `cellIs`) populate [`dxf`](Self::dxf) and leave
/// [`dxf_id`](Self::dxf_id) `None` until the writer assigns it. Rules parsed from
/// a source workbook populate [`dxf_id`](Self::dxf_id) (and resolve
/// [`dxf`](Self::dxf) from the source registry); unmodeled rule types keep their
/// inner XML in [`unmodeled_children`](Self::unmodeled_children) and any extra
/// attributes in [`unmodeled_attrs`](Self::unmodeled_attrs).
#[derive(Debug, Clone)]
pub struct ConditionalRule {
    /// The `type` attribute, e.g. `"expression"`, `"cellIs"`, `"colorScale"`.
    pub type_ext: String,
    /// Global, sequential priority. `None` until assigned by
    /// [`add_conditional_formatting`](crate::Worksheet::add_conditional_formatting)
    /// or parsed from a source workbook.
    pub priority: Option<u32>,
    /// The `dxfId` parsed from a source workbook, if any.
    pub dxf_id: Option<u32>,
    /// Serializes as `stopIfTrue="1"` when `true`; omitted otherwise.
    pub stop_if_true: bool,
    /// The `operator` attribute (e.g. `"greaterThan"`) for `cellIs` rules.
    pub operator: Option<String>,
    /// `<formula>` operands, in document order (stored unescaped).
    pub formulas: Vec<String>,
    /// The differential format this rule applies. Registered into the dxf
    /// registry at write time to obtain a `dxfId`.
    pub dxf: Option<Dxf>,
    /// Attributes not modeled above, preserved verbatim for round-trip.
    pub unmodeled_attrs: Vec<(String, String)>,
    /// Inner XML of unmodeled child elements (e.g. `<colorScale>…</colorScale>`),
    /// preserved verbatim for round-trip.
    pub unmodeled_children: String,
}

impl ConditionalRule {
    /// Build a FormulaRule (`cfRule type="expression"`) applying `dxf` when the
    /// formula evaluates truthy. Mirrors openpyxl's `FormulaRule`.
    pub fn formula(formulas: Vec<String>, dxf: Dxf) -> Self {
        Self {
            type_ext: "expression".to_string(),
            priority: None,
            dxf_id: None,
            stop_if_true: false,
            operator: None,
            formulas,
            dxf: Some(dxf),
            unmodeled_attrs: Vec::new(),
            unmodeled_children: String::new(),
        }
    }

    /// Build a CellIsRule (`cfRule type="cellIs"`) with the given comparison
    /// `operator` (e.g. `"greaterThan"`). Mirrors openpyxl's `CellIsRule`.
    pub fn cell_is(operator: &str, formulas: Vec<String>, dxf: Dxf) -> Self {
        Self {
            type_ext: "cellIs".to_string(),
            priority: None,
            dxf_id: None,
            stop_if_true: false,
            operator: Some(operator.to_string()),
            formulas,
            dxf: Some(dxf),
            unmodeled_attrs: Vec::new(),
            unmodeled_children: String::new(),
        }
    }

    /// Builder setter for `stopIfTrue`.
    #[must_use]
    pub fn with_stop_if_true(mut self, stop: bool) -> Self {
        self.stop_if_true = stop;
        self
    }

    /// An empty rule used as a parse accumulator.
    pub(crate) fn empty() -> Self {
        Self {
            type_ext: String::new(),
            priority: None,
            dxf_id: None,
            stop_if_true: false,
            operator: None,
            formulas: Vec::new(),
            dxf: None,
            unmodeled_attrs: Vec::new(),
            unmodeled_children: String::new(),
        }
    }
}

/// A `<conditionalFormatting sqref="…">` block: one cell range and its ordered
/// rules. Rules added for the same `sqref` are grouped into one block, matching
/// openpyxl's `ws.conditional_formatting.add`.
#[derive(Debug, Clone)]
pub struct ConditionalFormatting {
    /// The `sqref` (space-separated A1 ranges) this block applies to.
    pub sqref: String,
    /// The rules, in registration (and serialization) order.
    pub rules: Vec<ConditionalRule>,
}
