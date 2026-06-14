//! Shared quick-xml SAX helpers used across the reader/writer modules.

use crate::OxlsxError;
use quick_xml::events::BytesStart;

/// Read the unescaped value of attribute `key` from a `BytesStart` element.
///
/// Returns `Ok(None)` when the attribute is absent. Attribute checks are
/// disabled (`with_checks(false)`) to match the lenient SAX parsing used
/// throughout the crate.
pub(crate) fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, OxlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| OxlsxError::Xml(err.into()))?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}
