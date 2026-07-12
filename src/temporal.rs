use crate::{OxlsxError, Value};
use chrono::{Datelike, Duration, NaiveDate, NaiveTime, Timelike};

pub(crate) const DATE_FORMAT: &str = "yyyy-mm-dd";
pub(crate) const DATETIME_FORMAT: &str = "yyyy-mm-dd h:mm:ss";
pub(crate) const TIME_FORMAT: &str = "h:mm:ss";
pub(crate) const DURATION_FORMAT: &str = "[hh]:mm:ss";

const MILLIS_PER_DAY: i64 = 86_400_000;
const NANOS_PER_MILLI: i64 = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TemporalKind {
    Date,
    DateTime,
    Time,
    Duration,
}

impl TemporalKind {
    pub(crate) fn default_format(self) -> &'static str {
        match self {
            Self::Date => DATE_FORMAT,
            Self::DateTime => DATETIME_FORMAT,
            Self::Time => TIME_FORMAT,
            Self::Duration => DURATION_FORMAT,
        }
    }
}

pub(crate) fn value_kind(value: &Value) -> Option<TemporalKind> {
    match value {
        Value::Date(_) => Some(TemporalKind::Date),
        Value::DateTime(_) => Some(TemporalKind::DateTime),
        Value::Time(_) => Some(TemporalKind::Time),
        Value::Duration(_) => Some(TemporalKind::Duration),
        _ => None,
    }
}

pub(crate) fn format_kind(format: &str) -> Option<TemporalKind> {
    let normalized = strip_format_literals(first_format_section(format));
    if normalized.contains("[h]")
        || normalized.contains("[hh]")
        || normalized.contains("[m]")
        || normalized.contains("[mm]")
        || normalized.contains("[s]")
        || normalized.contains("[ss]")
    {
        return Some(TemporalKind::Duration);
    }
    let has_date = normalized.contains('y') || normalized.contains('d');
    let has_time = normalized.contains('h') || normalized.contains('s');
    match (has_date, has_time) {
        (true, true) => Some(TemporalKind::DateTime),
        (true, false) => Some(TemporalKind::Date),
        (false, true) => Some(TemporalKind::Time),
        (false, false) => None,
    }
}

fn first_format_section(format: &str) -> &str {
    let mut quoted = false;
    let mut bracketed = false;
    let mut escaped = false;
    for (index, ch) in format.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        match ch {
            '\\' | '_' | '*' => escaped = true,
            '"' if !bracketed => quoted = !quoted,
            '[' if !quoted => bracketed = true,
            ']' if !quoted => bracketed = false,
            ';' if !quoted && !bracketed => return &format[..index],
            _ => {}
        }
    }
    format
}

fn strip_format_literals(format: &str) -> String {
    let lowercase = format.to_ascii_lowercase();
    let mut chars = lowercase.chars().peekable();
    let mut normalized = String::with_capacity(format.len());
    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                for literal in chars.by_ref() {
                    if literal == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                let _ = chars.next();
            }
            '[' => {
                let mut bracket = String::new();
                for part in chars.by_ref() {
                    if part == ']' {
                        break;
                    }
                    bracket.push(part);
                }
                if matches!(bracket.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
                    normalized.push('[');
                    normalized.push_str(&bracket);
                    normalized.push(']');
                }
            }
            _ => normalized.push(ch),
        }
    }
    normalized
}

pub(crate) fn to_excel_serial(value: &Value) -> Result<f64, OxlsxError> {
    match value {
        Value::Date(date) => {
            validate_date(*date)?;
            Ok(date_days(*date) as f64)
        }
        Value::DateTime(datetime) => {
            validate_date(datetime.date())?;
            validate_time(datetime.time())?;
            let epoch = excel_epoch().and_hms_opt(0, 0, 0).expect("valid epoch");
            let millis = round_duration_to_millis(datetime.signed_duration_since(epoch))?;
            let normalized = epoch
                .checked_add_signed(Duration::milliseconds(millis))
                .ok_or_else(|| temporal_range_error("datetime"))?;
            validate_date(normalized.date())?;
            exact_serial_from_millis(adjust_early_date_millis(millis))
        }
        Value::Time(time) => {
            validate_time(*time)?;
            let millis = round_duration_to_millis(Duration::nanoseconds(
                i64::from(time.num_seconds_from_midnight()) * 1_000_000_000
                    + i64::from(time.nanosecond()),
            ))?;
            exact_serial_from_millis(millis.rem_euclid(MILLIS_PER_DAY))
        }
        Value::Duration(duration) => {
            let millis = round_duration_to_millis(*duration)?;
            exact_serial_from_millis(millis)
        }
        _ => Err(OxlsxError::InvalidFormat(
            "value is not temporal".to_string(),
        )),
    }
}

fn exact_serial_from_millis(millis: i64) -> Result<f64, OxlsxError> {
    let serial = millis as f64 / MILLIS_PER_DAY as f64;
    let recovered = (serial * MILLIS_PER_DAY as f64).round_ties_even();
    // `i64::MAX as f64` rounds UP to 2^63. Rust's float-to-int cast saturates
    // that value back to i64::MAX, which can hide a mismatch at Duration::MAX.
    // Reject the exclusive upper boundary before any cast; -2^63 is exact and
    // remains a valid inclusive lower boundary.
    const I64_MIN_F64: f64 = -9_223_372_036_854_775_808.0;
    const I64_MAX_EXCLUSIVE_F64: f64 = 9_223_372_036_854_775_808.0;
    if !recovered.is_finite()
        || !(I64_MIN_F64..I64_MAX_EXCLUSIVE_F64).contains(&recovered)
        || recovered as i64 != millis
    {
        return Err(OxlsxError::InvalidFormat(
            "temporal value cannot be represented exactly at XLSX millisecond precision"
                .to_string(),
        ));
    }
    Ok(serial)
}

fn validate_time(time: NaiveTime) -> Result<(), OxlsxError> {
    if time.nanosecond() >= 1_000_000_000 {
        return Err(OxlsxError::InvalidFormat(
            "Excel cannot represent leap-second temporal values".to_string(),
        ));
    }
    Ok(())
}

fn temporal_range_error(kind: &str) -> OxlsxError {
    OxlsxError::InvalidFormat(format!("{kind} is outside the XLSX temporal range"))
}

fn validate_date(date: NaiveDate) -> Result<(), OxlsxError> {
    if !(1..=9_999).contains(&date.year()) {
        return Err(OxlsxError::InvalidFormat(format!(
            "temporal year {} is outside the XLSX range 1..=9999",
            date.year()
        )));
    }
    Ok(())
}

pub(crate) fn from_excel_serial(serial: f64, kind: TemporalKind) -> Option<Value> {
    if !serial.is_finite() {
        return None;
    }
    let scaled = serial * MILLIS_PER_DAY as f64;
    if scaled < i64::MIN as f64 || scaled > i64::MAX as f64 {
        return None;
    }
    let mut millis = scaled.round_ties_even() as i64;
    match kind {
        TemporalKind::Date => {
            let mut days = serial.floor() as i64;
            if (1..60).contains(&days) {
                days += 1;
            }
            let date = excel_epoch().checked_add_signed(Duration::days(days))?;
            validate_date(date).ok()?;
            Some(Value::Date(date))
        }
        TemporalKind::DateTime => {
            if (MILLIS_PER_DAY..60 * MILLIS_PER_DAY).contains(&millis) {
                millis += MILLIS_PER_DAY;
            }
            let datetime = excel_epoch()
                .and_hms_opt(0, 0, 0)?
                .checked_add_signed(Duration::milliseconds(millis))?;
            validate_date(datetime.date()).ok()?;
            Some(Value::DateTime(datetime))
        }
        TemporalKind::Time if !(0.0..1.0).contains(&serial) => Some(Value::Number(serial)),
        TemporalKind::Time => Some(Value::Time(
            NaiveTime::from_hms_milli_opt(0, 0, 0, 0)?
                .overflowing_add_signed(Duration::milliseconds(millis))
                .0,
        )),
        TemporalKind::Duration => Some(Value::Duration(Duration::milliseconds(millis))),
    }
}

fn excel_epoch() -> NaiveDate {
    NaiveDate::from_ymd_opt(1899, 12, 30).expect("valid Excel epoch")
}

fn date_days(date: NaiveDate) -> i64 {
    let mut days = date.signed_duration_since(excel_epoch()).num_days();
    if (1..=60).contains(&days) {
        days -= 1;
    }
    days
}

fn adjust_early_date_millis(millis: i64) -> i64 {
    if (MILLIS_PER_DAY..=60 * MILLIS_PER_DAY).contains(&millis) {
        millis - MILLIS_PER_DAY
    } else {
        millis
    }
}

pub(crate) fn round_duration_to_millis(duration: Duration) -> Result<i64, OxlsxError> {
    let seconds = duration.num_seconds();
    let remainder = duration - Duration::seconds(seconds);
    let nanos = remainder.num_nanoseconds().ok_or_else(|| {
        OxlsxError::InvalidFormat("temporal value is outside XLSX range".to_string())
    })?;
    let base = seconds.checked_mul(1_000).ok_or_else(|| {
        OxlsxError::InvalidFormat("temporal value is outside XLSX range".to_string())
    })?;
    let quotient = nanos / NANOS_PER_MILLI;
    let residual = nanos % NANOS_PER_MILLI;
    let abs_residual = residual.unsigned_abs();
    let increment = if abs_residual > 500_000
        || (abs_residual == 500_000 && quotient.unsigned_abs() % 2 == 1)
    {
        residual.signum()
    } else {
        0
    };
    base.checked_add(quotient)
        .and_then(|value| value.checked_add(increment))
        .ok_or_else(|| {
            OxlsxError::InvalidFormat("temporal value is outside XLSX range".to_string())
        })
}

#[cfg(test)]
mod tests {
    use super::{
        TemporalKind, exact_serial_from_millis, format_kind, from_excel_serial, to_excel_serial,
    };
    use crate::Value;
    use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

    #[test]
    fn format_classification_ignores_colors_conditions_and_quoted_literals() {
        assert_eq!(format_kind("[Red]0.00"), None);
        assert_eq!(format_kind("[>=100]0;[Red]0"), None);
        assert_eq!(
            format_kind(r#"yyyy-mm-dd "days""#),
            Some(TemporalKind::Date)
        );
    }

    #[test]
    fn format_classification_uses_first_section_only() {
        assert_eq!(format_kind("0;yyyy-mm-dd"), None);
        assert_eq!(format_kind("yyyy-mm-dd;0"), Some(TemporalKind::Date));
    }

    #[test]
    fn rejects_lossy_large_duration_and_post_rounding_datetime_overflow() {
        let lossy = Value::Duration(Duration::milliseconds((1_i64 << 53) + 1));
        assert!(to_excel_serial(&lossy).is_err());

        let max = NaiveDateTime::new(
            NaiveDate::from_ymd_opt(9999, 12, 31).unwrap(),
            NaiveTime::from_hms_micro_opt(23, 59, 59, 999_500).unwrap(),
        );
        assert!(to_excel_serial(&Value::DateTime(max)).is_err());
        assert!(to_excel_serial(&Value::Duration(Duration::MAX)).is_err());
        assert!(exact_serial_from_millis(i64::MAX).is_err());
        assert!(exact_serial_from_millis(i64::MIN).is_ok());
    }

    #[test]
    fn rejects_chrono_leap_seconds_and_invalid_time_serial_components() {
        let leap = NaiveTime::from_hms_nano_opt(23, 59, 59, 1_500_000_000).unwrap();
        assert!(to_excel_serial(&Value::Time(leap)).is_err());
        assert_eq!(
            from_excel_serial(1.25, TemporalKind::Time),
            Some(Value::Number(1.25))
        );
        assert_eq!(
            from_excel_serial(-0.25, TemporalKind::Time),
            Some(Value::Number(-0.25))
        );
    }

    #[test]
    fn format_section_parser_honors_spacing_and_fill_escapes() {
        assert_eq!(format_kind("0_;yyyy-mm-dd"), Some(TemporalKind::Date));
        assert_eq!(format_kind("0*;yyyy-mm-dd"), Some(TemporalKind::Date));
        assert_eq!(format_kind("0_;0;yyyy-mm-dd"), None);
        assert_eq!(format_kind("0*;0;yyyy-mm-dd"), None);
    }
}
