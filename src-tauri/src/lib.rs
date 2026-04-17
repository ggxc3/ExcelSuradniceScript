use calamine::{open_workbook_auto, Data, Reader};
use serde::Deserialize;
use thiserror::Error;

#[derive(Debug, Deserialize)]
pub struct ProcessRequest {
    pub input_path: String,
    pub output_path: String,
    pub sheet: String,
    pub columns: String,
    pub start_row: usize,
    pub end_row: usize,
}

#[derive(Debug, Error)]
pub enum AppError {
    #[error("{0}")]
    Message(String),
}

fn parse_columns_spec(spec: &str) -> Result<Vec<(String, String)>, AppError> {
    let normalized = spec.to_uppercase().replace(' ', "");
    if normalized.is_empty() {
        return Err(AppError::Message(
            "Pole Stĺpce je povinné (napr. A-N,B-E,C-V)".into(),
        ));
    }

    normalized
        .split(',')
        .map(|part| {
            let split: Vec<&str> = part.split('-').collect();
            if split.len() != 2 || split[0].is_empty() || split[1].is_empty() {
                return Err(AppError::Message(format!(
                    "Neplatný formát stĺpcov: {part}"
                )));
            }
            if !matches!(split[1], "N" | "E" | "V") {
                return Err(AppError::Message(format!(
                    "Neplatný typ {} v {} (povolené: N, E, V)",
                    split[1], part
                )));
            }
            Ok((split[0].to_string(), split[1].to_string()))
        })
        .collect()
}

fn is_digit(c: char) -> bool {
    c.is_ascii_digit()
}

fn starts_with_two_digits(value: &str) -> bool {
    value.chars().take(2).all(is_digit) && value.chars().count() >= 2
}

fn format_coordinate(value: &str, coordinate_type: &str) -> Option<String> {
    if value.is_empty()
        || value.chars().count() < 5
        || value.chars().nth(1).is_some_and(|c| c == '.' || c == ',')
    {
        return None;
    }

    let mut text = value.replace(' ', "").replace(',', ".");
    if text.starts_with('0') {
        text.remove(0);
    }

    let first = text.chars().next()?;
    let second = text.chars().nth(1)?;
    if !(starts_with_two_digits(&text) || ((first == 'E' || first == 'N') && is_digit(second))) {
        return None;
    }

    if coordinate_type == "E" {
        text = text.replace('N', "");
        if !text.contains('E') {
            text = format!("{}E{}", &text[..2], &text[2..]);
        }
    } else {
        text = text.replace('E', "");
        if !text.contains('N') {
            text = format!("{}N{}", &text[..2], &text[2..]);
        }
    }

    Some(text)
}

fn format_height(value: &str) -> String {
    if value.is_empty() {
        return String::new();
    }

    let normalized = value.replace(' ', "").replace('.', ",");
    let parse_target = normalized.replace(',', ".");

    if let Ok(float_value) = parse_target.parse::<f64>() {
        return float_value.round().to_string();
    }

    normalized
        .chars()
        .take_while(|c| c.is_ascii_digit())
        .collect::<String>()
}

fn column_to_index(col: &str) -> usize {
    col.chars()
        .fold(0usize, |acc, c| acc * 26 + ((c as u8 - b'A' + 1) as usize))
}

fn index_to_column(mut index: usize) -> String {
    let mut col = String::new();
    while index > 0 {
        let rem = (index - 1) % 26;
        col.insert(0, (b'A' + rem as u8) as char);
        index = (index - 1) / 26;
    }
    col
}

fn data_to_string(data: &Data) -> String {
    match data {
        Data::Empty => String::new(),
        _ => data.to_string(),
    }
}

pub fn run(request: ProcessRequest) -> Result<String, AppError> {
    if request.input_path.trim().is_empty() {
        return Err(AppError::Message("Vyber vstupný Excel súbor".into()));
    }
    if request.output_path.trim().is_empty() {
        return Err(AppError::Message("Zvoľ cieľový súbor pre uloženie".into()));
    }
    if request.sheet.trim().is_empty() {
        return Err(AppError::Message("Vyber hárok".into()));
    }
    if request.start_row == 0 || request.end_row == 0 {
        return Err(AppError::Message(
            "Riadky musia byť kladné celé čísla".into(),
        ));
    }
    if request.start_row > request.end_row {
        return Err(AppError::Message(
            "Počiatočný riadok nesmie byť väčší ako koncový".into(),
        ));
    }

    let columns = parse_columns_spec(&request.columns)?;
    let mut workbook = open_workbook_auto(&request.input_path)
        .map_err(|e| AppError::Message(format!("Nepodarilo sa otvoriť vstupný súbor: {e}")))?;

    let range = workbook.worksheet_range(&request.sheet).map_err(|e| {
        AppError::Message(format!("Hárok {} sa v súbore nenašiel: {e}", request.sheet))
    })?;

    let mut rows: Vec<Vec<String>> = range
        .rows()
        .map(|r| r.iter().map(data_to_string).collect::<Vec<String>>())
        .collect();

    let mut effective_end = request.end_row.min(rows.len());

    for (column, coordinate_type) in columns {
        let col_index = column_to_index(&column);

        if coordinate_type == "V" {
            for row in request.start_row..=effective_end {
                if let Some(cell) = rows.get_mut(row - 1).and_then(|r| r.get_mut(col_index - 1)) {
                    *cell = format_height(cell);
                }
            }
            continue;
        }

        let mut delete_rows = Vec::new();
        for row in request.start_row..=effective_end {
            if let Some(cell) = rows.get_mut(row - 1).and_then(|r| r.get_mut(col_index - 1)) {
                if let Some(formatted) = format_coordinate(cell, &coordinate_type) {
                    *cell = formatted;
                } else {
                    delete_rows.push(row);
                }
            } else {
                delete_rows.push(row);
            }
        }

        for row in delete_rows.iter().rev() {
            rows.remove(*row - 1);
        }
        effective_end = effective_end.saturating_sub(delete_rows.len());
    }

    let mut target = umya_spreadsheet::new_file();
    let _ = target.new_sheet(&request.sheet);
    let sheet = target
        .get_sheet_by_name_mut(&request.sheet)
        .ok_or_else(|| AppError::Message("Nepodarilo sa pripraviť cieľový hárok".into()))?;

    for (ridx, row) in rows.iter().enumerate() {
        for (cidx, value) in row.iter().enumerate() {
            let coordinate = format!("{}{}", index_to_column(cidx + 1), ridx + 1);
            sheet.get_cell_mut(coordinate.as_str()).set_value(value);
        }
    }

    umya_spreadsheet::writer::xlsx::write(&target, &request.output_path)
        .map_err(|e| AppError::Message(format!("Nepodarilo sa uložiť výstupný súbor: {e}")))?;

    Ok(format!(
        "Hotovo. Súbor bol úspešne uložený: {}",
        request.output_path
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_columns() {
        let cols = parse_columns_spec("a-n, b-e, c-v").unwrap();
        assert_eq!(cols[0], ("A".to_string(), "N".to_string()));
        assert!(parse_columns_spec("A-X").is_err());
    }

    #[test]
    fn formats_values() {
        assert_eq!(format_coordinate("48123", "N"), Some("48N123".into()));
        assert_eq!(format_coordinate("48N123", "E"), Some("48E123".into()));
        assert_eq!(format_height("123.6m"), "123");
    }
}
