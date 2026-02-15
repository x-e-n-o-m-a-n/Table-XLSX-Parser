// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use base64::{engine::general_purpose::STANDARD, Engine as _};
use calamine::{Reader, Xlsx};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use rust_xlsxwriter::Workbook;
use serde_json::Value as JsonValue;
use std::collections::{BTreeMap, HashMap};
use std::fs::File;
use std::io::Cursor;
use std::io::Write;
use std::path::Path;

const MAX_EXCEL_COLUMNS: usize = 16_384;
const MAX_EXCEL_ROWS: u32 = 1_048_576;

#[derive(serde::Serialize)]
struct ConvertResponse {
    xml_content: String,
    status: String,
}

#[derive(serde::Serialize)]
struct SplitResponse {
    status: String,
    output_path: String,
    sheets_created: usize,
    rows_exported: usize,
    skipped_invalid_rows: usize,
}

#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[tauri::command]
fn parse_xlsx_from_bytes(file_data: String) -> Result<ConvertResponse, String> {
    println!("Received file_data, length: {}", file_data.len());
    let bytes = STANDARD
        .decode(&file_data)
        .map_err(|e| format!("Failed to decode base64: {}", e))?;
    println!("Decoded bytes length: {}", bytes.len());

    let mut workbook = Xlsx::new(Cursor::new(bytes)).map_err(|e| {
        println!("Failed to open workbook: {}", e);
        format!("Не удалось открыть файл: {}", e)
    })?;
    println!("Workbook opened successfully");

    let sheet_names = workbook.sheet_names();
    println!("Sheet names: {:?}", sheet_names);
    let sheet_name = sheet_names
        .first()
        .ok_or_else(|| "В файле нет листов".to_string())?
        .to_owned();
    println!("Using sheet: {}", sheet_name);

    let range = workbook.worksheet_range(&sheet_name).map_err(|e| {
        println!("Failed to get worksheet range: {}", e);
        format!("Ошибка при чтении листа: {}", e)
    })?;
    println!("Range dimensions: {}x{}", range.height(), range.width());
    if range.height() == 0 || range.width() == 0 {
        return Err("Лист пустой или не удалось прочитать данные".to_string());
    }

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer
        .write_event(Event::Start(BytesStart::new("workbook")))
        .map_err(|e| e.to_string())?;

    let drop_cols: [usize; 5] = [0, 1, 5, 6, 13];
    let mut row_count = 0;
    let mut cell_count = 0;
    for (ridx, row) in range.rows().enumerate() {
        if ridx == 0 {
            continue;
        }
        row_count += 1;
        writer
            .write_event(Event::Start(BytesStart::new("row")))
            .map_err(|e| e.to_string())?;

        for (cidx, cell) in row.iter().enumerate() {
            if drop_cols.contains(&cidx) {
                continue;
            }
            cell_count += 1;
            writer
                .write_event(Event::Start(BytesStart::new("cell")))
                .map_err(|e| e.to_string())?;

            let cell_value = format!("{}", cell);
            writer
                .write_event(Event::Text(BytesText::new(&cell_value)))
                .map_err(|e| e.to_string())?;

            writer
                .write_event(Event::End(BytesEnd::new("cell")))
                .map_err(|e| e.to_string())?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("row")))
            .map_err(|e| e.to_string())?;
    }
    println!("Processed {} rows, {} cells", row_count, cell_count);

    writer
        .write_event(Event::End(BytesEnd::new("workbook")))
        .map_err(|e| e.to_string())?;

    let result = writer.into_inner().into_inner();
    let xml_string = String::from_utf8(result).map_err(|e| e.to_string())?;
    println!("XML conversion successful, length: {}", xml_string.len());

    Ok(ConvertResponse {
        xml_content: xml_string,
        status: "Success".to_string(),
    })
}

#[tauri::command]
fn parse_xlsx_from_path(
    file_path: String,
    cols: Option<Vec<usize>>,
    mode: Option<String>,
    skip_header: Option<bool>,
    format: Option<String>,
) -> Result<ConvertResponse, String> {
    println!(
        "Received file_path: {} cols: {:?} mode: {:?} skip_header: {:?} format: {:?}",
        file_path, cols, mode, skip_header, format
    );
    let file = std::fs::File::open(&file_path).map_err(|e| format!("Ошибка открытия файла: {}", e))?;
    let mut workbook = calamine::Xlsx::new(std::io::BufReader::new(file))
        .map_err(|e| format!("Не удалось открыть файл: {}", e))?;
    let sheet_names = workbook.sheet_names();
    let sheet_name = sheet_names
        .first()
        .ok_or_else(|| "В файле нет листов".to_string())?
        .to_owned();
    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("Ошибка при чтении листа: {}", e))?;
    if range.height() == 0 || range.width() == 0 {
        return Err("Лист пустой или не удалось прочитать данные".to_string());
    }

    use std::collections::HashSet;
    let mut cols_set: HashSet<usize> = HashSet::new();
    if let Some(v) = cols {
        for n in v {
            if n == 0 {
                continue;
            }
            cols_set.insert(n - 1);
        }
    }

    let mode_str = mode.unwrap_or_else(|| "drop".to_string()).to_lowercase();
    let skip_header_flag = skip_header.unwrap_or(true);
    let fmt = format.unwrap_or_else(|| "xml".to_string()).to_lowercase();

    if fmt == "csv" {
        let mut lines: Vec<String> = Vec::new();
        for (ridx, row) in range.rows().enumerate() {
            if ridx == 0 && skip_header_flag {
                continue;
            }
            let mut cells_out: Vec<String> = Vec::new();
            for (cidx, cell) in row.iter().enumerate() {
                let include = if mode_str == "keep" {
                    cols_set.contains(&cidx)
                } else {
                    !cols_set.contains(&cidx)
                };
                if !include {
                    continue;
                }
                let cell_value = format!("{}", cell);
                let mut v = cell_value.replace('"', "\"\"");
                if v.contains(',') || v.contains('"') || v.contains('\n') || v.contains('\r') {
                    v = format!("\"{}\"", v);
                }
                cells_out.push(v);
            }
            lines.push(cells_out.join(","));
        }
        let csv = lines.join("\r\n");
        return Ok(ConvertResponse {
            xml_content: csv,
            status: "Success".to_string(),
        });
    }

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer
        .write_event(Event::Start(BytesStart::new("workbook")))
        .map_err(|e| e.to_string())?;

    for (ridx, row) in range.rows().enumerate() {
        if ridx == 0 && skip_header_flag {
            continue;
        }
        writer
            .write_event(Event::Start(BytesStart::new("row")))
            .map_err(|e| e.to_string())?;
        for (cidx, cell) in row.iter().enumerate() {
            let include = if mode_str == "keep" {
                cols_set.contains(&cidx)
            } else {
                !cols_set.contains(&cidx)
            };
            if !include {
                continue;
            }
            writer
                .write_event(Event::Start(BytesStart::new("cell")))
                .map_err(|e| e.to_string())?;
            let cell_value = format!("{}", cell);
            writer
                .write_event(Event::Text(BytesText::new(&cell_value)))
                .map_err(|e| e.to_string())?;
            writer
                .write_event(Event::End(BytesEnd::new("cell")))
                .map_err(|e| e.to_string())?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("row")))
            .map_err(|e| e.to_string())?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("workbook")))
        .map_err(|e| e.to_string())?;
    let result = writer.into_inner().into_inner();
    let xml_string = String::from_utf8(result).map_err(|e| e.to_string())?;
    Ok(ConvertResponse {
        xml_content: xml_string,
        status: "Success".to_string(),
    })
}

fn parse_order_and_operation(value: &str) -> Option<(String, String)> {
    let parts = value.split('/').map(|v| v.trim()).collect::<Vec<_>>();
    if parts.len() != 2 {
        return None;
    }
    if parts[0].is_empty() || parts[1].is_empty() {
        return None;
    }
    Some((parts[0].to_string(), parts[1].to_string()))
}

fn get_cell_value(row: &[calamine::Data], idx: usize) -> String {
    row.get(idx).map(|cell| format!("{}", cell)).unwrap_or_default()
}

fn sanitize_sheet_name(raw: &str) -> String {
    let trimmed = raw.trim();
    let src = if trimmed.is_empty() { "NO_ORDER" } else { trimmed };

    let mut cleaned = String::with_capacity(src.len());
    for ch in src.chars() {
        let is_invalid = matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']') || ch.is_control();
        cleaned.push(if is_invalid { '_' } else { ch });
    }

    let mut cleaned = cleaned.trim_matches('\'').trim().to_string();
    if cleaned.is_empty() {
        cleaned = "NO_ORDER".to_string();
    }
    cleaned.chars().take(31).collect()
}

fn make_unique_sheet_name(base: &str, used: &mut HashMap<String, usize>) -> String {
    let counter = used.entry(base.to_string()).or_insert(0);
    *counter += 1;
    if *counter == 1 {
        return base.to_string();
    }

    let suffix = format!("_{}", *counter);
    let keep_len = 31usize.saturating_sub(suffix.chars().count());
    let stem = base.chars().take(keep_len).collect::<String>();
    format!("{}{}", stem, suffix)
}

fn write_row(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row_index: u32,
    values: &[String],
) -> Result<(), String> {
    if row_index >= MAX_EXCEL_ROWS {
        return Err("Превышен лимит строк Excel".to_string());
    }
    for (cidx, value) in values.iter().enumerate() {
        if cidx >= MAX_EXCEL_COLUMNS {
            return Err("Превышен лимит столбцов Excel".to_string());
        }
        worksheet
            .write_string(row_index, cidx as u16, value)
            .map_err(|e| format!("Ошибка записи XLSX: {}", e))?;
    }
    Ok(())
}

#[tauri::command]
fn split_xlsx_by_order(file_path: String, output_path: String) -> Result<SplitResponse, String> {
    let output = Path::new(&output_path);
    let out_parent = output
        .parent()
        .ok_or_else(|| "Некорректный путь выходного файла".to_string())?;
    if !out_parent.exists() || !out_parent.is_dir() {
        return Err("Папка выходного файла не найдена".to_string());
    }

    let file = std::fs::File::open(&file_path).map_err(|e| format!("Ошибка открытия файла: {}", e))?;
    let mut workbook = calamine::Xlsx::new(std::io::BufReader::new(file))
        .map_err(|e| format!("Не удалось открыть файл: {}", e))?;

    let source_sheet = workbook
        .sheet_names()
        .first()
        .ok_or_else(|| "В файле нет листов".to_string())?
        .to_owned();
    let range = workbook
        .worksheet_range(&source_sheet)
        .map_err(|e| format!("Ошибка при чтении листа: {}", e))?;

    if range.height() == 0 || range.width() == 0 {
        return Err("Лист пустой или не удалось прочитать данные".to_string());
    }
    if range.width() <= 11 {
        return Err("Ожидаются данные минимум до столбца L".to_string());
    }

    let mut rows_iter = range.rows();
    let header_source = rows_iter
        .next()
        .ok_or_else(|| "Не удалось прочитать строку заголовка".to_string())?;

    let transformed_header = vec![
        get_cell_value(header_source, 2),
        get_cell_value(header_source, 3),
        "Номер заказа".to_string(),
        "Номер операции".to_string(),
        get_cell_value(header_source, 7),
        get_cell_value(header_source, 8),
        get_cell_value(header_source, 9),
        get_cell_value(header_source, 10),
        get_cell_value(header_source, 11),
    ];

    let mut groups: BTreeMap<String, Vec<Vec<String>>> = BTreeMap::new();
    let mut skipped_invalid_rows = 0usize;
    let mut rows_exported = 0usize;

    for row in rows_iter {
        let g_value = get_cell_value(row, 6);
        let Some((order_no, operation_no)) = parse_order_and_operation(&g_value) else {
            skipped_invalid_rows += 1;
            continue;
        };

        let transformed_row = vec![
            get_cell_value(row, 2),
            get_cell_value(row, 3),
            order_no.clone(),
            operation_no,
            get_cell_value(row, 7),
            get_cell_value(row, 8),
            get_cell_value(row, 9),
            get_cell_value(row, 10),
            get_cell_value(row, 11),
        ];

        groups.entry(order_no).or_default().push(transformed_row);
        rows_exported += 1;
    }

    if groups.is_empty() {
        return Err("Нет валидных строк: столбец G должен быть в формате <заказ>/<операция>".to_string());
    }

    let mut output_workbook = Workbook::new();
    let mut used_sheet_names: HashMap<String, usize> = HashMap::new();
    let mut sheets_created = 0usize;

    for (order_no, rows) in groups {
        let base_sheet_name = sanitize_sheet_name(&order_no);
        let sheet_name = make_unique_sheet_name(&base_sheet_name, &mut used_sheet_names);

        let worksheet = output_workbook.add_worksheet();
        worksheet
            .set_name(&sheet_name)
            .map_err(|e| format!("Ошибка имени листа '{}': {}", sheet_name, e))?;

        write_row(worksheet, 0, &transformed_header)?;
        for (idx, row_values) in rows.iter().enumerate() {
            write_row(worksheet, (idx + 1) as u32, row_values)?;
        }
        sheets_created += 1;
    }

    output_workbook
        .save(output)
        .map_err(|e| format!("Не удалось сохранить {}: {}", output.to_string_lossy(), e))?;

    Ok(SplitResponse {
        status: "Success".to_string(),
        output_path,
        sheets_created,
        rows_exported,
        skipped_invalid_rows,
    })
}

#[tauri::command]
fn save_xml_to_file(xml_content: String, file_path: String) -> Result<String, String> {
    let mut file = File::create(&file_path).map_err(|e| format!("Failed to create file: {}", e))?;

    file.write_all(xml_content.as_bytes())
        .map_err(|e| format!("Failed to write to file: {}", e))?;

    Ok("File saved successfully".to_string())
}

#[tauri::command]
fn save_xml_to_file_flexible(payload: JsonValue) -> Result<String, String> {
    let xml = payload
        .get("xml_content")
        .or_else(|| payload.get("xmlContent"))
        .and_then(|v| v.as_str())
        .ok_or_else(|| "missing required key xmlContent/xml_content".to_string())?;
    let path = payload
        .get("file_path")
        .or_else(|| payload.get("filePath"))
        .and_then(|v| v.as_str())
        .ok_or_else(|| "missing required key filePath/file_path".to_string())?;

    let mut file = File::create(path).map_err(|e| format!("Failed to create file: {}", e))?;
    file.write_all(xml.as_bytes())
        .map_err(|e| format!("Failed to write to file: {}", e))?;
    Ok("File saved successfully".to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            greet,
            parse_xlsx_from_bytes,
            parse_xlsx_from_path,
            split_xlsx_by_order,
            save_xml_to_file,
            save_xml_to_file_flexible
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
