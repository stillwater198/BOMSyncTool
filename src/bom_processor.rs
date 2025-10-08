use crate::{
    AutoCorrection, BomData, BomRow, ColumnDictionary, ColumnMapping, OverrideList,
    PreprocessRules, RegisteredNameEntry, RegisteredNameList, ValidationError, ValidationResult,
};
use calamine::{open_workbook, Reader, Xls, XlsError, Xlsx, XlsxError};
use csv::ReaderBuilder;
use encoding_rs::{SHIFT_JIS, UTF_8};
use rayon::prelude::*;
use serde::Serialize;
use serde_json;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::io::{Read, Seek};
use std::path::Path;
use thiserror::Error;

#[derive(Error, Debug)]
pub enum BomProcessorError {
    #[error("ファイル読み込みエラー: {0}")]
    FileReadError(String),
    #[error("ファイル形式エラー: {0}")]
    FormatError(String),
    #[error("エンコーディングエラー: {0}")]
    EncodingError(String),
    #[error("列指定エラー: {0}")]
    ColumnError(String),
}

const MAX_SAMPLE_ROWS: usize = 10;

#[derive(Debug, Clone, Serialize)]
pub struct FileAnalysis {
    pub headers: Vec<String>,
    pub suggested_mapping: Option<ColumnMapping>,
    pub sample_rows: Vec<Vec<String>>,
}

#[derive(Debug, Clone, Serialize)]
pub struct FilePreview {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
}

#[derive(Debug, Clone)]
pub struct LoadBomResult {
    pub bom: BomData,
    pub corrections: Vec<AutoCorrection>,
}

/// ファイル拡張子に基づいてBOMファイルを読み込む
pub async fn load_bom_file(
    file_path: &str,
    column_mapping: &ColumnMapping,
) -> Result<LoadBomResult, BomProcessorError> {
    let path = Path::new(file_path);
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    match extension.as_str() {
        "xlsx" | "xls" => load_excel_file(file_path, column_mapping).await,
        "csv" => load_csv_file(file_path, column_mapping).await,
        _ => Err(BomProcessorError::FormatError(
            "サポートされていないファイル形式です".to_string(),
        )),
    }
}

pub async fn analyze_bom_file(
    file_path: &str,
    dictionary: &ColumnDictionary,
) -> Result<FileAnalysis, BomProcessorError> {
    let path = Path::new(file_path);
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    match extension.as_str() {
        "xlsx" => analyze_excel_file(file_path, dictionary),
        "xls" => analyze_excel_file(file_path, dictionary),
        "csv" => analyze_csv_file(file_path, dictionary).await,
        _ => Err(BomProcessorError::FormatError(
            "サポートされていないファイル形式です".to_string(),
        )),
    }
}

pub async fn preview_raw_file(
    file_path: &str,
    limit: usize,
) -> Result<FilePreview, BomProcessorError> {
    let path = Path::new(file_path);
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    let capped_limit = limit.max(1).min(1000);

    match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(file_path)
                .map_err(|e: XlsxError| BomProcessorError::FileReadError(e.to_string()))?;
            preview_excel_workbook(&mut workbook, capped_limit)
        }
        "xls" => {
            let mut workbook: Xls<_> = open_workbook(file_path)
                .map_err(|e: XlsError| BomProcessorError::FileReadError(e.to_string()))?;
            preview_excel_workbook(&mut workbook, capped_limit)
        }
        "csv" => preview_csv_file(file_path, capped_limit).await,
        _ => Err(BomProcessorError::FormatError(
            "サポートされていないファイル形式です".to_string(),
        )),
    }
}

fn analyze_excel_file(
    file_path: &str,
    dictionary: &ColumnDictionary,
) -> Result<FileAnalysis, BomProcessorError> {
    let extension = Path::new(file_path)
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(file_path)
                .map_err(|e: XlsxError| BomProcessorError::FileReadError(e.to_string()))?;
            analyze_excel_workbook(&mut workbook, dictionary)
        }
        "xls" => {
            let mut workbook: Xls<_> = open_workbook(file_path)
                .map_err(|e: XlsError| BomProcessorError::FileReadError(e.to_string()))?;
            analyze_excel_workbook(&mut workbook, dictionary)
        }
        _ => Err(BomProcessorError::FormatError(
            "Excelファイルの拡張子が無効です".to_string(),
        )),
    }
}

fn analyze_excel_workbook<R, RS>(
    workbook: &mut R,
    dictionary: &ColumnDictionary,
) -> Result<FileAnalysis, BomProcessorError>
where
    R: Reader<RS>,
    RS: Read + Seek,
    R::Error: std::fmt::Display,
{
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| {
            BomProcessorError::FileReadError("ワークシートが見つかりません".to_string())
        })?
        .map_err(|e: R::Error| BomProcessorError::FileReadError(e.to_string()))?;

    let mut headers: Vec<String> = Vec::new();
    let mut sample_rows: Vec<Vec<String>> = Vec::new();

    for (row_idx, row) in range.rows().enumerate() {
        if row_idx == 0 {
            headers = row.iter().map(|cell| cell.to_string()).collect();
            continue;
        }
        if sample_rows.len() >= MAX_SAMPLE_ROWS {
            break;
        }
        let row_values: Vec<String> = row
            .iter()
            .map(|cell| standardize_string(&cell.to_string()))
            .collect();
        sample_rows.push(row_values);
    }

    let suggested_mapping = detect_column_mapping(&headers, &sample_rows, dictionary);

    Ok(FileAnalysis {
        headers,
        suggested_mapping,
        sample_rows,
    })
}

fn preview_excel_workbook<R, RS>(
    workbook: &mut R,
    limit: usize,
) -> Result<FilePreview, BomProcessorError>
where
    R: Reader<RS>,
    RS: Read + Seek,
    R::Error: std::fmt::Display,
{
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| {
            BomProcessorError::FileReadError("ワークシートが見つかりません".to_string())
        })?
        .map_err(|e: R::Error| BomProcessorError::FileReadError(e.to_string()))?;

    let mut headers: Vec<String> = Vec::new();
    let mut rows: Vec<Vec<String>> = Vec::new();

    for (row_idx, row) in range.rows().enumerate() {
        if row_idx == 0 {
            headers = row.iter().map(|cell| cell.to_string()).collect();
            continue;
        }
        if rows.len() >= limit {
            break;
        }
        rows.push(row.iter().map(|cell| cell.to_string()).collect());
    }

    Ok(FilePreview { headers, rows })
}

async fn analyze_csv_file(
    file_path: &str,
    dictionary: &ColumnDictionary,
) -> Result<FileAnalysis, BomProcessorError> {
    let content =
        fs::read(file_path).map_err(|e| BomProcessorError::FileReadError(e.to_string()))?;

    let decoded = if content.starts_with(&[0xEF, 0xBB, 0xBF]) {
        UTF_8.decode(&content[3..]).0
    } else if content.starts_with(&[0xFF, 0xFE]) {
        return Err(BomProcessorError::EncodingError(
            "UTF-16エンコーディングはサポートされていません".to_string(),
        ));
    } else {
        let utf8_result = UTF_8.decode(&content);
        if utf8_result.2 {
            utf8_result.0
        } else {
            SHIFT_JIS.decode(&content).0
        }
    };

    let mut reader = ReaderBuilder::new()
        .has_headers(true)
        .from_reader(decoded.as_bytes());

    let headers = reader
        .headers()
        .map_err(|e| BomProcessorError::FileReadError(e.to_string()))?
        .iter()
        .map(|h| h.to_string())
        .collect::<Vec<_>>();

    let mut sample_rows = Vec::new();
    for record in reader.records() {
        let record = record.map_err(|e| BomProcessorError::FileReadError(e.to_string()))?;
        let row_values: Vec<String> = record
            .iter()
            .map(|value| standardize_string(value))
            .collect();
        sample_rows.push(row_values);
        if sample_rows.len() >= MAX_SAMPLE_ROWS {
            break;
        }
    }

    let suggested_mapping = detect_column_mapping(&headers, &sample_rows, dictionary);

    Ok(FileAnalysis {
        headers,
        suggested_mapping,
        sample_rows,
    })
}

async fn preview_csv_file(file_path: &str, limit: usize) -> Result<FilePreview, BomProcessorError> {
    let content =
        fs::read(file_path).map_err(|e| BomProcessorError::FileReadError(e.to_string()))?;

    let decoded = if content.starts_with(&[0xEF, 0xBB, 0xBF]) {
        UTF_8.decode(&content[3..]).0
    } else if content.starts_with(&[0xFF, 0xFE]) {
        return Err(BomProcessorError::EncodingError(
            "UTF-16エンコーディングはサポートされていません".to_string(),
        ));
    } else {
        let utf8_result = UTF_8.decode(&content);
        if utf8_result.2 {
            utf8_result.0
        } else {
            SHIFT_JIS.decode(&content).0
        }
    };

    let mut reader = ReaderBuilder::new()
        .has_headers(true)
        .from_reader(decoded.as_bytes());

    let headers = reader
        .headers()
        .map_err(|e| BomProcessorError::FileReadError(e.to_string()))?
        .iter()
        .map(|h| h.to_string())
        .collect::<Vec<_>>();

    let mut rows = Vec::new();
    for record in reader.records() {
        if rows.len() >= limit {
            break;
        }
        let record = record.map_err(|e| BomProcessorError::FileReadError(e.to_string()))?;
        rows.push(record.iter().map(|value| value.to_string()).collect());
    }

    Ok(FilePreview { headers, rows })
}

/// Excelファイルを読み込む
async fn load_excel_file(
    file_path: &str,
    column_mapping: &ColumnMapping,
) -> Result<LoadBomResult, BomProcessorError> {
    let extension = Path::new(file_path)
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(file_path)
                .map_err(|e: XlsxError| BomProcessorError::FileReadError(e.to_string()))?;

            load_excel_workbook(&mut workbook, column_mapping)
        }
        "xls" => {
            let mut workbook: Xls<_> = open_workbook(file_path)
                .map_err(|e: XlsError| BomProcessorError::FileReadError(e.to_string()))?;

            load_excel_workbook(&mut workbook, column_mapping)
        }
        _ => Err(BomProcessorError::FormatError(
            "Excelファイルの拡張子が無効です".to_string(),
        )),
    }
}

/// Excelワークブックからデータを読み込む
fn load_excel_workbook<R, RS>(
    workbook: &mut R,
    column_mapping: &ColumnMapping,
) -> Result<LoadBomResult, BomProcessorError>
where
    R: Reader<RS>,
    RS: Read + Seek,
    R::Error: std::fmt::Display,
{
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| {
            BomProcessorError::FileReadError("ワークシートが見つかりません".to_string())
        })?
        .map_err(|e: R::Error| BomProcessorError::FileReadError(e.to_string()))?;

    let mut headers = Vec::new();
    let mut raw_rows: Vec<Vec<String>> = Vec::new();

    for (row_idx, row) in range.rows().enumerate() {
        if row_idx == 0 {
            headers = row.iter().map(|cell| cell.to_string()).collect();
            continue;
        }
        let row_values: Vec<String> = row.iter().map(|cell| cell.to_string()).collect();
        raw_rows.push(row_values);
    }

    build_bom_from_rows(headers, raw_rows, column_mapping)
}

fn detect_column_mapping(
    headers: &[String],
    rows: &[Vec<String>],
    dictionary: &ColumnDictionary,
) -> Option<ColumnMapping> {
    let max_columns = headers
        .len()
        .max(rows.iter().map(|row| row.len()).max().unwrap_or(0));

    if max_columns == 0 {
        return None;
    }

    let mut used: HashSet<usize> = HashSet::new();

    let part_idx = choose_column_from_dictionary("part_number", headers, rows, dictionary, &used)
        .map(|(idx, _)| idx)
        .or_else(|| find_text_column(max_columns, rows, &used))?;
    used.insert(part_idx);

    let model_idx = choose_column_from_dictionary("model_number", headers, rows, dictionary, &used)
        .map(|(idx, _)| idx)
        .or_else(|| find_text_column(max_columns, rows, &used))?;
    used.insert(model_idx);

    let manufacturer_idx =
        choose_column_from_dictionary("manufacturer", headers, rows, dictionary, &used)
            .map(|(idx, _)| idx);

    Some(ColumnMapping {
        part_number: part_idx,
        model_number: model_idx,
        manufacturer: manufacturer_idx,
    })
}

fn choose_column_from_dictionary(
    column_type: &str,
    headers: &[String],
    rows: &[Vec<String>],
    dictionary: &ColumnDictionary,
    used: &HashSet<usize>,
) -> Option<(usize, f32)> {
    let max_columns = headers
        .len()
        .max(rows.iter().map(|row| row.len()).max().unwrap_or(0));

    if max_columns == 0 {
        return None;
    }

    let patterns: Vec<String> = dictionary
        .patterns_for(column_type)
        .into_iter()
        .map(|p| normalize_token(&p))
        .filter(|p| !p.is_empty())
        .collect();

    let mut best: Option<(usize, f32)> = None;

    for idx in 0..max_columns {
        if used.contains(&idx) {
            continue;
        }
        let header_norm = headers
            .get(idx)
            .map(|h| normalize_token(h))
            .unwrap_or_default();

        let mut score = 0.0f32;

        if !patterns.is_empty() {
            let mut header_matches = 0f32;
            let mut value_ratio_total = 0f32;

            for pattern in &patterns {
                if pattern.is_empty() {
                    continue;
                }

                if header_norm.contains(pattern) || pattern.contains(&header_norm) {
                    header_matches += 1.0;
                    continue;
                }

                let (matches, total) = count_pattern_matches(idx, rows, pattern);
                if total > 0 {
                    value_ratio_total += matches as f32 / total as f32;
                }
            }

            let pattern_count = patterns.len() as f32;
            if pattern_count > 0.0 {
                score += (header_matches / pattern_count) * 2.0;
                score += value_ratio_total / pattern_count;
            }
        }

        if column_type.eq_ignore_ascii_case("part_number") {
            // Penalize columns with very few unique textual values
            let uniqueness = compute_uniqueness_ratio(idx, rows);
            if uniqueness > 0.0 {
                score += uniqueness * 0.3;
            }
        }

        if score <= 0.0 {
            continue;
        }

        match best {
            Some((_, best_score)) if score <= best_score => {}
            _ => best = Some((idx, score)),
        }
    }

    best
}

fn normalize_token(value: &str) -> String {
    value
        .trim()
        .chars()
        .filter(|c| !c.is_whitespace())
        .flat_map(|c| c.to_lowercase())
        .collect::<String>()
}

fn count_pattern_matches(col_idx: usize, rows: &[Vec<String>], pattern: &str) -> (usize, usize) {
    let mut matches = 0usize;
    let mut total = 0usize;
    for row in rows {
        if let Some(value) = row.get(col_idx) {
            let normalized = normalize_token(value);
            if normalized.is_empty() {
                continue;
            }
            total += 1;
            if normalized.contains(pattern) {
                matches += 1;
            }
        }
    }
    (matches, total)
}

fn compute_uniqueness_ratio(col_idx: usize, rows: &[Vec<String>]) -> f32 {
    let mut unique = HashSet::new();
    let mut total = 0usize;
    for row in rows {
        if let Some(value) = row.get(col_idx) {
            let normalized = normalize_token(value);
            if normalized.is_empty() {
                continue;
            }
            total += 1;
            unique.insert(normalized);
        }
    }
    if total == 0 {
        0.0
    } else {
        (unique.len() as f32 / total as f32).min(1.0)
    }
}

fn find_text_column(
    max_columns: usize,
    rows: &[Vec<String>],
    used: &HashSet<usize>,
) -> Option<usize> {
    for idx in 0..max_columns {
        if used.contains(&idx) {
            continue;
        }
        let mut score = 0usize;
        let mut total = 0usize;
        for row in rows {
            if let Some(value) = row.get(idx) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                total += 1;
                if trimmed.chars().any(|c| c.is_alphabetic()) {
                    score += 1;
                }
            }
        }
        if total > 0 && (score as f32 / total as f32) >= 0.3 {
            return Some(idx);
        }
    }

    for idx in 0..max_columns {
        if !used.contains(&idx) {
            return Some(idx);
        }
    }

    None
}

/// CSVファイルを読み込む
async fn load_csv_file(
    file_path: &str,
    column_mapping: &ColumnMapping,
) -> Result<LoadBomResult, BomProcessorError> {
    let content =
        fs::read(file_path).map_err(|e| BomProcessorError::FileReadError(e.to_string()))?;

    // エンコーディングを自動検出
    let (decoded_content, _, _) = if content.starts_with(&[0xEF, 0xBB, 0xBF]) {
        // UTF-8 BOM
        (UTF_8.decode(&content[3..]).0, UTF_8, true)
    } else if content.starts_with(&[0xFF, 0xFE]) {
        // UTF-16 LE BOM
        return Err(BomProcessorError::EncodingError(
            "UTF-16エンコーディングはサポートされていません".to_string(),
        ));
    } else {
        // まずUTF-8として試行
        let utf8_result = UTF_8.decode(&content);
        if utf8_result.2 {
            (utf8_result.0, UTF_8, false)
        } else {
            // UTF-8で失敗した場合はShift-JISとして試行
            let sjis_result = SHIFT_JIS.decode(&content);
            (sjis_result.0, SHIFT_JIS, false)
        }
    };

    let mut reader = ReaderBuilder::new()
        .has_headers(true)
        .from_reader(decoded_content.as_bytes());

    let mut headers = Vec::new();
    let mut raw_rows = Vec::new();

    // ヘッダーを取得
    if let Some(result) = reader.headers().ok() {
        headers = result.iter().map(|s| s.to_string()).collect();
    }

    // データ行を処理
    for result in reader.records() {
        let record = result.map_err(|e| BomProcessorError::FileReadError(e.to_string()))?;
        raw_rows.push(record.iter().map(|value| value.to_string()).collect());
    }

    build_bom_from_rows(headers, raw_rows, column_mapping)
}

fn build_bom_from_rows(
    mut headers: Vec<String>,
    raw_rows: Vec<Vec<String>>,
    column_mapping: &ColumnMapping,
) -> Result<LoadBomResult, BomProcessorError> {
    let mut max_required_index = column_mapping
        .part_number
        .max(column_mapping.model_number);

    if let Some(manufacturer_idx) = column_mapping.manufacturer {
        max_required_index = max_required_index.max(manufacturer_idx);
    }

    let mut max_columns = raw_rows
        .iter()
        .map(|row| row.len())
        .max()
        .unwrap_or(headers.len());
    if max_columns == 0 {
        max_columns = headers.len().max(max_required_index + 1);
    } else {
        max_columns = max_columns.max(max_required_index + 1);
    }

    if headers.len() < max_columns {
        for idx in headers.len()..max_columns {
            headers.push(format!("列{}", idx + 1));
        }
    }

    if headers.is_empty() {
        return Err(BomProcessorError::ColumnError(
            "ヘッダー行が存在しません".to_string(),
        ));
    }

    let mut rows = Vec::new();
    let mut corrections = Vec::new();

    for (row_idx, raw_row) in raw_rows.into_iter().enumerate() {
        let data_row_number = row_idx + 1;
        let mut pending: Vec<AutoCorrection> = Vec::new();
        let mut cells = vec![String::new(); headers.len()];

        for (col_idx, header) in headers.iter().enumerate() {
            let original_value = raw_row.get(col_idx).cloned().unwrap_or_default();
            let normalized = standardize_string(&original_value);
            let rule = string_correction_rule(col_idx, column_mapping);
            record_string_correction(
                &mut pending,
                data_row_number,
                col_idx,
                header,
                &original_value,
                &normalized,
                rule,
            );
            cells[col_idx] = normalized;
        }

        if column_mapping.part_number >= headers.len()
            || column_mapping.model_number >= headers.len()
            || column_mapping
                .manufacturer
                .map(|idx| idx >= headers.len())
                .unwrap_or(false)
        {
            return Err(BomProcessorError::ColumnError(
                "列番号の指定がヘッダー数を超えています".to_string(),
            ));
        }

        let part_number = cells[column_mapping.part_number].clone();
        if part_number.trim().is_empty() {
            continue;
        }

        let model_number = cells[column_mapping.model_number].clone();

        let mut attributes = HashMap::new();
        for (idx, header) in headers.iter().enumerate() {
            attributes.insert(header.clone(), cells.get(idx).cloned().unwrap_or_default());
        }

        rows.push(BomRow {
            part_number,
            model_number,
            attributes,
        });

        corrections.extend(pending.into_iter());
    }

    Ok(LoadBomResult {
        bom: BomData { headers, rows },
        corrections,
    })
}

fn string_correction_rule(column_index: usize, mapping: &ColumnMapping) -> &'static str {
    if column_index == mapping.part_number {
        "normalize_part_number"
    } else if column_index == mapping.model_number {
        "normalize_model_number"
    } else {
        "normalize_attribute"
    }
}

fn record_string_correction(
    bucket: &mut Vec<AutoCorrection>,
    row_number: usize,
    column_index: usize,
    column_name: &str,
    original: &str,
    normalized: &str,
    rule: &'static str,
) {
    if original == normalized {
        return;
    }
    bucket.push(AutoCorrection {
        row_number,
        column_index,
        column_name: column_name.to_string(),
        original_value: original.to_string(),
        corrected_value: normalized.to_string(),
        rule: rule.to_string(),
    });
}

/// 文字列を標準化する
pub fn standardize_string(input: &str) -> String {
    input
        .chars()
        .map(|c| {
            match c {
                // 全角数字を半角に変換
                '０'..='９' => char::from_u32(c as u32 - 0xFEE0).unwrap_or(c),
                // 全角英字を半角に変換
                'Ａ'..='Ｚ' => char::from_u32(c as u32 - 0xFEE0).unwrap_or(c),
                'ａ'..='ｚ' => char::from_u32(c as u32 - 0xFEE0).unwrap_or(c),
                // 改行文字を削除
                '\n' | '\r' => ' ',
                // その他の文字はそのまま
                _ => c,
            }
        })
        .collect::<String>()
        .replace(" ", "") // 空白を削除
        .to_uppercase() // 大文字に変換
}

/// 部品表データを並列処理で最適化
pub fn optimize_bom_data(bom_data: &mut BomData) {
    let mut part_map: HashMap<String, BomRow> = HashMap::new();

    for mut row in bom_data.rows.drain(..) {
        part_map
            .entry(row.part_number.clone())
            .and_modify(|existing_row| {
                for (key, value) in row.attributes.drain() {
                    existing_row.attributes.insert(key, value);
                }
            })
            .or_insert(row);
    }

    bom_data.rows = part_map.into_values().collect();

    // 並列処理でソート
    bom_data
        .rows
        .par_sort_by(|a, b| a.part_number.cmp(&b.part_number));
}

pub fn preprocess_bom_data(
    bom_data: &BomData,
    rules: &PreprocessRules,
) -> Result<BomData, BomProcessorError> {
    let mut processed_rows: Vec<BomRow> = Vec::new();

    for original in &bom_data.rows {
        let mut base_row = original.clone();

        base_row.part_number = apply_string_rules(&base_row.part_number, rules);
        base_row.model_number = apply_string_rules(&base_row.model_number, rules);

        for value in base_row.attributes.values_mut() {
            *value = apply_string_rules(value, rules);
        }

        let mut expanded_rows: Vec<BomRow> = Vec::new();

        if rules.expand_ranges {
            if let Some(expanded) = expand_ranges(&base_row.part_number) {
                let original_part = base_row.part_number.clone();
                for part in expanded {
                    let mut cloned = base_row.clone();
                    cloned.part_number = apply_string_rules(&part, rules);
                    replace_attribute_value(
                        &mut cloned.attributes,
                        &original_part,
                        &cloned.part_number,
                    );
                    expanded_rows.push(cloned);
                }
            }
        }

        if expanded_rows.is_empty() {
            expanded_rows.push(base_row);
        }

        processed_rows.extend(expanded_rows);
    }

    let mut result = bom_data.clone();
    result.rows = processed_rows;
    Ok(result)
}

fn apply_string_rules(value: &str, rules: &PreprocessRules) -> String {
    let mut result = value.to_string();
    if rules.remove_parentheses {
        result = remove_parentheses(&result);
    }
    if rules.fullwidth_to_halfwidth {
        result = fullwidth_to_halfwidth(&result);
    }
    if rules.lowercase_to_uppercase {
        result = result.to_uppercase();
    }
    result
}

fn replace_attribute_value(
    attributes: &mut HashMap<String, String>,
    original_value: &str,
    new_value: &str,
) {
    for value in attributes.values_mut() {
        if value == original_value {
            *value = new_value.to_string();
        }
    }
}

fn remove_parentheses(input: &str) -> String {
    input.replace('(', "").replace(')', "")
}

fn expand_ranges(input: &str) -> Option<Vec<String>> {
    if let Some(dash_pos) = input.find('-') {
        let prefix = &input[..dash_pos];
        let suffix = &input[dash_pos + 1..];

        if let (Some(start_num), Some(end_num)) = (extract_number(prefix), extract_number(suffix)) {
            if start_num < end_num && end_num - start_num <= 100 {
                let base = prefix
                    .trim_end_matches(|c: char| c.is_ascii_digit())
                    .to_string();
                let mut result = Vec::new();
                for i in start_num..=end_num {
                    result.push(format!("{}{}", base, i));
                }
                return Some(result);
            }
        }
    }
    None
}

fn extract_number(input: &str) -> Option<u32> {
    let digits: String = input
        .chars()
        .rev()
        .take_while(|c| c.is_ascii_digit())
        .collect::<Vec<_>>()
        .into_iter()
        .rev()
        .collect();
    digits.parse().ok()
}

fn fullwidth_to_halfwidth(input: &str) -> String {
    input
        .chars()
        .map(|c| match c {
            '０'..='９' | 'Ａ'..='Ｚ' | 'ａ'..='ｚ' => {
                char::from_u32(c as u32 - 0xFEE0).unwrap_or(c)
            }
            _ => c,
        })
        .collect()
}

pub async fn load_registered_name_csv(
    file_path: &str,
) -> Result<RegisteredNameList, BomProcessorError> {
    let content =
        fs::read(file_path).map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

    let (decoded_content, _, _) = if content.starts_with(&[0xEF, 0xBB, 0xBF]) {
        (UTF_8.decode(&content[3..]).0, UTF_8, true)
    } else {
        let utf8_result = UTF_8.decode(&content);
        if utf8_result.2 {
            (utf8_result.0, UTF_8, false)
        } else {
            let sjis_result = SHIFT_JIS.decode(&content);
            (sjis_result.0, SHIFT_JIS, false)
        }
    };

    let mut reader = ReaderBuilder::new()
        .has_headers(true)
        .from_reader(decoded_content.as_bytes());

    let mut entries = Vec::new();

    for result in reader.records() {
        let record = result.map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;
        if record.len() < 2 {
            continue;
        }
        entries.push(RegisteredNameEntry {
            part_model: record[0].to_string(),
            registered_name: record[1].to_string(),
        });
    }

    Ok(RegisteredNameList { entries })
}

pub async fn load_registered_name_json(
    file_path: &str,
) -> Result<RegisteredNameList, BomProcessorError> {
    let content = fs::read_to_string(file_path)
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

    let list: RegisteredNameList = serde_json::from_str(&content)
        .map_err(|e| BomProcessorError::FormatError(format!("JSON解析エラー: {}", e)))?;

    Ok(list)
}

pub async fn save_registered_name_csv(
    list: &RegisteredNameList,
    file_path: &str,
) -> Result<(), BomProcessorError> {
    let mut csv_data = Vec::new();
    csv_data.push(vec!["部品型番".to_string(), "登録名".to_string()]);

    for entry in &list.entries {
        csv_data.push(vec![
            entry.part_model.clone(),
            entry.registered_name.clone(),
        ]);
    }

    crate::file_handler::save_csv_file(&csv_data, file_path, "utf-8")
        .await
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

    Ok(())
}

pub async fn save_registered_name_json(
    list: &RegisteredNameList,
    file_path: &str,
) -> Result<(), BomProcessorError> {
    let json_content = serde_json::to_string_pretty(list)
        .map_err(|e| BomProcessorError::FormatError(format!("JSON生成エラー: {}", e)))?;

    fs::write(file_path, json_content)
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

    Ok(())
}

pub fn apply_registered_names_to_bom(
    bom_data: &mut BomData,
    registered_name_list: &Option<RegisteredNameList>,
    override_list: &Option<OverrideList>,
) {
    let override_map: HashMap<String, String> = override_list
        .as_ref()
        .map(|list| {
            list.entries
                .iter()
                .map(|entry| (entry.part_number.clone(), entry.registered_name.clone()))
                .collect()
        })
        .unwrap_or_default();

    let registered_name_map: HashMap<String, String> = registered_name_list
        .as_ref()
        .map(|list| {
            list.entries
                .iter()
                .map(|entry| (entry.part_model.clone(), entry.registered_name.clone()))
                .collect()
        })
        .unwrap_or_default();

    for row in &mut bom_data.rows {
        if let Some(override_name) = override_map.get(&row.part_number) {
            row.attributes
                .insert("登録名".to_string(), override_name.clone());
        } else if let Some(registered_name) = registered_name_map.get(&row.model_number) {
            row.attributes
                .insert("登録名".to_string(), registered_name.clone());
        }
    }
}

pub fn validate_bom_data(bom_data: &BomData) -> ValidationResult {
    let mut errors = Vec::new();

    for (index, row) in bom_data.rows.iter().enumerate() {
        let row_number = index + 1;

        if row.part_number.trim().is_empty() {
            errors.push(ValidationError {
                row_number,
                field: "部品番号".to_string(),
                message: "部品番号は必須です".to_string(),
            });
        }

        if row.model_number.trim().is_empty() {
            errors.push(ValidationError {
                row_number,
                field: "型番".to_string(),
                message: "型番は必須です".to_string(),
            });
        }

        let duplicate_count = bom_data
            .rows
            .iter()
            .filter(|r| r.part_number == row.part_number)
            .count();
        if duplicate_count > 1 {
            errors.push(ValidationError {
                row_number,
                field: "部品番号".to_string(),
                message: format!("部品番号 '{}' が重複しています", row.part_number),
            });
        }

        if !row
            .part_number
            .chars()
            .all(|c| c.is_alphanumeric() || c == '-' || c == '_')
        {
            errors.push(ValidationError {
                row_number,
                field: "部品番号".to_string(),
                message: "部品番号は英数字、ハイフン、アンダースコアのみ使用できます".to_string(),
            });
        }

        if !row
            .model_number
            .chars()
            .all(|c| c.is_alphanumeric() || c == '-' || c == '_' || c == '.')
        {
            errors.push(ValidationError {
                row_number,
                field: "型番".to_string(),
                message: "型番は英数字、ハイフン、アンダースコア、ピリオドのみ使用できます"
                    .to_string(),
            });
        }
    }

    ValidationResult {
        is_valid: errors.is_empty(),
        errors,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_standardize_string() {
        assert_eq!(standardize_string("ABC123"), "ABC123");
        assert_eq!(standardize_string("ＡＢＣ１２３"), "ABC123");
        assert_eq!(standardize_string("abc\n123"), "ABC123");
        assert_eq!(standardize_string("A B C"), "ABC");
    }

}
