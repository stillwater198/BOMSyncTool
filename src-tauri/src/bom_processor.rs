use crate::{BomData, BomRow, ColumnMapping, PreprocessRules, RegisteredNameList, RegisteredNameEntry, OverrideList, ValidationResult, ValidationError, CorrectionEntry, CorrectionType};
use calamine::{Reader, open_workbook, Xlsx, Xls};
use csv::ReaderBuilder;
use encoding_rs::{SHIFT_JIS, UTF_8};
use rayon::prelude::*;
use std::collections::HashMap;
use std::fs;
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


/// ファイル拡張子に基づいてBOMファイルを読み込む
pub async fn load_bom_file(
    file_path: &str,
    column_mapping: &ColumnMapping,
) -> Result<BomData, BomProcessorError> {
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

/// Excelファイルを読み込む
async fn load_excel_file(
    file_path: &str,
    column_mapping: &ColumnMapping,
) -> Result<BomData, BomProcessorError> {
    let extension = Path::new(file_path)
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(file_path)
                .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;
            
            load_excel_workbook(&mut workbook, column_mapping)
        }
        "xls" => {
            let mut workbook: Xls<_> = open_workbook(file_path)
                .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;
            
            load_excel_workbook(&mut workbook, column_mapping)
        }
        _ => Err(BomProcessorError::FormatError(
            "Excelファイルの拡張子が無効です".to_string(),
        )),
    }
}

/// Excelワークブックからデータを読み込む
fn load_excel_workbook<R>(
    workbook: &mut R,
    column_mapping: &ColumnMapping,
) -> Result<BomData, BomProcessorError>
where
    R: Reader<std::io::BufReader<std::fs::File>>,
    <R as Reader<std::io::BufReader<std::fs::File>>>::Error: std::fmt::Display,
{
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| BomProcessorError::FileReadError("ワークシートが見つかりません".to_string()))?
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

    let mut headers = Vec::new();
    let mut rows = Vec::new();

    for (row_idx, row) in range.rows().enumerate() {
        if row_idx == 0 {
            // ヘッダー行
            for cell in row {
                headers.push(cell.to_string());
            }
            continue;
        }

        let required_index = column_mapping
            .part_number
            .max(column_mapping.model_number);
        if row.len() <= required_index {
            continue; // 必要な列が存在しない場合はスキップ
        }

        let part_number = standardize_string(&row[column_mapping.part_number].to_string());
        let model_number = standardize_string(&row[column_mapping.model_number].to_string());

        if part_number.is_empty() {
            continue; // 部品番号が空の場合はスキップ
        }

        let mut attributes = HashMap::new();
        for (col_idx, header) in headers.iter().enumerate() {
            if col_idx < row.len() {
                attributes.insert(
                    header.clone(),
                    standardize_string(&row[col_idx].to_string()),
                );
            }
        }

        rows.push(BomRow {
            part_number,
            model_number,
            attributes,
        });
    }

    Ok(BomData { headers, rows })
}

/// CSVファイルを読み込む
async fn load_csv_file(
    file_path: &str,
    column_mapping: &ColumnMapping,
) -> Result<BomData, BomProcessorError> {
    let content = fs::read(file_path)
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

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
    let mut rows = Vec::new();

    // ヘッダーを取得
    if let Some(result) = reader.headers().ok() {
        headers = result.iter().map(|s| s.to_string()).collect();
    }

    // データ行を処理
    for result in reader.records() {
        let record = result.map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

        let required_index = column_mapping
            .part_number
            .max(column_mapping.model_number);
        if record.len() <= required_index {
            continue; // 必要な列が存在しない場合はスキップ
        }

        let part_number = standardize_string(&record[column_mapping.part_number]);
        let model_number = standardize_string(&record[column_mapping.model_number]);

        if part_number.is_empty() {
            continue; // 部品番号が空の場合はスキップ
        }

        let mut attributes = HashMap::new();
        for (col_idx, header) in headers.iter().enumerate() {
            if col_idx < record.len() {
                attributes.insert(
                    header.clone(),
                    standardize_string(&record[col_idx]),
                );
            }
        }

        rows.push(BomRow {
            part_number,
            model_number,
            attributes,
        });
    }

    Ok(BomData { headers, rows })
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


/// BOMデータの前処理を実行
pub fn preprocess_bom_data(bom_data: &BomData, rules: &PreprocessRules) -> Result<BomData, BomProcessorError> {
    let mut processed_data = bom_data.clone();
    
    // 各ルールに基づいて前処理を実行
    for row in &mut processed_data.rows {
        // 括弧削除
        if rules.remove_parentheses {
            let original_part = row.part_number.clone();
            let original_model = row.model_number.clone();
            
            row.part_number = remove_parentheses(&row.part_number);
            row.model_number = remove_parentheses(&row.model_number);
            
            // 修正ログに記録
            if original_part != row.part_number {
                log_correction(
                    "部品番号",
                    original_part,
                    row.part_number.clone(),
                    "括弧削除",
                    CorrectionType::Auto,
                );
            }
            if original_model != row.model_number {
                log_correction(
                    "型番",
                    original_model,
                    row.model_number.clone(),
                    "括弧削除",
                    CorrectionType::Auto,
                );
            }
        }
        
        // 範囲展開
        if rules.expand_ranges {
            // 範囲展開は部品番号に対してのみ適用
            if let Some(expanded) = expand_ranges(&row.part_number) {
                // 範囲展開された場合は複数行に分割
                // この実装では単純化のため、最初の部品番号のみを保持
                let original = row.part_number.clone();
                row.part_number = expanded[0].clone();
                
                if original != row.part_number {
                    log_correction(
                        "部品番号",
                        original,
                        row.part_number.clone(),
                        "範囲展開",
                        CorrectionType::Auto,
                    );
                }
            }
        }
        
        // 全角→半角変換
        if rules.fullwidth_to_halfwidth {
            let original_part = row.part_number.clone();
            let original_model = row.model_number.clone();
            
            row.part_number = fullwidth_to_halfwidth(&row.part_number);
            row.model_number = fullwidth_to_halfwidth(&row.model_number);
            
            if original_part != row.part_number {
                log_correction(
                    "部品番号",
                    original_part,
                    row.part_number.clone(),
                    "全角→半角変換",
                    CorrectionType::Auto,
                );
            }
            if original_model != row.model_number {
                log_correction(
                    "型番",
                    original_model,
                    row.model_number.clone(),
                    "全角→半角変換",
                    CorrectionType::Auto,
                );
            }
        }
        
        // 小文字→大文字変換
        if rules.lowercase_to_uppercase {
            let original_part = row.part_number.clone();
            let original_model = row.model_number.clone();
            
            row.part_number = row.part_number.to_uppercase();
            row.model_number = row.model_number.to_uppercase();
            
            if original_part != row.part_number {
                log_correction(
                    "部品番号",
                    original_part,
                    row.part_number.clone(),
                    "小文字→大文字変換",
                    CorrectionType::Auto,
                );
            }
            if original_model != row.model_number {
                log_correction(
                    "型番",
                    original_model,
                    row.model_number.clone(),
                    "小文字→大文字変換",
                    CorrectionType::Auto,
                );
            }
        }
        
        // 属性も同様に処理
        for (key, value) in &mut row.attributes {
            if rules.remove_parentheses {
                let original = value.clone();
                *value = remove_parentheses(value);
                if original != *value {
                    log_correction(
                        key,
                        original,
                        value.clone(),
                        "括弧削除",
                        CorrectionType::Auto,
                    );
                }
            }
            if rules.fullwidth_to_halfwidth {
                let original = value.clone();
                *value = fullwidth_to_halfwidth(value);
                if original != *value {
                    log_correction(
                        key,
                        original,
                        value.clone(),
                        "全角→半角変換",
                        CorrectionType::Auto,
                    );
                }
            }
            if rules.lowercase_to_uppercase {
                let original = value.clone();
                *value = value.to_uppercase();
                if original != *value {
                    log_correction(
                        key,
                        original,
                        value.clone(),
                        "小文字→大文字変換",
                        CorrectionType::Auto,
                    );
                }
            }
        }
    }
    
    Ok(processed_data)
}

// 修正ログを記録（実際の実装ではグローバルなログストレージに保存）
fn log_correction(
    column_name: &str,
    original_value: String,
    corrected_value: String,
    rule_applied: &str,
    correction_type: CorrectionType,
) {
    // 実際の実装では、グローバルなログストレージに保存
    // ここでは簡易的にコンソールに出力
    println!(
        "修正ログ: {} | {} -> {} | ルール: {} | 種別: {:?}",
        column_name, original_value, corrected_value, rule_applied, correction_type
    );
}

/// 括弧を削除する
fn remove_parentheses(input: &str) -> String {
    input.replace("(", "").replace(")", "")
}

/// 範囲を展開する（例: "C1-C3" → ["C1", "C2", "C3"]）
fn expand_ranges(input: &str) -> Option<Vec<String>> {
    if let Some(dash_pos) = input.find('-') {
        let prefix = &input[..dash_pos];
        let suffix = &input[dash_pos + 1..];
        
        // 数字部分を抽出
        if let (Some(start_num), Some(end_num)) = (extract_number(prefix), extract_number(suffix)) {
            if start_num < end_num && end_num - start_num <= 100 { // 安全制限
                let mut result = Vec::new();
                for i in start_num..=end_num {
                    result.push(format!("{}{}", prefix.trim_end_matches(char::is_numeric), i));
                }
                return Some(result);
            }
        }
    }
    None
}

/// 文字列から数字を抽出する
fn extract_number(input: &str) -> Option<u32> {
    input.chars()
        .rev()
        .take_while(|c| c.is_ascii_digit())
        .collect::<String>()
        .chars()
        .rev()
        .collect::<String>()
        .parse()
        .ok()
}

/// 全角文字を半角に変換する
fn fullwidth_to_halfwidth(input: &str) -> String {
    input.chars().map(|c| {
        match c {
            '０'..='９' => char::from_u32(c as u32 - 0xFEE0).unwrap_or(c),
            'Ａ'..='Ｚ' => char::from_u32(c as u32 - 0xFEE0).unwrap_or(c),
            'ａ'..='ｚ' => char::from_u32(c as u32 - 0xFEE0).unwrap_or(c),
            _ => c,
        }
    }).collect()
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
    bom_data.rows.par_sort_by(|a, b| a.part_number.cmp(&b.part_number));
}

/// 登録名リストをCSVから読み込む
pub async fn load_registered_name_csv(file_path: &str) -> Result<RegisteredNameList, BomProcessorError> {
    let content = fs::read(file_path)
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;

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

/// 登録名リストをJSONから読み込む
pub async fn load_registered_name_json(file_path: &str) -> Result<RegisteredNameList, BomProcessorError> {
    let content = fs::read_to_string(file_path)
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;
    
    let list: RegisteredNameList = serde_json::from_str(&content)
        .map_err(|e| BomProcessorError::FormatError(format!("JSON解析エラー: {}", e)))?;
    
    Ok(list)
}

/// 登録名リストをCSVに保存
pub async fn save_registered_name_csv(list: &RegisteredNameList, file_path: &str) -> Result<(), BomProcessorError> {
    let mut csv_data = Vec::new();
    
    // ヘッダー行
    csv_data.push(vec!["部品型番".to_string(), "登録名".to_string()]);
    
    // データ行
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

/// 登録名リストをJSONに保存
pub async fn save_registered_name_json(list: &RegisteredNameList, file_path: &str) -> Result<(), BomProcessorError> {
    let json_content = serde_json::to_string_pretty(list)
        .map_err(|e| BomProcessorError::FormatError(format!("JSON生成エラー: {}", e)))?;
    
    fs::write(file_path, json_content)
        .map_err(|e| BomProcessorError::FileReadError(format!("{}", e)))?;
    
    Ok(())
}

/// 登録名をBOMデータに適用
pub fn apply_registered_names_to_bom(
    bom_data: &mut BomData,
    registered_name_list: &Option<RegisteredNameList>,
    override_list: &Option<OverrideList>,
) {
    // 上書きリストからマップを作成
    let override_map: HashMap<String, String> = override_list
        .as_ref()
        .map(|list| {
            list.entries
                .iter()
                .map(|entry| (entry.part_number.clone(), entry.registered_name.clone()))
                .collect()
        })
        .unwrap_or_default();

    // 登録名リストからマップを作成
    let registered_name_map: HashMap<String, String> = registered_name_list
        .as_ref()
        .map(|list| {
            list.entries
                .iter()
                .map(|entry| (entry.part_model.clone(), entry.registered_name.clone()))
                .collect()
        })
        .unwrap_or_default();

    // 各BOM行に登録名を適用
    for row in &mut bom_data.rows {
        // 適用順序: override → 登録名リスト → デフォルト
        if let Some(override_name) = override_map.get(&row.part_number) {
            row.attributes.insert("登録名".to_string(), override_name.clone());
        } else if let Some(registered_name) = registered_name_map.get(&row.model_number) {
            row.attributes.insert("登録名".to_string(), registered_name.clone());
        }
    }
}

/// BOMデータのバリデーションを実行
pub fn validate_bom_data(bom_data: &BomData) -> ValidationResult {
    let mut errors = Vec::new();
    
    for (index, row) in bom_data.rows.iter().enumerate() {
        let row_number = index + 1; // 1ベースの行番号
        
        // 部品番号のバリデーション
        if row.part_number.trim().is_empty() {
            errors.push(ValidationError {
                row_number,
                field: "部品番号".to_string(),
                message: "部品番号は必須です".to_string(),
            });
        }
        
        // 型番のバリデーション
        if row.model_number.trim().is_empty() {
            errors.push(ValidationError {
                row_number,
                field: "型番".to_string(),
                message: "型番は必須です".to_string(),
            });
        }
        
        // 部品番号の重複チェック
        let duplicate_count = bom_data.rows
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
        
        // 部品番号の形式チェック（英数字とハイフンのみ）
        if !row.part_number.chars().all(|c| c.is_alphanumeric() || c == '-' || c == '_') {
            errors.push(ValidationError {
                row_number,
                field: "部品番号".to_string(),
                message: "部品番号は英数字、ハイフン、アンダースコアのみ使用できます".to_string(),
            });
        }
        
        // 型番の形式チェック
        if !row.model_number.chars().all(|c| c.is_alphanumeric() || c == '-' || c == '_' || c == '.') {
            errors.push(ValidationError {
                row_number,
                field: "型番".to_string(),
                message: "型番は英数字、ハイフン、アンダースコア、ピリオドのみ使用できます".to_string(),
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
