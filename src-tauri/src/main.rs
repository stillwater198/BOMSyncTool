// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::Mutex;
use tauri::api::dialog::FileDialogBuilder;
use tauri::Manager;
use tauri::State;
use tokio::sync::oneshot;

mod bom_processor;
mod comparison;
mod file_handler;
mod synthesis;

use bom_processor::{load_bom_file, BomProcessorError};
use comparison::*;
use synthesis::*;

// アプリケーションの状態管理
#[derive(Debug, Default)]
pub struct AppState {
    pub bom_a: Mutex<Option<BomData>>,
    pub bom_b: Mutex<Option<BomData>>,
    pub comparison_result: Mutex<Option<ComparisonResult>>,
    pub synthesis_result: Mutex<Option<SynthesisResult>>,
    pub registered_name_list: Mutex<Option<RegisteredNameList>>,
    pub override_list: Mutex<Option<OverrideList>>,
    // プレビュー用のCSVデータ
    pub preview_csv_a: Mutex<Option<String>>,
    pub preview_csv_b: Mutex<Option<String>>,
    // 修正ログ
    pub correction_log: Mutex<Vec<CorrectionEntry>>,
}

// 部品データ構造
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BomData {
    pub headers: Vec<String>,
    pub rows: Vec<BomRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BomRow {
    pub part_number: String,
    pub model_number: String,
    pub attributes: HashMap<String, String>,
}

// 列指定の構造体
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnMapping {
    pub part_number: usize,
    pub model_number: usize,
    #[serde(default)]
    pub manufacturer: Option<usize>,
}

// 比較結果
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ComparisonResult {
    pub common_parts: Vec<ComparisonRow>,
    pub a_only_parts: Vec<ComparisonRow>,
    pub b_only_parts: Vec<ComparisonRow>,
    pub modified_parts: Vec<ComparisonRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ComparisonRow {
    pub part_number: String,
    pub model_a: String,
    pub model_b: String,
    pub status: String,      // "common", "a_only", "b_only", "modified"
    pub change_type: String, // "ADDED", "REMOVED", "MODIFIED", "UNCHANGED"
}

// 合成結果
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SynthesisResult {
    pub rows: Vec<SynthesisRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SynthesisRow {
    pub part_number: String,
    pub model_a: String,
    pub model_b: String,
    pub status: String, // "common", "missing_a", "missing_b"
}

// 進捗イベントを送信するヘルパー関数
fn emit_progress(app_handle: &tauri::AppHandle, progress: u32) {
    let _ = app_handle.emit_all(
        "progress_update",
        serde_json::json!({
            "progress": progress
        }),
    );
}

// ファイル読み込みコマンド（進捗付き）
#[tauri::command]
async fn load_file(
    file_path: String,
    column_mapping: ColumnMapping,
    side: String, // "a" or "b"
    state: State<'_, AppState>,
    app_handle: tauri::AppHandle,
) -> Result<String, String> {
    emit_progress(&app_handle, 10);

    match load_bom_file(&file_path, &column_mapping).await {
        Ok(bom_data) => {
            emit_progress(&app_handle, 70);

            // CSVデータを生成
            let csv_content = bom_data_to_csv(&bom_data);

            emit_progress(&app_handle, 90);

            if side == "a" {
                *state.bom_a.lock().unwrap() = Some(bom_data);
                *state.preview_csv_a.lock().unwrap() = Some(csv_content);
            } else if side == "b" {
                *state.bom_b.lock().unwrap() = Some(bom_data);
                *state.preview_csv_b.lock().unwrap() = Some(csv_content);
            } else {
                return Err("無効なサイド指定です".to_string());
            }

            emit_progress(&app_handle, 100);

            Ok(format!("部品表{}を読み込みました", side.to_uppercase()))
        }
        Err(e) => Err(format!("ファイル読み込みエラー: {}", e)),
    }
}

// BOMデータをCSV文字列に変換
fn bom_data_to_csv(bom_data: &BomData) -> String {
    let mut csv_content = String::new();

    // ヘッダー行
    csv_content.push_str(&bom_data.headers.join(","));
    csv_content.push('\n');

    for row in &bom_data.rows {
        let row_data: Vec<String> = bom_data
            .headers
            .iter()
            .map(|header| row.attributes.get(header).cloned().unwrap_or_default())
            .collect();
        csv_content.push_str(&row_data.join(","));
        csv_content.push('\n');
    }

    csv_content
}

// 比較実行コマンド（進捗付き）
#[tauri::command]
async fn compare_boms(
    state: State<'_, AppState>,
    app_handle: tauri::AppHandle,
) -> Result<ComparisonResult, String> {
    emit_progress(&app_handle, 10);

    let (bom_a, bom_b) = {
        let bom_a = state.bom_a.lock().unwrap();
        let bom_b = state.bom_b.lock().unwrap();
        (bom_a.clone(), bom_b.clone())
    };

    emit_progress(&app_handle, 30);

    match (bom_a, bom_b) {
        (Some(a), Some(b)) => {
            emit_progress(&app_handle, 50);
            let result = perform_comparison(&a, &b).await;
            emit_progress(&app_handle, 80);

            *state.comparison_result.lock().unwrap() = Some(result.clone());
            emit_progress(&app_handle, 100);

            Ok(result)
        }
        _ => Err("部品表AまたはBが読み込まれていません".to_string()),
    }
}

// 合成実行コマンド（進捗付き）
#[tauri::command]
async fn synthesize_boms(
    state: State<'_, AppState>,
    app_handle: tauri::AppHandle,
) -> Result<SynthesisResult, String> {
    emit_progress(&app_handle, 10);

    let (bom_a, bom_b) = {
        let bom_a = state.bom_a.lock().unwrap();
        let bom_b = state.bom_b.lock().unwrap();
        (bom_a.clone(), bom_b.clone())
    };

    emit_progress(&app_handle, 30);

    match (bom_a, bom_b) {
        (Some(a), Some(b)) => {
            emit_progress(&app_handle, 50);
            let result = perform_synthesis(&a, &b).await;
            emit_progress(&app_handle, 80);

            *state.synthesis_result.lock().unwrap() = Some(result.clone());
            emit_progress(&app_handle, 100);

            Ok(result)
        }
        _ => Err("部品表AまたはBが読み込まれていません".to_string()),
    }
}

// 結果保存コマンド
#[tauri::command]
async fn save_result(
    file_path: String,
    format: String,      // "csv" or "txt"
    result_type: String, // "comparison" or "synthesis"
    state: State<'_, AppState>,
) -> Result<String, String> {
    match result_type.as_str() {
        "comparison" => {
            let comparison = {
                let comparison = state.comparison_result.lock().unwrap();
                comparison.clone()
            };
            match comparison {
                Some(result) => save_comparison_result(&result, &file_path, &format).await,
                None => Err("比較結果がありません".to_string()),
            }
        }
        "synthesis" => {
            let synthesis = {
                let synthesis = state.synthesis_result.lock().unwrap();
                synthesis.clone()
            };
            match synthesis {
                Some(result) => save_synthesis_result(&result, &file_path, &format).await,
                None => Err("合成結果がありません".to_string()),
            }
        }
        _ => Err("無効な結果タイプです".to_string()),
    }
}

// BOM前処理コマンド
#[tauri::command]
async fn preprocess_bom(
    bom_data: BomData,
    rules: PreprocessRules,
    state: State<'_, AppState>,
) -> Result<BomData, String> {
    let processed_data = bom_processor::preprocess_bom_data(&bom_data, &rules)
        .map_err(|e| format!("前処理エラー: {}", e))?;
    Ok(processed_data)
}

// BOMスナップショット取得コマンド
#[tauri::command]
async fn get_bom_snapshot(
    side: String,
    state: State<'_, AppState>,
) -> Result<Option<BomData>, String> {
    match side.as_str() {
        "a" => {
            let bom_a = state.bom_a.lock().unwrap();
            Ok(bom_a.clone())
        }
        "b" => {
            let bom_b = state.bom_b.lock().unwrap();
            Ok(bom_b.clone())
        }
        _ => Err("無効なサイド指定です".to_string()),
    }
}

// BOMデータ更新コマンド
#[tauri::command]
async fn update_bom_data(
    side: String,
    bom_data: BomData,
    state: State<'_, AppState>,
) -> Result<String, String> {
    match side.as_str() {
        "a" => {
            *state.bom_a.lock().unwrap() = Some(bom_data);
            Ok("部品表Aを更新しました".to_string())
        }
        "b" => {
            *state.bom_b.lock().unwrap() = Some(bom_data);
            Ok("部品表Bを更新しました".to_string())
        }
        _ => Err("無効なサイド指定です".to_string()),
    }
}

// 前処理ルール構造体
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PreprocessRules {
    pub remove_parentheses: bool,
    pub expand_ranges: bool,
    pub fullwidth_to_halfwidth: bool,
    pub lowercase_to_uppercase: bool,
}

// 登録名リスト構造体
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RegisteredNameList {
    pub entries: Vec<RegisteredNameEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RegisteredNameEntry {
    pub part_model: String,
    pub registered_name: String,
}

// 上書きリスト構造体
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OverrideList {
    pub entries: Vec<OverrideEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OverrideEntry {
    pub part_number: String,
    pub registered_name: String,
}

// バリデーションエラー構造体
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationError {
    pub row_number: usize,
    pub field: String,
    pub message: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationResult {
    pub is_valid: bool,
    pub errors: Vec<ValidationError>,
}

// 修正ログエントリ
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CorrectionEntry {
    pub timestamp: String,
    pub row_number: usize,
    pub column_name: String,
    pub original_value: String,
    pub corrected_value: String,
    pub rule_applied: String,
    pub correction_type: CorrectionType,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CorrectionType {
    Auto,       // 自動修正
    Manual,     // 手動修正
    Validation, // バリデーション修正
}

// プレビュー用のCSVデータ構造
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PreviewData {
    pub csv_content: String,
    pub headers: Vec<String>,
    pub row_count: usize,
    pub validation_errors: Vec<ValidationError>,
}

// 登録名リスト読み込みコマンド
#[tauri::command]
async fn load_registered_name_list(
    file_path: String,
    format: String,
    state: State<'_, AppState>,
) -> Result<String, String> {
    let list = match format.as_str() {
        "csv" => bom_processor::load_registered_name_csv(&file_path)
            .await
            .map_err(|e| format!("CSV読み込みエラー: {}", e))?,
        "json" => bom_processor::load_registered_name_json(&file_path)
            .await
            .map_err(|e| format!("JSON読み込みエラー: {}", e))?,
        _ => return Err("サポートされていないフォーマットです".to_string()),
    };

    *state.registered_name_list.lock().unwrap() = Some(list);
    Ok("登録名リストを読み込みました".to_string())
}

// 登録名リスト保存コマンド
#[tauri::command]
async fn save_registered_name_list(
    file_path: String,
    format: String,
    state: State<'_, AppState>,
) -> Result<String, String> {
    let list = {
        let registered_name_list = state.registered_name_list.lock().unwrap();
        registered_name_list.clone()
    };

    match list {
        Some(list) => {
            match format.as_str() {
                "csv" => bom_processor::save_registered_name_csv(&list, &file_path)
                    .await
                    .map_err(|e| format!("CSV保存エラー: {}", e))?,
                "json" => bom_processor::save_registered_name_json(&list, &file_path)
                    .await
                    .map_err(|e| format!("JSON保存エラー: {}", e))?,
                _ => return Err("サポートされていないフォーマットです".to_string()),
            }
            Ok("登録名リストを保存しました".to_string())
        }
        None => Err("登録名リストがありません".to_string()),
    }
}

// 登録名リストをBOMに適用コマンド
#[tauri::command]
async fn apply_registered_names(
    side: String,
    state: State<'_, AppState>,
) -> Result<String, String> {
    let registered_name_list = {
        let registered_name_list = state.registered_name_list.lock().unwrap();
        registered_name_list.clone()
    };

    let override_list = {
        let override_list = state.override_list.lock().unwrap();
        override_list.clone()
    };

    match side.as_str() {
        "a" => {
            let mut bom_a = {
                let bom_a = state.bom_a.lock().unwrap();
                bom_a.clone()
            };

            if let Some(mut bom) = bom_a {
                bom_processor::apply_registered_names_to_bom(
                    &mut bom,
                    &registered_name_list,
                    &override_list,
                );
                *state.bom_a.lock().unwrap() = Some(bom);
                Ok("部品表Aに登録名を適用しました".to_string())
            } else {
                Err("部品表Aが読み込まれていません".to_string())
            }
        }
        "b" => {
            let mut bom_b = {
                let bom_b = state.bom_b.lock().unwrap();
                bom_b.clone()
            };

            if let Some(mut bom) = bom_b {
                bom_processor::apply_registered_names_to_bom(
                    &mut bom,
                    &registered_name_list,
                    &override_list,
                );
                *state.bom_b.lock().unwrap() = Some(bom);
                Ok("部品表Bに登録名を適用しました".to_string())
            } else {
                Err("部品表Bが読み込まれていません".to_string())
            }
        }
        _ => Err("無効なサイド指定です".to_string()),
    }
}

// 上書きリスト適用コマンド
#[tauri::command]
async fn apply_overrides(
    overrides: OverrideList,
    state: State<'_, AppState>,
) -> Result<String, String> {
    *state.override_list.lock().unwrap() = Some(overrides);
    Ok("上書きリストを適用しました".to_string())
}

// 登録名リスト取得コマンド
#[tauri::command]
async fn get_registered_name_list(
    state: State<'_, AppState>,
) -> Result<Option<RegisteredNameList>, String> {
    let registered_name_list = state.registered_name_list.lock().unwrap();
    Ok(registered_name_list.clone())
}

// 上書きリスト取得コマンド
#[tauri::command]
async fn get_override_list(state: State<'_, AppState>) -> Result<Option<OverrideList>, String> {
    let override_list = state.override_list.lock().unwrap();
    Ok(override_list.clone())
}

// BOMデータバリデーションコマンド
#[tauri::command]
async fn validate_bom_data(
    side: String,
    state: State<'_, AppState>,
) -> Result<ValidationResult, String> {
    let bom_data = match side.as_str() {
        "a" => {
            let bom_a = state.bom_a.lock().unwrap();
            match bom_a.as_ref() {
                Some(bom) => bom.clone(),
                None => return Err("部品表Aが読み込まれていません".to_string()),
            }
        }
        "b" => {
            let bom_b = state.bom_b.lock().unwrap();
            match bom_b.as_ref() {
                Some(bom) => bom.clone(),
                None => return Err("部品表Bが読み込まれていません".to_string()),
            }
        }
        _ => return Err("無効なサイド指定です".to_string()),
    };

    let validation_result = bom_processor::validate_bom_data(&bom_data);
    Ok(validation_result)
}

// プレビューデータ取得コマンド
#[tauri::command]
async fn get_preview_data(side: String, state: State<'_, AppState>) -> Result<PreviewData, String> {
    let (bom_data, preview_csv) = match side.as_str() {
        "a" => {
            let bom_a = state.bom_a.lock().unwrap();
            let preview_csv_a = state.preview_csv_a.lock().unwrap();
            (bom_a.clone(), preview_csv_a.clone())
        }
        "b" => {
            let bom_b = state.bom_b.lock().unwrap();
            let preview_csv_b = state.preview_csv_b.lock().unwrap();
            (bom_b.clone(), preview_csv_b.clone())
        }
        _ => return Err("無効なサイド指定です".to_string()),
    };

    match (bom_data, preview_csv) {
        (Some(bom), Some(csv)) => {
            // バリデーション実行
            let validation_result = bom_processor::validate_bom_data(&bom);

            Ok(PreviewData {
                csv_content: csv,
                headers: bom.headers,
                row_count: bom.rows.len(),
                validation_errors: validation_result.errors,
            })
        }
        _ => Err("プレビュー対象のデータがありません".to_string()),
    }
}

// 修正ログ取得コマンド
#[tauri::command]
async fn get_correction_log(state: State<'_, AppState>) -> Result<Vec<CorrectionEntry>, String> {
    let log = state.correction_log.lock().unwrap();
    Ok(log.clone())
}

// 修正ログをCSVで出力
#[tauri::command]
async fn export_correction_log_csv(
    file_path: String,
    state: State<'_, AppState>,
) -> Result<String, String> {
    let log = state.correction_log.lock().unwrap();

    let mut csv_content = String::new();
    csv_content.push_str("タイムスタンプ,行番号,列名,元の値,修正後の値,適用ルール,修正種別\n");

    for entry in log.iter() {
        csv_content.push_str(&format!(
            "{},{},{},{},{},{},{}\n",
            entry.timestamp,
            entry.row_number,
            entry.column_name,
            entry.original_value,
            entry.corrected_value,
            entry.rule_applied,
            match entry.correction_type {
                CorrectionType::Auto => "自動",
                CorrectionType::Manual => "手動",
                CorrectionType::Validation => "バリデーション",
            }
        ));
    }

    std::fs::write(&file_path, csv_content)
        .map_err(|e| format!("修正ログの保存に失敗しました: {}", e))?;

    Ok(format!("修正ログを保存しました: {}", file_path))
}

// データクリアコマンド（拡張版）
#[tauri::command]
async fn clear_data(mode: String, state: State<'_, AppState>) -> Result<String, String> {
    match mode.as_str() {
        "all" => {
            // 全クリア
            *state.bom_a.lock().unwrap() = None;
            *state.bom_b.lock().unwrap() = None;
            *state.comparison_result.lock().unwrap() = None;
            *state.synthesis_result.lock().unwrap() = None;
            *state.registered_name_list.lock().unwrap() = None;
            *state.override_list.lock().unwrap() = None;
            *state.preview_csv_a.lock().unwrap() = None;
            *state.preview_csv_b.lock().unwrap() = None;
            *state.correction_log.lock().unwrap() = Vec::new();
            Ok("全データをクリアしました".to_string())
        }
        "session_keep" => {
            // セッション残しクリア（ETC1〜3相当のデータを保持）
            *state.bom_a.lock().unwrap() = None;
            *state.bom_b.lock().unwrap() = None;
            *state.comparison_result.lock().unwrap() = None;
            *state.synthesis_result.lock().unwrap() = None;
            *state.preview_csv_a.lock().unwrap() = None;
            *state.preview_csv_b.lock().unwrap() = None;
            *state.correction_log.lock().unwrap() = Vec::new();
            // registered_name_list と override_list は保持
            Ok("セッションデータを保持してクリアしました".to_string())
        }
        _ => Err("無効なクリアモードです".to_string()),
    }
}

// シートクリアコマンド（後方互換性のため残す）
#[tauri::command]
async fn clear_sheets(state: State<'_, AppState>) -> Result<String, String> {
    clear_data("all".to_string(), state).await
}

fn main() {
    tauri::Builder::default()
        .manage(AppState::default())
        .invoke_handler(tauri::generate_handler![
            load_file,
            compare_boms,
            synthesize_boms,
            save_result,
            clear_sheets,
            clear_data,
            open_file_dialog,
            preprocess_bom,
            get_bom_snapshot,
            update_bom_data,
            load_registered_name_list,
            save_registered_name_list,
            apply_registered_names,
            apply_overrides,
            get_registered_name_list,
            get_override_list,
            validate_bom_data,
            get_preview_data,
            get_correction_log,
            export_correction_log_csv
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[tauri::command]
async fn open_file_dialog() -> Result<Option<String>, String> {
    let (tx, rx) = oneshot::channel();
    FileDialogBuilder::new()
        .set_can_create_directories(false)
        .add_filter("BOM ファイル", &["csv", "xls", "xlsx"])
        .pick_file(move |file| {
            let _ = tx.send(file.map(|p| p.to_string_lossy().into_owned()));
        });

    match rx.await {
        Ok(path) => Ok(path),
        Err(_) => Ok(None),
    }
}
