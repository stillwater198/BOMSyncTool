// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use chrono::Utc;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::env;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use tauri::State;
use tauri_plugin_dialog::DialogExt;
use tokio::sync::oneshot;

mod bom_processor;
mod comparison;
mod file_handler;
mod session;
mod synthesis;
use comparison::*;
use session::{
    collect_snapshots, delete_snapshot, load_snapshot, save_snapshot, SessionKind, SessionSnapshot,
};
use synthesis::*;

const SETTINGS_DIR: &str = "../sessions/settings";
const SETTINGS_FILE_NAME: &str = "bom_settings.json";
const SETTINGS_ACTIONS: &[&str] = &["copy_above", "expand_range", "replace_with", "ignore"];

const DICTIONARY_DIR: &str = "../dictionary";
const DICTIONARY_FILE_NAME: &str = "custom_dict.json";
const AUTO_PREVIEW_LIMIT: usize = 15;

// アプリケーションの状態管理
#[derive(Debug)]
pub struct AppState {
    pub bom_a: Mutex<Option<BomData>>,
    pub bom_b: Mutex<Option<BomData>>,
    pub comparison_result: Mutex<Option<ComparisonResult>>,
    pub synthesis_result: Mutex<Option<SynthesisResult>>,
    pub registered_name_list: Mutex<Option<RegisteredNameList>>,
    pub override_list: Mutex<Option<OverrideList>>,
    pub file_a_path: Mutex<Option<String>>,
    pub file_b_path: Mutex<Option<String>>,
    pub column_mapping_a: Mutex<Option<ColumnMapping>>,
    pub column_mapping_b: Mutex<Option<ColumnMapping>>,
    pub settings: Mutex<AppSettings>,
    pub column_dictionary: Mutex<ColumnDictionary>,
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
    #[serde(default)]
    pub modified_parts: Vec<ComparisonRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ComparisonRow {
    pub part_number: String,
    pub model_a: String,
    pub model_b: String,
    pub status: String, // "common", "a_only", "b_only"
    #[serde(default = "default_change_type")]
    pub change_type: String, // "ADDED", "REMOVED", "MODIFIED", "UNCHANGED"
}

fn default_change_type() -> String {
    "UNCHANGED".to_string()
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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PreprocessRules {
    pub remove_parentheses: bool,
    pub expand_ranges: bool,
    pub fullwidth_to_halfwidth: bool,
    pub lowercase_to_uppercase: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct RegisteredNameList {
    pub entries: Vec<RegisteredNameEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RegisteredNameEntry {
    pub part_model: String,
    pub registered_name: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct OverrideList {
    pub entries: Vec<OverrideEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OverrideEntry {
    pub part_number: String,
    pub registered_name: String,
}

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

impl Default for AppState {
    fn default() -> Self {
        let settings = load_settings_from_disk().unwrap_or_default();
        let dictionary =
            load_dictionary_from_disk().unwrap_or_else(|_| default_column_dictionary());
        Self {
            bom_a: Mutex::new(None),
            bom_b: Mutex::new(None),
            comparison_result: Mutex::new(None),
            synthesis_result: Mutex::new(None),
            registered_name_list: Mutex::new(None),
            override_list: Mutex::new(None),
            file_a_path: Mutex::new(None),
            file_b_path: Mutex::new(None),
            column_mapping_a: Mutex::new(None),
            column_mapping_b: Mutex::new(None),
            settings: Mutex::new(settings),
            column_dictionary: Mutex::new(dictionary),
        }
    }
}

#[derive(Debug, Serialize)]
struct LoadFileResponse {
    message: String,
    side: String,
    preview: Option<PreviewTable>,
}

#[derive(Debug, Serialize)]
struct AnalyzeFileResponse {
    headers: Vec<String>,
    suggested_mapping: Option<ColumnMapping>,
    sample_rows: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
struct SessionListItem {
    id: String,
    label: Option<String>,
    created_at: String,
    file_a_name: Option<String>,
    file_b_name: Option<String>,
}

// 自動修正情報
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AutoCorrection {
    pub row_number: usize,
    pub column_index: usize,
    pub column_name: String,
    pub original_value: String,
    pub corrected_value: String,
    pub rule: String,
}

// 設定情報
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct FormatRule {
    pub pattern: String,
    pub action: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct AppSettings {
    pub makers: Vec<String>,
    pub format_rules: Vec<FormatRule>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ColumnDictionaryEntry {
    pub column_type: String,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub patterns: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ColumnDictionary {
    pub columns: Vec<ColumnDictionaryEntry>,
}

#[derive(Debug, Clone, Serialize)]
pub struct PreviewTable {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
    pub total_rows: usize,
}

impl ColumnDictionary {
    pub fn entry_for(&self, column_type: &str) -> Option<&ColumnDictionaryEntry> {
        let needle = column_type.trim().to_lowercase();
        self.columns
            .iter()
            .find(|entry| entry.column_type.trim().eq_ignore_ascii_case(&needle))
    }

    pub fn patterns_for(&self, column_type: &str) -> Vec<String> {
        self.entry_for(column_type)
            .map(|entry| {
                entry
                    .patterns
                    .iter()
                    .map(|p| p.trim().to_string())
                    .filter(|p| !p.is_empty())
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default()
    }
}

// ファイル読み込み結果
#[derive(Debug, Clone)]
pub struct LoadBomResult {
    pub bom: BomData,
    pub corrections: Vec<AutoCorrection>,
}

#[derive(Debug, Serialize)]
struct SessionRestoreResponse {
    message: String,
    file_a_path: Option<String>,
    file_b_path: Option<String>,
    column_mapping_a: Option<ColumnMapping>,
    column_mapping_b: Option<ColumnMapping>,
    comparison_result: Option<ComparisonResult>,
    synthesis_result: Option<SynthesisResult>,
    bom_a_headers: Option<Vec<String>>,
    bom_b_headers: Option<Vec<String>>,
}

#[derive(Debug, Serialize)]
struct CompareResponse {
    result: ComparisonResult,
    stats: HashMap<String, usize>,
}

#[derive(Debug, Deserialize)]
struct PreprocessRequest {
    side: Option<String>,
    bom_data: Option<BomSnapshot>,
    rules: PreprocessRules,
    persist: Option<bool>,
}

#[derive(Debug, Serialize)]
struct PreprocessResponse {
    bom_data: BomSnapshot,
}

#[derive(Debug, Deserialize)]
struct SetOverridesRequest {
    entry: Option<OverrideEntry>,
    entries: Option<Vec<OverrideEntry>>,
    remove_part_number: Option<String>,
    replace: Option<bool>,
}

#[derive(Debug, Serialize)]
struct OverrideListResponse {
    overrides: OverrideList,
    message: String,
}

#[derive(Debug, Serialize)]
struct RegisteredNameListResponse {
    list: RegisteredNameList,
    message: String,
}

#[derive(Debug, Serialize)]
struct MessageResponse {
    message: String,
}

// ファイル読み込みコマンド
#[tauri::command]
async fn load_file(
    file_path: String,
    column_mapping: ColumnMapping,
    side: String, // "a" or "b"
    state: State<'_, AppState>,
) -> Result<LoadFileResponse, String> {
    let side_normalized = side.to_lowercase();
    if side_normalized != "a" && side_normalized != "b" {
        return Err("無効なサイド指定です".to_string());
    }

    match bom_processor::load_bom_file(&file_path, &column_mapping).await {
        Ok(load_result) => {
            let bom_data = load_result.bom;

            let preview = match generate_preprocessed_preview(&bom_data, &column_mapping) {
                Ok(table) => Some(table),
                Err(err) => {
                    println!(
                        "[load_file][preview_error] side={}, path={}, err={}",
                        side_normalized, file_path, err
                    );
                    None
                }
            };

            println!("[load_file] side={}, path={}", side_normalized, file_path);
            if side_normalized == "a" {
                *state.bom_a.lock().unwrap() = Some(bom_data.clone());
                *state.file_a_path.lock().unwrap() = Some(file_path.clone());
                *state.column_mapping_a.lock().unwrap() = Some(column_mapping.clone());
            } else {
                *state.bom_b.lock().unwrap() = Some(bom_data.clone());
                *state.file_b_path.lock().unwrap() = Some(file_path.clone());
                *state.column_mapping_b.lock().unwrap() = Some(column_mapping.clone());
            }

            *state.comparison_result.lock().unwrap() = None;
            *state.synthesis_result.lock().unwrap() = None;

            save_auto_session(&state)?;

            Ok(LoadFileResponse {
                message: format!("部品表{}を読み込みました", side_normalized.to_uppercase()),
                side: side_normalized,
                preview,
            })
        }
        Err(e) => {
            println!(
                "[load_file][error] side={}, path={}, err={}",
                side_normalized, file_path, e
            );
            Err(format!("ファイル読み込みエラー: {}", e))
        }
    }
}

#[tauri::command]
async fn analyze_file(
    file_path: String,
    state: State<'_, AppState>,
) -> Result<AnalyzeFileResponse, String> {
    let dictionary = state.column_dictionary.lock().unwrap().clone();
    let analysis = bom_processor::analyze_bom_file(&file_path, &dictionary)
        .await
        .map_err(|e| format!("ファイル解析エラー: {e}"))?;

    Ok(AnalyzeFileResponse {
        headers: analysis.headers,
        suggested_mapping: analysis.suggested_mapping,
        sample_rows: analysis.sample_rows,
    })
}

#[tauri::command]
async fn preview_file(
    file_path: String,
    limit: Option<usize>,
) -> Result<bom_processor::FilePreview, String> {
    let row_limit = limit.unwrap_or(200);
    bom_processor::preview_raw_file(&file_path, row_limit)
        .await
        .map_err(|e| format!("プレビュー取得エラー: {e}"))
}

// 比較実行コマンド
fn fetch_boms(state: &State<'_, AppState>) -> Result<(BomData, BomData), String> {
    let bom_a = state
        .bom_a
        .lock()
        .map_err(|_| "部品表Aのロックに失敗しました".to_string())?
        .clone();
    let bom_b = state
        .bom_b
        .lock()
        .map_err(|_| "部品表Bのロックに失敗しました".to_string())?
        .clone();

    match (bom_a, bom_b) {
        (Some(a), Some(b)) => Ok((a, b)),
        _ => Err("部品表AまたはBが読み込まれていません".to_string()),
    }
}

fn get_bom_from_state(state: &State<'_, AppState>, side: &str) -> Result<Option<BomData>, String> {
    match side {
        "a" => state
            .bom_a
            .lock()
            .map_err(|_| "部品表Aのロックに失敗しました".to_string())
            .map(|guard| guard.clone()),
        "b" => state
            .bom_b
            .lock()
            .map_err(|_| "部品表Bのロックに失敗しました".to_string())
            .map(|guard| guard.clone()),
        _ => Err("サイド指定が無効です".to_string()),
    }
}

#[tauri::command]
async fn compare_boms(state: State<'_, AppState>) -> Result<ComparisonResult, String> {
    let (a, b) = fetch_boms(&state)?;
    let result = perform_comparison(&a, &b);
    *state.comparison_result.lock().unwrap() = Some(result.clone());
    Ok(result)
}

#[tauri::command]
async fn compare_with_comments(state: State<'_, AppState>) -> Result<CompareResponse, String> {
    let (a, b) = fetch_boms(&state)?;
    let result = perform_comparison(&a, &b);
    let stats = get_comparison_stats(&result);
    *state.comparison_result.lock().unwrap() = Some(result.clone());
    Ok(CompareResponse { result, stats })
}

// 合成実行コマンド
#[tauri::command]
async fn synthesize_boms(state: State<'_, AppState>) -> Result<SynthesisResult, String> {
    let (bom_a, bom_b) = {
        let bom_a_guard = state.bom_a.lock().unwrap();
        let bom_b_guard = state.bom_b.lock().unwrap();
        (bom_a_guard.clone(), bom_b_guard.clone())
    };

    match (bom_a, bom_b) {
        (Some(a), Some(b)) => {
            let result = perform_synthesis(&a, &b);
            *state.synthesis_result.lock().unwrap() = Some(result.clone());
            Ok(result)
        }
        _ => Err("部品表AまたはBが読み込まれていません".to_string()),
    }
}

#[tauri::command]
async fn preprocess_bom(
    request: PreprocessRequest,
    state: State<'_, AppState>,
) -> Result<PreprocessResponse, String> {
    let side = request.side.as_ref().map(|s| s.to_lowercase());
    let persist = request.persist.unwrap_or(side.is_some());

    if persist && side.is_none() {
        return Err("前処理結果を保存する場合は対象サイドを指定してください".to_string());
    }

    let maybe_bom = if let Some(snapshot) = request.bom_data {
        Some(BomData::from(snapshot))
    } else if let Some(ref side_key) = side {
        get_bom_from_state(&state, side_key)?
    } else {
        None
    };

    let source_bom = maybe_bom.ok_or_else(|| "前処理対象の部品表がありません".to_string())?;

    let processed_bom = bom_processor::preprocess_bom_data(&source_bom, &request.rules)
        .map_err(|e| format!("前処理エラー: {e}"))?;

    if persist {
        if let Some(ref side_key) = side {
            match side_key.as_str() {
                "a" => {
                    *state.bom_a.lock().unwrap() = Some(processed_bom.clone());
                }
                "b" => {
                    *state.bom_b.lock().unwrap() = Some(processed_bom.clone());
                }
                _ => return Err("サイド指定が無効です".to_string()),
            }
            *state.comparison_result.lock().unwrap() = None;
            save_auto_session(&state)?;
        }
    }

    Ok(PreprocessResponse {
        bom_data: BomSnapshot::from(processed_bom),
    })
}

#[tauri::command]
async fn update_bom_data(
    side: String,
    bom_data: BomSnapshot,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let side_key = side.to_lowercase();
    let bom: BomData = bom_data.into();

    match side_key.as_str() {
        "a" => {
            *state.bom_a.lock().unwrap() = Some(bom);
        }
        "b" => {
            *state.bom_b.lock().unwrap() = Some(bom);
        }
        _ => return Err("サイド指定が無効です".to_string()),
    }

    *state.comparison_result.lock().unwrap() = None;
    save_auto_session(&state)?;

    Ok(MessageResponse {
        message: format!("部品表{}を更新しました", side_key.to_uppercase()),
    })
}

#[tauri::command(name = "load_registered_name_list")]
async fn load_registered_name_list_cmd(
    file_path: String,
    format: String,
    state: State<'_, AppState>,
) -> Result<RegisteredNameListResponse, String> {
    let format_norm = format.to_lowercase();
    let list = match format_norm.as_str() {
        "csv" => bom_processor::load_registered_name_csv(&file_path)
            .await
            .map_err(|e| format!("CSV読み込みエラー: {e}"))?,
        "json" => bom_processor::load_registered_name_json(&file_path)
            .await
            .map_err(|e| format!("JSON読み込みエラー: {e}"))?,
        _ => return Err("サポートされていないフォーマットです".to_string()),
    };

    *state.registered_name_list.lock().unwrap() = Some(list.clone());
    save_auto_session(&state)?;

    Ok(RegisteredNameListResponse {
        list,
        message: "登録名リストを読み込みました".to_string(),
    })
}

#[tauri::command(name = "save_registered_name_list")]
async fn save_registered_name_list_cmd(
    file_path: String,
    format: String,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let list = state
        .registered_name_list
        .lock()
        .unwrap()
        .clone()
        .ok_or_else(|| "登録名リストがありません".to_string())?;

    let format_norm = format.to_lowercase();
    match format_norm.as_str() {
        "csv" => bom_processor::save_registered_name_csv(&list, &file_path)
            .await
            .map_err(|e| format!("CSV保存エラー: {e}"))?,
        "json" => bom_processor::save_registered_name_json(&list, &file_path)
            .await
            .map_err(|e| format!("JSON保存エラー: {e}"))?,
        _ => return Err("サポートされていないフォーマットです".to_string()),
    }

    Ok(MessageResponse {
        message: "登録名リストを保存しました".to_string(),
    })
}

#[tauri::command]
async fn apply_registered_names(
    side: String,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let side_key = side.to_lowercase();
    let registered_list = state.registered_name_list.lock().unwrap().clone();
    let overrides = state.override_list.lock().unwrap().clone();

    match side_key.as_str() {
        "a" => {
            let mut bom_lock = state.bom_a.lock().unwrap();
            if let Some(ref mut bom) = *bom_lock {
                bom_processor::apply_registered_names_to_bom(bom, &registered_list, &overrides);
            } else {
                return Err("部品表Aが読み込まれていません".to_string());
            }
        }
        "b" => {
            let mut bom_lock = state.bom_b.lock().unwrap();
            if let Some(ref mut bom) = *bom_lock {
                bom_processor::apply_registered_names_to_bom(bom, &registered_list, &overrides);
            } else {
                return Err("部品表Bが読み込まれていません".to_string());
            }
        }
        _ => return Err("サイド指定が無効です".to_string()),
    }

    *state.comparison_result.lock().unwrap() = None;
    save_auto_session(&state)?;

    Ok(MessageResponse {
        message: format!("部品表{}に登録名を適用しました", side_key.to_uppercase()),
    })
}

#[tauri::command]
async fn load_settings(state: State<'_, AppState>) -> Result<AppSettings, String> {
    let settings = state.settings.lock().unwrap().clone();
    Ok(settings)
}

#[tauri::command]
async fn save_settings(
    settings: AppSettings,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let normalized = normalize_settings(settings)?;
    write_settings_to_disk(&normalized)?;
    *state.settings.lock().unwrap() = normalized;

    Ok(MessageResponse {
        message: "設定を保存しました".to_string(),
    })
}

#[tauri::command]
async fn import_settings(
    file_path: String,
    state: State<'_, AppState>,
) -> Result<AppSettings, String> {
    let path = Path::new(&file_path);
    if !path.exists() {
        return Err("設定ファイルが見つかりません".to_string());
    }

    let content = fs::read_to_string(path)
        .map_err(|e| format!("設定ファイルの読み込みに失敗しました: {e}"))?;

    let raw: AppSettings = serde_json::from_str(&content)
        .map_err(|e| format!("設定ファイルの解析に失敗しました: {e}"))?;

    let normalized = normalize_settings(raw)?;
    write_settings_to_disk(&normalized)?;
    *state.settings.lock().unwrap() = normalized.clone();

    Ok(normalized)
}

#[tauri::command]
async fn export_settings(
    file_path: String,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let settings = state.settings.lock().unwrap().clone();
    let json = serde_json::to_string_pretty(&settings)
        .map_err(|e| format!("設定JSONの生成に失敗しました: {e}"))?;

    let path = Path::new(&file_path);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("ディレクトリの作成に失敗しました: {e}"))?;
    }

    fs::write(path, json).map_err(|e| format!("設定ファイルの書き込みに失敗しました: {e}"))?;

    Ok(MessageResponse {
        message: format!("設定をエクスポートしました: {}", file_path),
    })
}

#[tauri::command]
async fn load_column_dictionary(state: State<'_, AppState>) -> Result<ColumnDictionary, String> {
    Ok(state.column_dictionary.lock().unwrap().clone())
}

#[tauri::command]
async fn save_column_dictionary(
    dictionary: ColumnDictionary,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let normalized = normalize_dictionary(dictionary)?;
    write_dictionary_to_disk(&normalized)?;
    *state.column_dictionary.lock().unwrap() = normalized;

    Ok(MessageResponse {
        message: "辞書を保存しました".to_string(),
    })
}

#[tauri::command]
async fn import_column_dictionary(
    file_path: String,
    state: State<'_, AppState>,
) -> Result<ColumnDictionary, String> {
    let path = Path::new(&file_path);
    if !path.exists() {
        return Err("辞書ファイルが見つかりません".to_string());
    }

    let content = fs::read_to_string(path)
        .map_err(|e| format!("辞書ファイルの読み込みに失敗しました: {e}"))?;

    let raw: ColumnDictionary = serde_json::from_str(&content)
        .map_err(|e| format!("辞書ファイルの解析に失敗しました: {e}"))?;

    let normalized = normalize_dictionary(raw)?;
    write_dictionary_to_disk(&normalized)?;
    *state.column_dictionary.lock().unwrap() = normalized.clone();

    Ok(normalized)
}

#[tauri::command]
async fn export_column_dictionary(
    file_path: String,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let dictionary = state.column_dictionary.lock().unwrap().clone();
    let json = serde_json::to_string_pretty(&dictionary)
        .map_err(|e| format!("辞書JSONの生成に失敗しました: {e}"))?;

    let path = Path::new(&file_path);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("ディレクトリの作成に失敗しました: {e}"))?;
    }

    fs::write(path, json).map_err(|e| format!("辞書ファイルの書き込みに失敗しました: {e}"))?;

    Ok(MessageResponse {
        message: format!("辞書をエクスポートしました: {}", file_path),
    })
}

#[tauri::command]
async fn get_processed_preview(
    side: String,
    state: State<'_, AppState>,
) -> Result<PreviewTable, String> {
    let side_key = side.to_lowercase();
    let bom = match side_key.as_str() {
        "a" => state
            .bom_a
            .lock()
            .map_err(|_| "部品表Aのロックに失敗しました".to_string())?
            .clone(),
        "b" => state
            .bom_b
            .lock()
            .map_err(|_| "部品表Bのロックに失敗しました".to_string())?
            .clone(),
        _ => return Err("無効なサイド指定です".to_string()),
    };

    let mapping = match side_key.as_str() {
        "a" => state
            .column_mapping_a
            .lock()
            .map_err(|_| "列指定Aのロックに失敗しました".to_string())?
            .clone(),
        "b" => state
            .column_mapping_b
            .lock()
            .map_err(|_| "列指定Bのロックに失敗しました".to_string())?
            .clone(),
        _ => None,
    };

    let bom = bom.ok_or_else(|| match side_key.as_str() {
        "a" => "部品表Aが読み込まれていません".to_string(),
        "b" => "部品表Bが読み込まれていません".to_string(),
        _ => "無効なサイド指定です".to_string(),
    })?;

    let mapping = mapping.ok_or_else(|| match side_key.as_str() {
        "a" => "部品表Aの列設定が未指定です".to_string(),
        "b" => "部品表Bの列設定が未指定です".to_string(),
        _ => "無効なサイド指定です".to_string(),
    })?;

    generate_preprocessed_preview(&bom, &mapping)
}

#[tauri::command]
async fn set_overrides(
    request: SetOverridesRequest,
    state: State<'_, AppState>,
) -> Result<OverrideListResponse, String> {
    let mut guard = state.override_list.lock().unwrap();
    let mut overrides = guard.clone().unwrap_or_default();

    if let Some(part_number) = request.remove_part_number.as_ref() {
        overrides
            .entries
            .retain(|entry| &entry.part_number != part_number);
    }

    if let Some(entries) = request.entries {
        if request.replace.unwrap_or(true) {
            overrides.entries = entries;
        } else {
            for entry in entries {
                upsert_override_entry(&mut overrides, entry);
            }
        }
    }

    if let Some(entry) = request.entry {
        upsert_override_entry(&mut overrides, entry);
    }

    overrides
        .entries
        .sort_by(|a, b| a.part_number.cmp(&b.part_number));

    *guard = Some(overrides.clone());
    save_auto_session(&state)?;

    Ok(OverrideListResponse {
        overrides,
        message: "上書きリストを更新しました".to_string(),
    })
}

fn upsert_override_entry(list: &mut OverrideList, entry: OverrideEntry) {
    if let Some(existing) = list
        .entries
        .iter_mut()
        .find(|e| e.part_number == entry.part_number)
    {
        existing.registered_name = entry.registered_name;
    } else {
        list.entries.push(entry);
    }
}

#[tauri::command]
async fn apply_overrides_ipc(
    side: String,
    state: State<'_, AppState>,
) -> Result<MessageResponse, String> {
    let side_key = side.to_lowercase();
    let registered_list = state.registered_name_list.lock().unwrap().clone();
    let overrides = state.override_list.lock().unwrap().clone();

    match side_key.as_str() {
        "a" => {
            let mut bom_lock = state.bom_a.lock().unwrap();
            if let Some(ref mut bom) = *bom_lock {
                bom_processor::apply_registered_names_to_bom(bom, &registered_list, &overrides);
            } else {
                return Err("部品表Aが読み込まれていません".to_string());
            }
        }
        "b" => {
            let mut bom_lock = state.bom_b.lock().unwrap();
            if let Some(ref mut bom) = *bom_lock {
                bom_processor::apply_registered_names_to_bom(bom, &registered_list, &overrides);
            } else {
                return Err("部品表Bが読み込まれていません".to_string());
            }
        }
        _ => return Err("サイド指定が無効です".to_string()),
    }

    *state.comparison_result.lock().unwrap() = None;
    save_auto_session(&state)?;

    Ok(MessageResponse {
        message: format!("部品表{}に上書きを適用しました", side_key.to_uppercase()),
    })
}

#[tauri::command(name = "get_registered_name_list")]
async fn get_registered_name_list_cmd(
    state: State<'_, AppState>,
) -> Result<Option<RegisteredNameList>, String> {
    Ok(state.registered_name_list.lock().unwrap().clone())
}

#[tauri::command(name = "get_override_list")]
async fn get_override_list_cmd(state: State<'_, AppState>) -> Result<Option<OverrideList>, String> {
    Ok(state.override_list.lock().unwrap().clone())
}

#[tauri::command]
async fn validate_bom_data(
    side: Option<String>,
    bom_data: Option<BomSnapshot>,
    state: State<'_, AppState>,
) -> Result<ValidationResult, String> {
    let bom = if let Some(snapshot) = bom_data {
        BomData::from(snapshot)
    } else if let Some(side_value) = side {
        let side_key = side_value.to_lowercase();
        get_bom_from_state(&state, &side_key)?
            .ok_or_else(|| format!("部品表{}が読み込まれていません", side_key.to_uppercase()))?
    } else {
        return Err("バリデーション対象の部品表が指定されていません".to_string());
    };

    Ok(bom_processor::validate_bom_data(&bom))
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
            let comparison = state.comparison_result.lock().unwrap().clone();
            match comparison {
                Some(result) => save_comparison_result(&result, &file_path, &format).await,
                None => Err("比較結果がありません".to_string()),
            }
        }
        "synthesis" => {
            let synthesis = state.synthesis_result.lock().unwrap().clone();
            match synthesis {
                Some(result) => save_synthesis_result(&result, &file_path, &format).await,
                None => Err("合成結果がありません".to_string()),
            }
        }
        _ => Err("無効な結果タイプです".to_string()),
    }
}

#[tauri::command]
async fn clear_data(mode: String, state: State<'_, AppState>) -> Result<MessageResponse, String> {
    match mode.to_lowercase().as_str() {
        "all" => {
            *state.bom_a.lock().unwrap() = None;
            *state.bom_b.lock().unwrap() = None;
            *state.comparison_result.lock().unwrap() = None;
            *state.synthesis_result.lock().unwrap() = None;
            *state.registered_name_list.lock().unwrap() = None;
            *state.override_list.lock().unwrap() = None;
            *state.file_a_path.lock().unwrap() = None;
            *state.file_b_path.lock().unwrap() = None;
            *state.column_mapping_a.lock().unwrap() = None;
            *state.column_mapping_b.lock().unwrap() = None;
            save_auto_session(&state)?;
            Ok(MessageResponse {
                message: "全データをクリアしました".to_string(),
            })
        }
        "session_keep" => {
            *state.bom_a.lock().unwrap() = None;
            *state.bom_b.lock().unwrap() = None;
            *state.comparison_result.lock().unwrap() = None;
            *state.synthesis_result.lock().unwrap() = None;
            *state.file_a_path.lock().unwrap() = None;
            *state.file_b_path.lock().unwrap() = None;
            *state.column_mapping_a.lock().unwrap() = None;
            *state.column_mapping_b.lock().unwrap() = None;
            save_auto_session(&state)?;
            Ok(MessageResponse {
                message: "登録名と上書きを保持してクリアしました".to_string(),
            })
        }
        _ => Err("無効なクリアモードです".to_string()),
    }
}

// シートクリアコマンド（後方互換）
#[tauri::command]
async fn clear_sheets(state: State<'_, AppState>) -> Result<String, String> {
    clear_data("all".to_string(), state)
        .await
        .map(|resp| resp.message)
}

#[tauri::command]
async fn list_sessions(kind: String) -> Result<Vec<SessionListItem>, String> {
    let kind_enum = parse_session_kind(&kind)?;
    let summaries = collect_snapshots(kind_enum)?;
    Ok(summaries
        .into_iter()
        .map(|summary| SessionListItem {
            id: summary.id,
            label: summary.label,
            created_at: summary.created_at.to_rfc3339(),
            file_a_name: summary.file_a_name,
            file_b_name: summary.file_b_name,
        })
        .collect())
}

#[tauri::command]
async fn save_manual_session(
    label: Option<String>,
    state: State<'_, AppState>,
) -> Result<Vec<SessionListItem>, String> {
    let cleaned_label = label.and_then(|l| {
        let trimmed = l.trim().to_string();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed)
        }
    });
    let snapshot = create_snapshot(&state, true, cleaned_label);
    let _ = save_snapshot(snapshot, SessionKind::Manual)?;
    list_sessions("manual".to_string()).await
}

#[tauri::command]
async fn restore_session(
    kind: String,
    id: String,
    state: State<'_, AppState>,
) -> Result<SessionRestoreResponse, String> {
    let kind_enum = parse_session_kind(&kind)?;
    let snapshot = load_snapshot(kind_enum, &id)?;
    apply_snapshot(&state, &snapshot);

    Ok(SessionRestoreResponse {
        message: "セッションを復元しました".to_string(),
        file_a_path: snapshot.file_a_path.clone(),
        file_b_path: snapshot.file_b_path.clone(),
        column_mapping_a: snapshot.column_mapping_a.clone(),
        column_mapping_b: snapshot.column_mapping_b.clone(),
        comparison_result: snapshot.comparison_result.clone(),
        synthesis_result: snapshot.synthesis_result.clone(),
        bom_a_headers: snapshot.bom_a.as_ref().map(|b| b.headers.clone()),
        bom_b_headers: snapshot.bom_b.as_ref().map(|b| b.headers.clone()),
    })
}

#[tauri::command]
async fn delete_session_command(kind: String, id: String) -> Result<Vec<SessionListItem>, String> {
    let kind_enum = parse_session_kind(&kind)?;
    delete_snapshot(kind_enum, &id)?;
    list_sessions(kind).await
}

fn main() {
    ensure_watcher_ignore();
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .manage(AppState::default())
        .invoke_handler(tauri::generate_handler![
            open_file_dialog,
            load_file,
            analyze_file,
            preview_file,
            compare_boms,
            compare_with_comments,
            synthesize_boms,
            preprocess_bom,
            update_bom_data,
            save_result,
            load_registered_name_list_cmd,
            save_registered_name_list_cmd,
            apply_registered_names,
            set_overrides,
            apply_overrides_ipc,
            get_registered_name_list_cmd,
            get_override_list_cmd,
            validate_bom_data,
            load_settings,
            save_settings,
            import_settings,
            export_settings,
            load_column_dictionary,
            save_column_dictionary,
            import_column_dictionary,
            export_column_dictionary,
            get_processed_preview,
            clear_sheets,
            clear_data,
            list_sessions,
            save_manual_session,
            restore_session,
            delete_session_command,
            log_client_event,
            generate_cad_file,
            get_bom_snapshot,
            save_file_dialog,
            open_settings_import_dialog
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

fn ensure_watcher_ignore() {
    const IGNORE_ENTRY: &str = "sessions/**";
    const IGNORE_ENTRY_PARENT: &str = "../sessions/**";
    match env::var("TAURI_DEV_WATCHER_IGNORE") {
        Ok(current) => {
            let mut entries: Vec<String> = current
                .split(';')
                .map(|entry| entry.trim().to_string())
                .filter(|entry| !entry.is_empty())
                .collect();
            if !entries
                .iter()
                .any(|entry| entry.eq_ignore_ascii_case(IGNORE_ENTRY))
            {
                entries.push(IGNORE_ENTRY.to_string());
            }
            if !entries
                .iter()
                .any(|entry| entry.eq_ignore_ascii_case(IGNORE_ENTRY_PARENT))
            {
                entries.push(IGNORE_ENTRY_PARENT.to_string());
            }
            let new_value = entries.join(";");
            env::set_var("TAURI_DEV_WATCHER_IGNORE", new_value);
        }
        Err(_) => {
            let value = format!("{};{}", IGNORE_ENTRY, IGNORE_ENTRY_PARENT);
            env::set_var("TAURI_DEV_WATCHER_IGNORE", value);
        }
    }
}

#[tauri::command]
async fn open_file_dialog(app: tauri::AppHandle) -> Result<Option<String>, String> {
    let (tx, rx) = oneshot::channel();
    app.dialog()
        .file()
        .set_title("部品表ファイルを選択")
        .set_can_create_directories(false)
        .add_filter("BOM ファイル", &["csv", "xls", "xlsx"])
        .pick_file(move |file| {
            let path = file
                .and_then(|fp| fp.into_path().ok())
                .map(|p| p.to_string_lossy().into_owned());
            let _ = tx.send(path);
        });

    match rx.await {
        Ok(path) => Ok(path),
        Err(_) => Ok(None),
    }
}

#[tauri::command]
async fn open_settings_import_dialog(app: tauri::AppHandle) -> Result<Option<String>, String> {
    let (tx, rx) = oneshot::channel();
    app.dialog()
        .file()
        .set_title("設定ファイルを選択")
        .set_can_create_directories(false)
        .add_filter("設定ファイル", &["json"])
        .pick_file(move |file| {
            let path = file
                .and_then(|fp| fp.into_path().ok())
                .map(|p| p.to_string_lossy().into_owned());
            let _ = tx.send(path);
        });

    match rx.await {
        Ok(path) => Ok(path),
        Err(_) => Ok(None),
    }
}

fn parse_session_kind(kind: &str) -> Result<SessionKind, String> {
    match kind.to_lowercase().as_str() {
        "auto" => Ok(SessionKind::Auto),
        "manual" => Ok(SessionKind::Manual),
        _ => Err("不明なセッション種別です".to_string()),
    }
}

fn create_snapshot(
    state: &AppState,
    include_results: bool,
    label: Option<String>,
) -> SessionSnapshot {
    let bom_a = state.bom_a.lock().unwrap().clone();
    let bom_b = state.bom_b.lock().unwrap().clone();
    let comparison = if include_results {
        state.comparison_result.lock().unwrap().clone()
    } else {
        None
    };
    let synthesis = if include_results {
        state.synthesis_result.lock().unwrap().clone()
    } else {
        None
    };
    let registered_name_list = state.registered_name_list.lock().unwrap().clone();
    let override_list = state.override_list.lock().unwrap().clone();

    SessionSnapshot {
        id: String::new(),
        label,
        created_at: Utc::now(),
        file_a_path: state.file_a_path.lock().unwrap().clone(),
        file_b_path: state.file_b_path.lock().unwrap().clone(),
        column_mapping_a: state.column_mapping_a.lock().unwrap().clone(),
        column_mapping_b: state.column_mapping_b.lock().unwrap().clone(),
        bom_a,
        bom_b,
        comparison_result: comparison,
        synthesis_result: synthesis,
        registered_name_list,
        override_list,
    }
}

fn apply_snapshot(state: &AppState, snapshot: &SessionSnapshot) {
    *state.bom_a.lock().unwrap() = snapshot.bom_a.clone();
    *state.bom_b.lock().unwrap() = snapshot.bom_b.clone();
    *state.file_a_path.lock().unwrap() = snapshot.file_a_path.clone();
    *state.file_b_path.lock().unwrap() = snapshot.file_b_path.clone();
    *state.column_mapping_a.lock().unwrap() = snapshot.column_mapping_a.clone();
    *state.column_mapping_b.lock().unwrap() = snapshot.column_mapping_b.clone();
    *state.comparison_result.lock().unwrap() = snapshot.comparison_result.clone();
    *state.synthesis_result.lock().unwrap() = snapshot.synthesis_result.clone();
    *state.registered_name_list.lock().unwrap() = snapshot.registered_name_list.clone();
    *state.override_list.lock().unwrap() = snapshot.override_list.clone();
}

fn save_auto_session(state: &AppState) -> Result<(), String> {
    let bom_a_exists = state.bom_a.lock().unwrap().is_some();
    let bom_b_exists = state.bom_b.lock().unwrap().is_some();
    if !bom_a_exists && !bom_b_exists {
        return Ok(());
    }

    let snapshot = create_snapshot(state, false, None);
    let _ = save_snapshot(snapshot, SessionKind::Auto)?;
    Ok(())
}

fn settings_file_path() -> PathBuf {
    Path::new(SETTINGS_DIR).join(SETTINGS_FILE_NAME)
}

fn load_settings_from_disk() -> Result<AppSettings, String> {
    let path = settings_file_path();
    if !path.exists() {
        return Ok(AppSettings::default());
    }

    let content = fs::read_to_string(&path)
        .map_err(|e| format!("設定ファイルの読み込みに失敗しました: {e}"))?;

    if content.trim().is_empty() {
        return Ok(AppSettings::default());
    }

    let raw: AppSettings = serde_json::from_str(&content)
        .map_err(|e| format!("設定ファイルの解析に失敗しました: {e}"))?;

    normalize_settings(raw)
}

fn normalize_settings(settings: AppSettings) -> Result<AppSettings, String> {
    let mut makers = Vec::new();
    let mut maker_seen: HashSet<String> = HashSet::new();

    for maker in settings.makers.into_iter() {
        let trimmed = maker.trim().to_string();
        if trimmed.is_empty() {
            return Err("メーカー名に空の値は使用できません".to_string());
        }
        if maker_seen.insert(trimmed.clone()) {
            makers.push(trimmed);
        }
    }

    let mut rules = Vec::new();
    let mut rule_seen: HashSet<(String, String)> = HashSet::new();

    for rule in settings.format_rules.into_iter() {
        let pattern = rule.pattern.trim().to_string();
        let action = rule.action.trim().to_lowercase();
        if !SETTINGS_ACTIONS.contains(&action.as_str()) {
            return Err(format!("無効な処理方法です: {action}"));
        }
        let key = (pattern.clone(), action.clone());
        if rule_seen.insert(key) {
            rules.push(FormatRule { pattern, action });
        }
    }

    Ok(AppSettings {
        makers,
        format_rules: rules,
    })
}

fn write_settings_to_disk(settings: &AppSettings) -> Result<(), String> {
    let path = settings_file_path();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("設定フォルダの作成に失敗しました: {e}"))?;
    }

    let json = serde_json::to_string_pretty(settings)
        .map_err(|e| format!("設定JSONの生成に失敗しました: {e}"))?;

    fs::write(&path, json).map_err(|e| format!("設定ファイルの保存に失敗しました: {e}"))?;

    Ok(())
}

fn dictionary_file_path() -> PathBuf {
    Path::new(DICTIONARY_DIR).join(DICTIONARY_FILE_NAME)
}

fn load_dictionary_from_disk() -> Result<ColumnDictionary, String> {
    let path = dictionary_file_path();
    if !path.exists() {
        let defaults = default_column_dictionary();
        let _ = write_dictionary_to_disk(&defaults);
        return Ok(defaults);
    }

    let content = fs::read_to_string(&path)
        .map_err(|e| format!("辞書ファイルの読み込みに失敗しました: {e}"))?;

    if content.trim().is_empty() {
        return Ok(default_column_dictionary());
    }

    let raw: ColumnDictionary = serde_json::from_str(&content)
        .map_err(|e| format!("辞書ファイルの解析に失敗しました: {e}"))?;

    normalize_dictionary(raw)
}

fn normalize_dictionary(dictionary: ColumnDictionary) -> Result<ColumnDictionary, String> {
    let mut merged: BTreeMap<String, ColumnDictionaryEntry> = BTreeMap::new();

    for entry in dictionary.columns.into_iter() {
        let key = entry.column_type.trim().to_lowercase();
        if key.is_empty() {
            return Err("列タイプ名は必須です".to_string());
        }

        let display_name = entry.display_name.and_then(|name| {
            let trimmed = name.trim().to_string();
            if trimmed.is_empty() {
                None
            } else {
                Some(trimmed)
            }
        });

        let mut patterns_set: HashSet<String> = HashSet::new();
        for pattern in entry.patterns.into_iter() {
            let trimmed = pattern.trim().to_string();
            if trimmed.is_empty() {
                continue;
            }
            patterns_set.insert(trimmed);
        }

        let target = merged
            .entry(key.clone())
            .or_insert_with(|| ColumnDictionaryEntry {
                column_type: key.clone(),
                display_name: display_name.clone(),
                patterns: Vec::new(),
            });

        if display_name.is_some() {
            target.display_name = display_name.clone();
        }

        for pattern in patterns_set {
            if !target
                .patterns
                .iter()
                .any(|existing| existing.eq_ignore_ascii_case(&pattern))
            {
                target.patterns.push(pattern);
            }
        }
    }

    let mut columns: Vec<ColumnDictionaryEntry> = merged
        .into_values()
        .map(|mut entry| {
            entry.patterns.sort();
            entry
        })
        .collect();

    if columns.is_empty() {
        columns = default_column_dictionary().columns;
    }

    Ok(ColumnDictionary { columns })
}

fn write_dictionary_to_disk(dictionary: &ColumnDictionary) -> Result<(), String> {
    let path = dictionary_file_path();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("辞書フォルダの作成に失敗しました: {e}"))?;
    }

    let json = serde_json::to_string_pretty(dictionary)
        .map_err(|e| format!("辞書JSONの生成に失敗しました: {e}"))?;

    fs::write(&path, json).map_err(|e| format!("辞書ファイルの保存に失敗しました: {e}"))?;

    Ok(())
}

fn default_column_dictionary() -> ColumnDictionary {
    ColumnDictionary {
        columns: vec![
            ColumnDictionaryEntry {
                column_type: "part_number".to_string(),
                display_name: Some("部品番号".to_string()),
                patterns: vec![
                    "部品番号".to_string(),
                    "品番".to_string(),
                    "部番".to_string(),
                    "part number".to_string(),
                    "part no".to_string(),
                    "part#".to_string(),
                    "reference".to_string(),
                    "refdes".to_string(),
                ],
            },
            ColumnDictionaryEntry {
                column_type: "model_number".to_string(),
                display_name: Some("型番/部品名".to_string()),
                patterns: vec![
                    "型番".to_string(),
                    "部品名".to_string(),
                    "品名".to_string(),
                    "item".to_string(),
                    "part name".to_string(),
                    "description".to_string(),
                ],
            },
            ColumnDictionaryEntry {
                column_type: "manufacturer".to_string(),
                display_name: Some("メーカー".to_string()),
                patterns: vec![
                    "ﾒｰｶｰ".to_string(),
                    "メーカー".to_string(),
                    "maker".to_string(),
                    "manufacturer".to_string(),
                    "vendor".to_string(),
                ],
            },
        ],
    }
}

fn generate_preprocessed_preview(
    bom: &BomData,
    column_mapping: &ColumnMapping,
) -> Result<PreviewTable, String> {
    let default_rules = PreprocessRules {
        remove_parentheses: true,
        expand_ranges: true,
        fullwidth_to_halfwidth: true,
        lowercase_to_uppercase: true,
    };

    let processed = bom_processor::preprocess_bom_data(bom, &default_rules)
        .map_err(|e| format!("前処理エラー: {e}"))?;

    let headers = if processed.headers.is_empty() {
        (0..3)
            .map(|idx| format!("列{}", idx + 1))
            .collect::<Vec<_>>()
    } else {
        processed.headers.clone()
    };

    let total_rows = processed.rows.len();
    let limit = AUTO_PREVIEW_LIMIT.min(total_rows);

    let part_header = headers.get(column_mapping.part_number).cloned();
    let model_header = headers.get(column_mapping.model_number).cloned();
    let manufacturer_header = column_mapping
        .manufacturer
        .and_then(|idx| headers.get(idx).cloned());

    let mut rows = Vec::new();
    for row in processed.rows.iter().take(limit) {
        let mut line = Vec::with_capacity(headers.len());
        for header in headers.iter() {
            let mut value = row.attributes.get(header).cloned().unwrap_or_default();

            if value.is_empty() {
                if part_header.as_ref().map(String::as_str) == Some(header.as_str()) {
                    value = row.part_number.clone();
                } else if model_header.as_ref().map(String::as_str) == Some(header.as_str()) {
                    value = row.model_number.clone();
                } else if manufacturer_header
                    .as_ref()
                    .map(String::as_str)
                    == Some(header.as_str())
                {
                    value = row.attributes.get(header).cloned().unwrap_or_default();
                }
            }

            line.push(value);
        }
        rows.push(line);
    }

    Ok(PreviewTable {
        headers,
        rows,
        total_rows,
    })
}

#[tauri::command]
async fn log_client_event(level: String, message: String) -> Result<(), String> {
    println!("[client {level}] {message}");
    Ok(())
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BomSnapshot {
    pub headers: Vec<String>,
    pub rows: Vec<BomRow>,
}

impl From<BomData> for BomSnapshot {
    fn from(value: BomData) -> Self {
        Self {
            headers: value.headers,
            rows: value.rows,
        }
    }
}

impl From<BomSnapshot> for BomData {
    fn from(value: BomSnapshot) -> Self {
        Self {
            headers: value.headers,
            rows: value.rows,
        }
    }
}

#[tauri::command]
async fn get_bom_snapshot(
    side: String,
    state: State<'_, AppState>,
) -> Result<Option<BomSnapshot>, String> {
    let side = side.to_lowercase();
    let snapshot = match side.as_str() {
        "a" => state
            .bom_a
            .lock()
            .map_err(|_| "部品表Aのロックに失敗しました".to_string())?
            .clone()
            .map(BomSnapshot::from),
        "b" => state
            .bom_b
            .lock()
            .map_err(|_| "部品表Bのロックに失敗しました".to_string())?
            .clone()
            .map(BomSnapshot::from),
        _ => {
            return Err("サイド指定が無効です".to_string());
        }
    };
    Ok(snapshot)
}

#[tauri::command]
async fn generate_cad_file(
    format: String,
    snapshot: BomSnapshot,
    output_path: Option<String>,
) -> Result<String, String> {
    let format = CadFormat::parse(&format)?;
    let bom: BomData = snapshot.into();
    if bom.rows.is_empty() {
        return Err("出力対象の部品表にデータがありません".to_string());
    }

    let content = build_cad_output(&format, &bom);
    let target_path = determine_cad_output_path(&format, output_path)?;
    if let Some(parent) = target_path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("出力ディレクトリを作成できません: {e}"))?;
    }
    let mut file =
        File::create(&target_path).map_err(|e| format!("CADファイルを作成できません: {e}"))?;
    file.write_all(content.join("\n").as_bytes())
        .map_err(|e| format!("CADファイルの書き込みに失敗しました: {e}"))?;

    Ok(target_path.to_string_lossy().to_string())
}

#[derive(Debug, Clone, Copy)]
enum CadFormat {
    Pads,
    Bd,
    Pws,
}

impl CadFormat {
    fn parse(input: &str) -> Result<Self, String> {
        match input.trim().to_uppercase().as_str() {
            "PADS" => Ok(CadFormat::Pads),
            "BD" => Ok(CadFormat::Bd),
            "PWS" => Ok(CadFormat::Pws),
            other => Err(format!("未対応のCADフォーマットです: {other}")),
        }
    }

    fn default_extension(&self) -> &'static str {
        match self {
            CadFormat::Pads => "pads",
            CadFormat::Bd => "bd",
            CadFormat::Pws => "pws",
        }
    }

    fn display_name(&self) -> &'static str {
        match self {
            CadFormat::Pads => "PADS",
            CadFormat::Bd => "BD",
            CadFormat::Pws => "PWS",
        }
    }
}

fn build_cad_output(format: &CadFormat, bom: &BomData) -> Vec<String> {
    let mut lines = Vec::new();
    match format {
        CadFormat::Pads => {
            lines.push("!KYODEN BOM TOOL CAD EXPORT - PADS".to_string());
            lines.push("PART_NUMBER\tMODEL_NUMBER".to_string());
            for row in &bom.rows {
                lines.push(format!("{}\t{}", row.part_number, row.model_number));
            }
        }
        CadFormat::Bd => {
            lines.push("# Kyoden BOM Tool CAD Export (BD)".to_string());
            lines.push("PART_NUMBER,MODEL_NUMBER".to_string());
            for row in &bom.rows {
                lines.push(format!("{},{}", row.part_number, row.model_number));
            }
        }
        CadFormat::Pws => {
            lines.push("# Kyoden BOM Tool CAD Export (PWS)".to_string());
            lines.push("[Component List]".to_string());
            for row in &bom.rows {
                lines.push(format!("{}={}", row.part_number, row.model_number));
            }
        }
    }
    if !bom.headers.is_empty() {
        lines.push("".to_string());
        lines.push("# Attributes".to_string());
        for row in &bom.rows {
            if row.attributes.is_empty() {
                continue;
            }
            let attrs = row
                .attributes
                .iter()
                .map(|(k, v)| format!("{k}={v}"))
                .collect::<Vec<_>>()
                .join(", ");
            lines.push(format!("{} => {}", row.part_number, attrs));
        }
    }
    lines
}

fn determine_cad_output_path(
    format: &CadFormat,
    provided: Option<String>,
) -> Result<PathBuf, String> {
    if let Some(path) = provided {
        let path = PathBuf::from(path);
        if path.extension().is_none() {
            let mut with_ext = path.clone();
            with_ext.set_extension(format.default_extension());
            return Ok(with_ext);
        }
        return Ok(path);
    }

    let base_dir = Path::new("../sessions/cad");
    let timestamp = chrono::Local::now().format("%Y%m%d_%H%M%S");
    let file_name = format!(
        "cad_{}_{}.{}",
        format.display_name().to_lowercase(),
        timestamp,
        format.default_extension()
    );
    Ok(base_dir.join(file_name))
}

#[derive(Debug, Clone, Deserialize)]
struct DialogFilter {
    name: String,
    extensions: Vec<String>,
}

#[tauri::command]
async fn save_file_dialog(
    app: tauri::AppHandle,
    default_path: Option<String>,
    filters: Option<Vec<DialogFilter>>,
) -> Result<Option<String>, String> {
    let (tx, rx) = oneshot::channel();
    let mut builder = app.dialog().file();

    if let Some(path) = default_path.as_ref() {
        let pb = PathBuf::from(path);
        if let Some(parent) = pb.parent() {
            builder = builder.set_directory(parent.to_path_buf());
        }
        if let Some(file_name) = pb.file_name() {
            builder = builder.set_file_name(file_name.to_string_lossy().to_string());
        }
    }

    if let Some(filters) = filters {
        for filter in filters {
            let extension_refs: Vec<&str> = filter.extensions.iter().map(|s| s.as_str()).collect();
            builder = builder.add_filter(filter.name, &extension_refs);
        }
    }

    builder.save_file(move |file| {
        let path = file
            .and_then(|fp| fp.into_path().ok())
            .map(|p| p.to_string_lossy().into_owned());
        let _ = tx.send(path);
    });

    match rx.await {
        Ok(result) => Ok(result),
        Err(_) => Ok(None),
    }
}
