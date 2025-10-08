use chrono::{DateTime, Utc};
use rand::{distributions::Alphanumeric, Rng};
use serde::{Deserialize, Serialize};
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use crate::{
    BomData, ColumnMapping, ComparisonResult, OverrideList, RegisteredNameList, SynthesisResult,
};

const AUTO_DIR: &str = "../sessions/auto";
const MANUAL_DIR: &str = "../sessions/manual";
const AUTO_LIMIT: usize = 10;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionSnapshot {
    pub id: String,
    pub label: Option<String>,
    pub created_at: DateTime<Utc>,
    pub file_a_path: Option<String>,
    pub file_b_path: Option<String>,
    pub column_mapping_a: Option<ColumnMapping>,
    pub column_mapping_b: Option<ColumnMapping>,
    pub bom_a: Option<BomData>,
    pub bom_b: Option<BomData>,
    pub comparison_result: Option<ComparisonResult>,
    pub synthesis_result: Option<SynthesisResult>,
    pub registered_name_list: Option<RegisteredNameList>,
    pub override_list: Option<OverrideList>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionSummary {
    pub id: String,
    pub label: Option<String>,
    pub created_at: DateTime<Utc>,
    pub file_a_name: Option<String>,
    pub file_b_name: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SessionKind {
    Auto,
    Manual,
}

impl SessionKind {
    fn directory(&self) -> &'static str {
        match self {
            SessionKind::Auto => AUTO_DIR,
            SessionKind::Manual => MANUAL_DIR,
        }
    }
}

fn ensure_directory(path: &Path) -> Result<(), String> {
    fs::create_dir_all(path).map_err(|e| format!("ディレクトリ作成に失敗しました: {e}"))
}

fn session_dir(kind: SessionKind) -> Result<PathBuf, String> {
    let dir = PathBuf::from(kind.directory());
    ensure_directory(&dir)?;
    Ok(dir)
}

fn generate_id() -> String {
    let rand_str: String = rand::thread_rng()
        .sample_iter(&Alphanumeric)
        .take(8)
        .map(char::from)
        .collect();
    format!("{}-{}", Utc::now().timestamp(), rand_str)
}

fn snapshot_to_summary(snapshot: &SessionSnapshot) -> SessionSummary {
    SessionSummary {
        id: snapshot.id.clone(),
        label: snapshot.label.clone(),
        created_at: snapshot.created_at,
        file_a_name: snapshot
            .file_a_path
            .as_ref()
            .map(|p| {
                Path::new(p)
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
            })
            .flatten(),
        file_b_name: snapshot
            .file_b_path
            .as_ref()
            .map(|p| {
                Path::new(p)
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
            })
            .flatten(),
    }
}

pub fn save_snapshot(
    mut snapshot: SessionSnapshot,
    kind: SessionKind,
) -> Result<SessionSummary, String> {
    if snapshot.id.is_empty() {
        snapshot.id = generate_id();
    }
    let dir = session_dir(kind)?;
    let path = dir.join(format!("{}.json", snapshot.id));
    let mut file =
        File::create(&path).map_err(|e| format!("セッション保存ファイルを作成できません: {e}"))?;
    let json = serde_json::to_string_pretty(&snapshot)
        .map_err(|e| format!("セッションのシリアライズに失敗しました: {e}"))?;
    file.write_all(json.as_bytes())
        .map_err(|e| format!("セッション保存に失敗しました: {e}"))?;

    if kind == SessionKind::Auto {
        prune_auto_sessions()?;
    }

    Ok(snapshot_to_summary(&snapshot))
}

fn prune_auto_sessions() -> Result<(), String> {
    let dir = session_dir(SessionKind::Auto)?;
    let mut snapshots = collect_snapshots(SessionKind::Auto)?;
    snapshots.sort_by(|a, b| b.created_at.cmp(&a.created_at));
    if snapshots.len() <= AUTO_LIMIT {
        return Ok(());
    }
    for summary in snapshots.into_iter().skip(AUTO_LIMIT) {
        let path = dir.join(format!("{}.json", summary.id));
        let _ = fs::remove_file(path);
    }
    Ok(())
}

fn read_snapshot(path: &Path) -> Result<SessionSnapshot, String> {
    let mut file = File::open(path).map_err(|e| format!("セッションを開けません: {e}"))?;
    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| format!("セッションの読み込みに失敗しました: {e}"))?;
    serde_json::from_str(&content).map_err(|e| format!("セッションの解析に失敗しました: {e}"))
}

pub fn collect_snapshots(kind: SessionKind) -> Result<Vec<SessionSummary>, String> {
    let dir = session_dir(kind)?;
    let mut summaries = Vec::new();
    for entry in fs::read_dir(&dir)
        .map_err(|e| format!("セッションディレクトリの読み込みに失敗しました: {e}"))?
    {
        let entry =
            entry.map_err(|e| format!("ディレクトリエントリの読み込みに失敗しました: {e}"))?;
        let path = entry.path();
        if path.extension().and_then(|ext| ext.to_str()) != Some("json") {
            continue;
        }
        if let Ok(snapshot) = read_snapshot(&path) {
            summaries.push(snapshot_to_summary(&snapshot));
        }
    }
    summaries.sort_by(|a, b| b.created_at.cmp(&a.created_at));
    Ok(summaries)
}

pub fn load_snapshot(kind: SessionKind, id: &str) -> Result<SessionSnapshot, String> {
    let dir = session_dir(kind)?;
    let path = dir.join(format!("{}.json", id));
    read_snapshot(&path)
}

pub fn delete_snapshot(kind: SessionKind, id: &str) -> Result<(), String> {
    let dir = session_dir(kind)?;
    let path = dir.join(format!("{}.json", id));
    fs::remove_file(&path).map_err(|e| format!("セッションの削除に失敗しました: {e}"))
}
