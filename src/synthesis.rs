use crate::{BomData, SynthesisResult, SynthesisRow};
use rayon::prelude::*;
use std::collections::{HashMap, HashSet};

/// 部品表AとBを合成して代替合成部品表を作成する
pub fn perform_synthesis(bom_a: &BomData, bom_b: &BomData) -> SynthesisResult {
    let map_a: HashMap<String, &crate::BomRow> = bom_a
        .rows
        .iter()
        .map(|row| (row.part_number.clone(), row))
        .collect();
    let map_b: HashMap<String, &crate::BomRow> = bom_b
        .rows
        .iter()
        .map(|row| (row.part_number.clone(), row))
        .collect();

    let mut all_part_numbers: HashSet<String> = HashSet::new();
    all_part_numbers.extend(map_a.keys().cloned());
    all_part_numbers.extend(map_b.keys().cloned());

    let mut rows: Vec<SynthesisRow> = all_part_numbers
        .par_iter()
        .map(|part_number| {
            let row_a = map_a.get(part_number);
            let row_b = map_b.get(part_number);

            match (row_a, row_b) {
                (Some(a), Some(b)) => SynthesisRow {
                    part_number: part_number.clone(),
                    model_a: a.model_number.clone(),
                    model_b: b.model_number.clone(),
                    status: "common".to_string(),
                },
                (Some(a), None) => SynthesisRow {
                    part_number: part_number.clone(),
                    model_a: a.model_number.clone(),
                    model_b: String::new(),
                    status: "missing_b".to_string(),
                },
                (None, Some(b)) => SynthesisRow {
                    part_number: part_number.clone(),
                    model_a: String::new(),
                    model_b: b.model_number.clone(),
                    status: "missing_a".to_string(),
                },
                (None, None) => SynthesisRow {
                    part_number: part_number.clone(),
                    model_a: String::new(),
                    model_b: String::new(),
                    status: "unknown".to_string(),
                },
            }
        })
        .collect();

    rows.par_sort_by(|a, b| a.part_number.cmp(&b.part_number));

    SynthesisResult { rows }
}

/// 合成結果をCSV形式で保存
pub async fn save_synthesis_result(
    result: &SynthesisResult,
    file_path: &str,
    format: &str,
) -> Result<String, String> {
    match format {
        "csv" => {
            let mut csv_data = Vec::new();
            csv_data.push(vec![
                "部品番号".to_string(),
                "型番A".to_string(),
                "型番B".to_string(),
                "ステータス".to_string(),
            ]);

            for row in &result.rows {
                csv_data.push(vec![
                    row.part_number.clone(),
                    row.model_a.clone(),
                    row.model_b.clone(),
                    get_status_text(&row.status),
                ]);
            }

            crate::file_handler::save_csv_file(&csv_data, file_path, "utf-8")
                .await
                .map_err(|e| format!("CSV保存エラー: {e}"))?;
        }
        "txt" => {
            let mut content = String::new();
            content.push_str("=== 代替合成部品表 ===\n\n");

            let stats = get_synthesis_stats(result);
            content.push_str(&format!(
                "総部品数: {}件\n",
                stats.get("total").copied().unwrap_or(0)
            ));
            content.push_str(&format!(
                "共通部品: {}件\n",
                stats.get("common").copied().unwrap_or(0)
            ));
            content.push_str(&format!(
                "A欠品: {}件\n",
                stats.get("missing_a").copied().unwrap_or(0)
            ));
            content.push_str(&format!(
                "B欠品: {}件\n\n",
                stats.get("missing_b").copied().unwrap_or(0)
            ));

            content.push_str("=== 部品一覧 ===\n");
            for row in &result.rows {
                content.push_str(&format!(
                    "{} | {} | {} | {}\n",
                    row.part_number,
                    row.model_a,
                    row.model_b,
                    get_status_text(&row.status)
                ));
            }

            crate::file_handler::save_txt_file(&content, file_path, "utf-8")
                .await
                .map_err(|e| format!("TXT保存エラー: {e}"))?;
        }
        _ => return Err("サポートされていないフォーマットです".to_string()),
    }

    Ok("合成結果を保存しました".to_string())
}

/// 合成結果の統計情報を取得
pub fn get_synthesis_stats(result: &SynthesisResult) -> HashMap<String, usize> {
    let mut stats = HashMap::new();
    stats.insert("total".to_string(), result.rows.len());
    stats.insert(
        "common".to_string(),
        result
            .rows
            .iter()
            .filter(|row| row.status == "common")
            .count(),
    );
    stats.insert(
        "missing_a".to_string(),
        result
            .rows
            .iter()
            .filter(|row| row.status == "missing_a")
            .count(),
    );
    stats.insert(
        "missing_b".to_string(),
        result
            .rows
            .iter()
            .filter(|row| row.status == "missing_b")
            .count(),
    );
    stats
}

/// ステータスで合成結果をフィルタリングする
pub fn filter_synthesis_result(result: &SynthesisResult, status: Option<&str>) -> SynthesisResult {
    let filtered_rows = match status {
        Some(status_filter) if !status_filter.trim().is_empty() => result
            .rows
            .iter()
            .filter(|row| row.status.eq_ignore_ascii_case(status_filter))
            .cloned()
            .collect(),
        _ => result.rows.clone(),
    };

    SynthesisResult {
        rows: filtered_rows,
    }
}

/// 指定したステータスの部品を抽出する
pub fn collect_missing_parts(result: &SynthesisResult) -> (Vec<SynthesisRow>, Vec<SynthesisRow>) {
    let missing_a = result
        .rows
        .iter()
        .filter(|row| row.status == "missing_a")
        .cloned()
        .collect();
    let missing_b = result
        .rows
        .iter()
        .filter(|row| row.status == "missing_b")
        .cloned()
        .collect();

    (missing_a, missing_b)
}

fn get_status_text(status: &str) -> String {
    match status {
        "common" => "共通".to_string(),
        "missing_a" => "A欠品".to_string(),
        "missing_b" => "B欠品".to_string(),
        _ => "不明".to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{BomData, BomRow};
    use std::collections::HashMap;

    fn create_test_bom_a() -> BomData {
        BomData {
            headers: vec!["部品番号".to_string(), "型番".to_string()],
            rows: vec![
                BomRow {
                    part_number: "PART001".to_string(),
                    model_number: "MODEL001".to_string(),
                    attributes: HashMap::new(),
                },
                BomRow {
                    part_number: "PART002".to_string(),
                    model_number: "MODEL002".to_string(),
                    attributes: HashMap::new(),
                },
            ],
        }
    }

    fn create_test_bom_b() -> BomData {
        BomData {
            headers: vec!["部品番号".to_string(), "型番".to_string()],
            rows: vec![
                BomRow {
                    part_number: "PART001".to_string(),
                    model_number: "MODEL001".to_string(),
                    attributes: HashMap::new(),
                },
                BomRow {
                    part_number: "PART003".to_string(),
                    model_number: "MODEL003".to_string(),
                    attributes: HashMap::new(),
                },
            ],
        }
    }

    #[test]
    fn test_perform_synthesis() {
        let bom_a = create_test_bom_a();
        let bom_b = create_test_bom_b();

        let result = perform_synthesis(&bom_a, &bom_b);

        assert_eq!(result.rows.len(), 3);

        let part001 = result
            .rows
            .iter()
            .find(|r| r.part_number == "PART001")
            .unwrap();
        assert_eq!(part001.status, "common");

        let part002 = result
            .rows
            .iter()
            .find(|r| r.part_number == "PART002")
            .unwrap();
        assert_eq!(part002.status, "missing_b");

        let part003 = result
            .rows
            .iter()
            .find(|r| r.part_number == "PART003")
            .unwrap();
        assert_eq!(part003.status, "missing_a");
    }

    #[test]
    fn test_get_synthesis_stats() {
        let result = SynthesisResult { rows: vec![] };
        let stats = get_synthesis_stats(&result);

        assert_eq!(stats.get("total").unwrap(), &0);
        assert_eq!(stats.get("common").unwrap(), &0);
    }

    #[test]
    fn test_filter_synthesis_result() {
        let row1 = SynthesisRow {
            part_number: "PART001".to_string(),
            model_a: "MODEL001".to_string(),
            model_b: "MODEL001".to_string(),
            status: "common".to_string(),
        };

        let row2 = SynthesisRow {
            part_number: "PART002".to_string(),
            model_a: "MODEL002".to_string(),
            model_b: String::new(),
            status: "missing_b".to_string(),
        };

        let result = SynthesisResult {
            rows: vec![row1, row2],
        };

        let filtered = filter_synthesis_result(&result, Some("common"));
        assert_eq!(filtered.rows.len(), 1);
        assert_eq!(filtered.rows[0].part_number, "PART001");
    }
}
