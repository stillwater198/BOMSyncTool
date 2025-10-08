use crate::{BomData, ComparisonResult, ComparisonRow};
use rayon::prelude::*;
use std::collections::HashMap;

/// 部品表AとBを比較する
pub async fn perform_comparison(bom_a: &BomData, bom_b: &BomData) -> ComparisonResult {
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

    let (common_parts, a_only_parts) = rayon::join(
        || find_common_parts(&map_a, &map_b),
        || find_a_only_parts(&map_a, &map_b),
    );
    let (b_only_parts, modified_parts) = rayon::join(
        || find_b_only_parts(&map_a, &map_b),
        || find_modified_parts(&map_a, &map_b),
    );

    ComparisonResult {
        common_parts,
        a_only_parts,
        b_only_parts,
        modified_parts,
    }
}

fn find_common_parts(
    map_a: &HashMap<String, &crate::BomRow>,
    map_b: &HashMap<String, &crate::BomRow>,
) -> Vec<ComparisonRow> {
    map_a
        .par_iter()
        .filter(|(part_number, _)| map_b.contains_key(*part_number))
        .map(|(part_number, row_a)| {
            let row_b = map_b.get(part_number).unwrap();
            let is_modified = row_a.model_number != row_b.model_number;
            ComparisonRow {
                part_number: part_number.clone(),
                model_a: row_a.model_number.clone(),
                model_b: row_b.model_number.clone(),
                status: if is_modified {
                    "modified".to_string()
                } else {
                    "common".to_string()
                },
                change_type: if is_modified {
                    "MODIFIED".to_string()
                } else {
                    "UNCHANGED".to_string()
                },
            }
        })
        .collect()
}

fn find_a_only_parts(
    map_a: &HashMap<String, &crate::BomRow>,
    map_b: &HashMap<String, &crate::BomRow>,
) -> Vec<ComparisonRow> {
    map_a
        .par_iter()
        .filter(|(part_number, _)| !map_b.contains_key(*part_number))
        .map(|(part_number, row_a)| ComparisonRow {
            part_number: part_number.clone(),
            model_a: row_a.model_number.clone(),
            model_b: String::new(),
            status: "a_only".to_string(),
            change_type: "REMOVED".to_string(),
        })
        .collect()
}

fn find_b_only_parts(
    map_a: &HashMap<String, &crate::BomRow>,
    map_b: &HashMap<String, &crate::BomRow>,
) -> Vec<ComparisonRow> {
    map_b
        .par_iter()
        .filter(|(part_number, _)| !map_a.contains_key(*part_number))
        .map(|(part_number, row_b)| ComparisonRow {
            part_number: part_number.clone(),
            model_a: String::new(),
            model_b: row_b.model_number.clone(),
            status: "b_only".to_string(),
            change_type: "ADDED".to_string(),
        })
        .collect()
}

fn find_modified_parts(
    map_a: &HashMap<String, &crate::BomRow>,
    map_b: &HashMap<String, &crate::BomRow>,
) -> Vec<ComparisonRow> {
    map_a
        .par_iter()
        .filter(|(part_number, _)| map_b.contains_key(*part_number))
        .filter(|(_, row_a)| {
            let row_b = map_b.get((*part_number).as_str()).unwrap();
            row_a.model_number != row_b.model_number
        })
        .map(|(part_number, row_a)| {
            let row_b = map_b.get(part_number).unwrap();
            ComparisonRow {
                part_number: part_number.clone(),
                model_a: row_a.model_number.clone(),
                model_b: row_b.model_number.clone(),
                status: "modified".to_string(),
                change_type: "MODIFIED".to_string(),
            }
        })
        .collect()
}

pub async fn save_comparison_result(
    result: &ComparisonResult,
    file_path: &str,
    format: &str,
) -> Result<String, String> {
    let mut csv_data = Vec::new();

    csv_data.push(vec![
        "部品番号".to_string(),
        "型番A".to_string(),
        "型番B".to_string(),
        "ステータス".to_string(),
        "差分種別".to_string(),
    ]);

    for row in result
        .common_parts
        .iter()
        .chain(result.a_only_parts.iter())
        .chain(result.b_only_parts.iter())
        .chain(result.modified_parts.iter())
    {
        csv_data.push(vec![
            row.part_number.clone(),
            row.model_a.clone(),
            row.model_b.clone(),
            get_status_text(&row.status),
            get_change_type_text(&row.change_type),
        ]);
    }

    match format {
        "csv" => {
            crate::file_handler::save_csv_file(&csv_data, file_path, "utf-8")
                .await
                .map_err(|e| format!("CSV保存エラー: {}", e))?;
        }
        "txt" => {
            let mut content = String::new();
            content.push_str("=== 部品表比較結果 ===\n\n");
            content.push_str(&format!("共通部品: {}件\n", result.common_parts.len()));
            content.push_str(&format!("Aのみ部品: {}件\n", result.a_only_parts.len()));
            content.push_str(&format!("Bのみ部品: {}件\n", result.b_only_parts.len()));
            content.push_str(&format!("変更部品: {}件\n\n", result.modified_parts.len()));

            content.push_str("=== 共通/変更部品 ===\n");
            for row in &result.common_parts {
                content.push_str(&format!(
                    "{} | {} | {} | {}\n",
                    row.part_number,
                    row.model_a,
                    row.model_b,
                    get_change_type_text(&row.change_type)
                ));
            }
            for row in &result.modified_parts {
                content.push_str(&format!(
                    "{} | {} | {} | {}\n",
                    row.part_number,
                    row.model_a,
                    row.model_b,
                    get_change_type_text(&row.change_type)
                ));
            }

            content.push_str("\n=== Aのみ部品 ===\n");
            for row in &result.a_only_parts {
                content.push_str(&format!("{} | {}\n", row.part_number, row.model_a));
            }

            content.push_str("\n=== Bのみ部品 ===\n");
            for row in &result.b_only_parts {
                content.push_str(&format!("{} | {}\n", row.part_number, row.model_b));
            }

            crate::file_handler::save_txt_file(&content, file_path, "utf-8")
                .await
                .map_err(|e| format!("TXT保存エラー: {}", e))?;
        }
        _ => return Err("サポートされていないフォーマットです".to_string()),
    }

    Ok("比較結果を保存しました".to_string())
}

fn get_status_text(status: &str) -> String {
    match status {
        "common" => "共通部品".to_string(),
        "a_only" => "Aのみ".to_string(),
        "b_only" => "Bのみ".to_string(),
        "modified" => "変更".to_string(),
        _ => status.to_string(),
    }
}

fn get_change_type_text(change_type: &str) -> String {
    match change_type {
        "ADDED" => "追加".to_string(),
        "REMOVED" => "削除".to_string(),
        "MODIFIED" => "変更".to_string(),
        "UNCHANGED" => "変更なし".to_string(),
        _ => change_type.to_string(),
    }
}

pub fn get_comparison_stats(result: &ComparisonResult) -> HashMap<String, usize> {
    let mut stats = HashMap::new();
    stats.insert("common".to_string(), result.common_parts.len());
    stats.insert("a_only".to_string(), result.a_only_parts.len());
    stats.insert("b_only".to_string(), result.b_only_parts.len());
    stats.insert("modified".to_string(), result.modified_parts.len());
    stats.insert(
        "total_a".to_string(),
        result.common_parts.len() + result.a_only_parts.len(),
    );
    stats.insert(
        "total_b".to_string(),
        result.common_parts.len() + result.b_only_parts.len(),
    );
    stats
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

    #[tokio::test]
    async fn test_perform_comparison() {
        let bom_a = create_test_bom_a();
        let bom_b = create_test_bom_b();

        let result = perform_comparison(&bom_a, &bom_b).await;

        assert_eq!(result.common_parts.len(), 1);
        assert_eq!(result.a_only_parts.len(), 1);
        assert_eq!(result.b_only_parts.len(), 1);

        assert_eq!(result.common_parts[0].part_number, "PART001");
        assert_eq!(result.a_only_parts[0].part_number, "PART002");
        assert_eq!(result.b_only_parts[0].part_number, "PART003");
    }

    #[test]
    fn test_get_comparison_stats() {
        let result = ComparisonResult {
            common_parts: vec![],
            a_only_parts: vec![],
            b_only_parts: vec![],
            modified_parts: vec![],
        };

        let stats = get_comparison_stats(&result);
        assert_eq!(stats.get("common").unwrap(), &0);
        assert_eq!(stats.get("a_only").unwrap(), &0);
        assert_eq!(stats.get("b_only").unwrap(), &0);
        assert_eq!(stats.get("modified").unwrap(), &0);
    }
}
