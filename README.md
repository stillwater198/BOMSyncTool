# BOMSyncTool - 部品表比較・変換・合成ツール

## 概要

BOMSyncToolは、部品表の比較・変換・合成を行うデスクトップアプリケーションです。RustとTauriを使用して開発され、Windows、macOS、Linuxで動作します。

## 主な機能

### 1. ファイル読み込み
- **サポートフォーマット**: Excel (.xlsx/.xls), CSV (UTF-8/Shift-JIS)
- **入力方法**: ファイル選択ダイアログまたはドラッグ&ドロップ
- **列指定**: 部品番号、型番、数量の列をGUIで指定可能

### 2. データ標準化
- 改行削除
- 全角→半角変換
- 空白削除
- 空欄セル補完
- 大文字・小文字区別なし比較

### 3. 比較機能
- 部品表AとBの照合
- 共通部品、Aのみ部品、Bのみ部品の分類
- 結果をテーブル形式で表示

### 4. 合成機能
- 代替合成部品表の作成
- 数量の合算または別表示オプション
- 欠品フラグ付与

### 5. 出力機能
- CSV/TXT形式での保存
- partECO、partCCF、partMSF形式対応
- PartsSaver用出力

### 6. 補助機能
- シートクリア
- 登録名リスト（JSON形式）
- 統計情報表示

## システム要件

### 必要環境
- **OS**: Windows 10/11, macOS 10.15+, Linux (Ubuntu 18.04+)
- **メモリ**: 4GB以上推奨
- **ストレージ**: 100MB以上の空き容量
- **ネットワーク**: 不要（完全オフライン動作）

### 依存関係
- 追加ランタイム不要
- インターネット接続不要
- VBA/ActiveX不要

## インストール

### 1. リリース版のダウンロード
[Releases](https://github.com/your-repo/kyoden-bom-tool/releases)から各OS向けの実行ファイルをダウンロードしてください。

### 2. ビルド方法（開発者向け）

#### 前提条件
- Rust 1.70+
- Node.js 16+ (フロントエンド開発用)

#### ビルド手順
```bash
# リポジトリをクローン
git clone https://github.com/your-repo/kyoden-bom-tool.git
cd kyoden-bom-tool

# Rustの依存関係をインストール
cd src-tauri
cargo build --release

# アプリケーションをビルド
cargo tauri build
```

## 使い方

### 1. アプリケーションの起動
- Windows: `kyoden-bom-tool.exe`を実行
- macOS: `Kyoden BOM Tool.app`を起動
- Linux: `kyoden-bom-tool`を実行

### 2. ファイル読み込み
1. 「ファイルを選択」ボタンをクリック
2. 部品表ファイル（Excel/CSV）を選択
3. 列マッピングで部品番号、型番、数量の列を指定
4. 「部品表Aを読み込み」または「部品表Bを読み込み」をクリック

### 3. 比較実行
1. 部品表AとBの両方を読み込み
2. 「比較実行」ボタンをクリック
3. 結果を「比較結果」タブで確認

### 4. 合成実行
1. 部品表AとBの両方を読み込み
2. 数量の処理方法を選択（合算/別表示）
3. 「合成実行」ボタンをクリック
4. 結果を「合成結果」タブで確認

### 5. 結果保存
1. 保存したい結果のタブを選択
2. 保存形式（CSV/TXT）を選択
3. 「保存」ボタンをクリック
4. 保存先を指定

## ファイル形式

### 入力ファイル
- **Excel**: .xlsx, .xls
- **CSV**: UTF-8, Shift-JISエンコーディング対応

### 出力ファイル
- **CSV**: UTF-8エンコーディング
- **TXT**: UTF-8またはShift-JISエンコーディング
- **JSON**: 登録名リスト用

## サンプルデータ

`sample_data/`フォルダにテスト用のサンプルファイルが含まれています：

- `bom_a.csv`: 部品表Aのサンプル
- `bom_b.csv`: 部品表Bのサンプル
- `bom_large.csv`: 大量データテスト用

## パフォーマンス

- **処理速度**: 10万行の部品表を1秒以内で比較・合成
- **メモリ使用量**: 最小限に最適化
- **並列処理**: rayonクレートによる高速化

## トラブルシューティング

### よくある問題

#### 1. ファイル読み込みエラー
- **原因**: ファイル形式が不正、列指定が不完全
- **解決策**: ファイル形式を確認し、必要な列を正しく指定

#### 2. 列指定エラー
- **原因**: 部品番号、型番、数量の列が未指定
- **解決策**: 各列をコンボボックスで選択

#### 3. 保存エラー
- **原因**: 権限不足、パスが無効
- **解決策**: 保存先の権限を確認し、有効なパスを指定

#### 4. メモリ不足
- **原因**: 非常に大きなファイルの処理
- **解決策**: ファイルを分割して処理

### ログ確認
エラーが発生した場合は、アプリケーションのログを確認してください：
- Windows: `%APPDATA%/kyoden-bom-tool/logs/`
- macOS: `~/Library/Logs/kyoden-bom-tool/`
- Linux: `~/.local/share/kyoden-bom-tool/logs/`

## 技術仕様

### アーキテクチャ
- **バックエンド**: Rust
- **フロントエンド**: HTML/CSS/JavaScript
- **フレームワーク**: Tauri 2.0
- **UI**: DataTables.js

### 主要クレート
- `calamine`: Excelファイル読み込み
- `csv`: CSVファイル処理
- `rayon`: 並列処理
- `encoding_rs`: 文字エンコーディング変換
- `serde_json`: JSON処理

### データ構造
```rust
struct BomRow {
    part_number: String,
    model_number: String,
    attributes: HashMap<String, String>,
}
```

## ライセンス

MIT License

## 貢献

プルリクエストやイシューの報告を歓迎します。

## 更新履歴

### v0.1.0 (2025-10-12予定)
- 初回リリース
- 基本的な比較・合成機能
- Excel/CSV読み込み対応
- 多形式出力対応

## サポート

- **Issues**: [GitHub Issues](https://github.com/your-repo/kyoden-bom-tool/issues)
- **Documentation**: [Wiki](https://github.com/your-repo/kyoden-bom-tool/wiki)
- **Email**: support@kyoden-tool.com

## 開発者情報

- **開発**: Sota Hemmi
- **言語**: Rust, JavaScript, HTML, CSS
- **フレームワーク**: Tauri
- **ライセンス**: MIT
