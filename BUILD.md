# ビルド手順

## 前提条件

- Rust 1.70以上
- Node.js 16以上（フロントエンド開発用）
- 各OSの開発ツールチェーン

## ビルド手順

### 1. 依存関係のインストール

```bash
# Rustのインストール（未インストールの場合）
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
source $HOME/.cargo/env

# Tauri CLIのインストール
cargo install tauri-cli --locked
```

### 2. プロジェクトのビルド

```bash
# プロジェクトディレクトリに移動
cd /path/to/KyodenBomTool/src-tauri

# 依存関係のチェック
cargo check

# 開発ビルド
cargo tauri dev

# 本番ビルド
cargo tauri build
```

### 3. 出力ファイル

本番ビルド後、以下の場所に実行ファイルが生成されます：

- **Windows**: `src-tauri/target/release/bundle/msi/Kyoden BOM Tool_0.1.0_x64_en-US.msi`
- **macOS**: `src-tauri/target/release/bundle/dmg/Kyoden BOM Tool_0.1.0_x64.dmg`
- **Linux**: `src-tauri/target/release/bundle/deb/kyoden-bom-tool_0.1.0_amd64.deb`

## トラブルシューティング

### よくある問題

1. **コンパイルエラー**
   ```bash
   # 依存関係を更新
   cargo update
   
   # クリーンビルド
   cargo clean
   cargo build
   ```

2. **アイコンファイルエラー**
   - `src-tauri/icons/`ディレクトリに適切なPNGファイルが配置されていることを確認

3. **フロントエンドエラー**
   - `dist/`ディレクトリにHTML/CSS/JSファイルが存在することを確認

## 開発環境

### 開発モードでの実行

```bash
# 開発サーバーを起動
cargo tauri dev
```

### デバッグ

```bash
# デバッグビルド
cargo build

# リリースビルド
cargo build --release
```

## 配布用パッケージ作成

### 全プラットフォーム向けビルド

```bash
# 全プラットフォーム向けビルド
cargo tauri build --target x86_64-pc-windows-msvc
cargo tauri build --target x86_64-apple-darwin
cargo tauri build --target x86_64-unknown-linux-gnu
```

### サイズ最適化

```bash
# バイナリサイズの最適化
cargo tauri build --release
strip target/release/kyoden-bom-tool
```

## テスト

### 単体テスト

```bash
# 全テストを実行
cargo test

# 特定のテストを実行
cargo test test_name
```

### 統合テスト

```bash
# サンプルデータでのテスト
cargo run -- --test-mode
```

## パフォーマンス最適化

### リリースビルドの最適化

`Cargo.toml`に以下を追加：

```toml
[profile.release]
lto = true
codegen-units = 1
panic = "abort"
```

### メモリ使用量の最適化

```bash
# メモリプロファイリング
cargo install cargo-profdata
cargo profdata --bin kyoden-bom-tool
```
