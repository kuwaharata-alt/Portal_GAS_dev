# プロジェクト構成（詳細）

## 概要
このドキュメントは `clone_kibanPortal` プロジェクトの構成と主要ファイル、開発／デプロイ手順をまとめたものです。

## 主要ファイルと役割
- `0000_Code.js` – エントリ／共通初期化
- `0001_gs_initial.js` ～ `00XX_gs_*.js` – サーバー（GAS）側スクリプト
- `10XX_css_*.html` – スタイル断片（`include()` で読み込む）
- `20XX_html_*.html` – HTML テンプレート断片
- `30XX_js_*.html` – クライアント側 JS 断片（HTML 内に埋め込む）
- `index.html` – 単一エントリ（`include()` で断片を結合）
- `appsscript.json` – GAS マニフェスト（WebApp: `access: DOMAIN`）
- `.clasp.json` – clasp 設定
- `9999_index生成.js` – インデックス生成ユーティリティ（任意）

## シート（スプレッドシート）と用途
- `Display` – 表示用設定（タブ／アクセス権など）
- `Master` / `DB` – マスタデータ、メンバー情報、祝日など
- 各シート名はソース内（例: `INPUT_CONFIG.SHEET_MEMBER`）を参照してください。

## 開発ワークフロー（推奨）
1. 変更をローカルで編集
2. `clasp push` で GAS に反映
3. スクリプトエディタまたは WebApp で動作確認
4. 必要に応じてデプロイ

## 命名規則（推奨）
- プレフィックスで機能別管理（`0XXX`: GAS、`10XX`: CSS、`20XX`: HTML、`30XX`: client JS、`99XX`: ユーティリティ）
- ファイル名は `prefix_description.ext` 形式
- 一時ファイルは `archive/` へ移動または削除

## 整理・改善提案（次のステップ）
- `README.md` を充実化（完了済）
- `archive/` フォルダを作成して `package.json` 等を格納
- ESLint/Prettier のルールを決めてコード整形を実行
- 自動テストや CI（必要なら）を検討

## 備考
- `無題.js` と `編集用.js` は一時ファイルとして扱う（`README.md` に記載済）。

