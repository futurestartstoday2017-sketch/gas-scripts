# CLAUDE.md

このファイルは、このリポジトリで作業する Claude Code 向けのガイドです。

## プロジェクト概要

Google Apps Script (GAS) のスクリプト群を管理するリポジトリ。

- 言語: JavaScript（GAS ランタイム V8）
- ローカル開発・デプロイ: [clasp](https://github.com/google/clasp) を使用
- 応答・コメント・コミットメッセージは日本語で書く

## ディレクトリ構成（想定）

```
gas-scripts/
├── CLAUDE.md
├── README.md
├── .gitignore
└── <プロジェクト名>/        # GAS プロジェクトごとにフォルダを分ける
    ├── .clasp.json          # scriptId を含む（Git 管理対象外）
    ├── appsscript.json      # マニフェスト
    └── *.js / *.gs / *.html
```

## よく使うコマンド

```bash
npm install -g @google/clasp   # clasp のインストール
clasp login                    # Google アカウントでログイン
clasp clone <scriptId>         # 既存プロジェクトの取得
clasp pull                     # GAS エディタ側の変更をローカルへ取得
clasp push                     # ローカルの変更を GAS へ反映
clasp open                     # ブラウザで GAS エディタを開く
clasp logs                     # 実行ログの確認
```

## コーディング規約

- `const` / `let` を使い、`var` は使わない
- 関数名・変数名は camelCase、定数は UPPER_SNAKE_CASE
- トリガーから呼ばれる関数や公開関数には JSDoc で用途を書く
- GAS の実行時間上限（6 分）を意識し、大量データはバッチ処理する
- `SpreadsheetApp` 等への呼び出しはループ内で繰り返さず、`getValues()` / `setValues()` でまとめて読み書きする
- 外部 API 呼び出しは `UrlFetchApp` を使い、`muteHttpExceptions` とエラーハンドリングを付ける

## 秘密情報の扱い

- API キー・トークン・パスワードはコードに直接書かず、`PropertiesService` のスクリプトプロパティに保存する
- `.clasp.json`、`.clasprc.json`、認証情報ファイルは Git にコミットしない（`.gitignore` に追加する）
- コミット前に差分に秘密情報が含まれていないか確認する

## Git 運用ルール

**コードを変更するたびに、コミットして GitHub にプッシュすること。**

1. 変更が一区切りついたら（1 つの修正・機能追加ごとに）コミットする
2. コミット後は必ず `git push` で GitHub のリモートリポジトリに反映する
3. 変更をローカルに溜め込まない。作業の終わりにはプッシュ済みの状態にする

### コミットメッセージ

- 日本語で、何をなぜ変更したかを簡潔に書く
- 1 行目は要約（50 文字程度まで）、必要なら空行を挟んで詳細を書く
- 例: `売上集計スクリプトに月次集計機能を追加`

### 手順

```bash
git add <変更したファイル>
git commit -m "変更内容の要約"
git push
```

### 注意事項

- プッシュ前に秘密情報（API キー、scriptId を含む `.clasp.json` 等）が含まれていないか確認する
- プッシュが失敗した場合（リモートの方が新しい等）は、`git pull --rebase` で取り込んでから再度プッシュする。`--force` での強制プッシュは行わない
- `git push` と `clasp push` は別物。GAS 側への反映が必要な場合は `clasp push` も実行する
