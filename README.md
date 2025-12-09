# PCA2ICS

PCA公益法人会計V.12 → ICS財務処理db (db仕訳形式) 変換スクリプト

## 機能

- 202509.csv形式のデータをICS財務処理db形式に変換
- 科目コード変換 (PCAコード → ICSコード)
- 税区分変換 (PCA公益 → ICS db形式)
- 複合仕訳の単純仕訳への分解
- 日付フォーマット変換

## セットアップ

### 1. 依存パッケージのインストール

```bash
npm install
```

### 2. claspへのログイン

```bash
npm run login
```

### 3. スクリプトIDの設定

`.clasp.json.example`をコピーして`.clasp.json`を作成し、`scriptId`を設定してください。

```bash
cp .clasp.json.example .clasp.json
```

`.clasp.json`を編集して実際のスクリプトIDを設定：

```json
{
  "scriptId": "YOUR_ACTUAL_SCRIPT_ID",
  "rootDir": "./dist"
}
```

スクリプトIDは、Google Apps Scriptエディタの「プロジェクトの設定」から確認できます。

> **Note**: `.clasp.json`はgit管理外です（セキュリティのため）。

## 開発

### ビルド

TypeScriptをコンパイルして`dist/`フォルダに出力します。

```bash
npm run build
```

### 監視モード

ファイル変更時に自動でコンパイルします。

```bash
npm run watch
```

### デプロイ

ビルドしてGoogle Apps Scriptにプッシュします。

```bash
npm run deploy
```

## プロジェクト構成

```
pca2ics/
├── src/                    # TypeScriptソースファイル
│   ├── types.ts           # 型定義
│   ├── config.ts          # 設定
│   ├── utils.ts           # ユーティリティ関数
│   ├── dataLoader.ts      # データ読み込み関数
│   ├── converter.ts       # 変換ロジック
│   ├── output.ts          # 出力関数
│   └── main.ts            # メインエントリポイント
├── dist/                   # コンパイル出力（git管理外）
├── .clasp.json            # clasp設定（git管理外、要作成）
├── .clasp.json.example    # clasp設定のサンプル
├── tsconfig.json          # TypeScript設定
├── package.json           # npm設定
└── README.md
```

## Google Spreadsheet側の準備

以下のシートを用意してください：

1. **元データシート** (例: `202509`) - PCA公益法人会計からエクスポートしたCSVデータ
2. **科目対応表** - 勘定科目名、ICSコード、PCAコードの対応表
3. **税区分マッピング** (自動作成可) - PCA税区分コードとICS税区分コードの対応表

## 使用方法

1. スプレッドシートを開く
2. メニューから「PCA→ICS変換」→「変換実行」を選択
3. 元データシートを選択
4. 変換結果は「ICS変換結果」シートに出力されます

## GitHub Actions

mainブランチへのプッシュ時に自動でGoogle Apps Scriptにデプロイされます。

### シークレットの設定

GitHub Actionsでデプロイするには、以下のシークレットを設定してください：

- `CLASP_TOKEN`: claspの認証トークン（`~/.clasprc.json`の内容）
- `SCRIPT_ID`: Google Apps ScriptのスクリプトID
