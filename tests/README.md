# Tests ディレクトリ

Nodenberg API Server のテストファイルと関連リソースを格納しています。

---

## 📁 ファイル一覧

### テストスクリプト

#### `test-api.js`
**役割**: API全エンドポイントの自動テストスイート（CLIテストツール）

**機能**:
- 全6つのAPIエンドポイントを自動テスト
- テスト結果の詳細レポートを表示
- Excel/PDFファイルの生成と保存

**テスト内容**:
1. ヘルスチェック (`GET /health`)
2. プレースホルダー検出（シンプル版）
3. プレースホルダー検出（詳細版）
4. テンプレート情報取得
5. Excel生成（test9メソッド使用）
6. PDF生成（LibreOffice使用）

**実行方法**:
```bash
# デフォルト設定で実行（localhost:3000）
npm test

# カスタムAPI URLで実行
API_URL=http://localhost:3100 node tests/test-api.js

# カスタムテンプレートで実行
TEMPLATE_PATH=/path/to/template.xlsx node tests/test-api.js
```

**出力例**:
```
╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║   Nodenberg API Test Suite                                   ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝

[1/6] Health Check
      GET /health
      ✅ Success (15ms)

[2/6] Placeholder Detection
      POST /template/placeholders
      ✅ Success (125ms)
      {
        "success": true,
        "placeholders": ["会社名", "日付", "金額", "担当者"]
      }

...

╔══════════════════════════════════════════════════════════════╗
║   Test Summary                                               ║
╚══════════════════════════════════════════════════════════════╝

✅ Passed: 6
❌ Failed: 0
```

---

#### `verify-print-settings.ts`
**役割**: test9メソッドの印刷設定保存機能を検証

**機能**:
- Excel印刷設定の保存確認
- test9メソッドによるプレースホルダー置換後の設定保持検証

**検証項目**:
1. **PageSetup（ページ設定）**
   - 用紙サイズ（A4）
   - 印刷方向（横向き）
   - ページに合わせる設定
   - 余白設定

2. **RowBreaks（改ページ位置）**
   - 手動改ページの保持確認
   - ※ExcelJSの制限により完全な検証は困難

3. **Views（表示設定）**
   - ウィンドウ枠の固定
   - ズーム倍率（85%）

4. **Properties（シートプロパティ）**
   - タブの色（赤色）

**実行方法**:
```bash
npx tsx tests/verify-print-settings.ts
```

**出力例**:
```
Starting verification...
Template created with:
- PageSetup (Landscape, A4)
- RowBreaks (5, 10, 15)
- Views (Frozen, Zoom 85%)
- TabColor (Red)

Running ExcelGenerator...
✅ Verification PASSED: All settings preserved.
```

**重要性**:
test9メソッドの核心機能である「印刷設定100%保存」を検証する重要なテストです。通常のExcelライブラリ（ExcelJS、xlsx等）では印刷設定が失われる問題に対し、test9メソッドが正しく動作していることを確認します。

---

#### `verify-section-table.ts`
**役割**: 通常置換、legacy table、section複数行、sectionページ送りを検証

**検証項目**:
1. 通常プレースホルダー `{{会社名}}`
2. legacy table `{{#明細.項目}}`
3. section table `{{##請求.明細.項目}}`
4. 複数行ブロック複製時の数式行シフト
5. `Print_Area` の再計算
6. tall cell を含む section 明細のページ送り

**実行方法**:
```bash
npx tsx tests/verify-section-table.ts
```

---

#### `verify-image-placeholder.ts`
**役割**: section 内画像差し込み、複数レコード画像、旧記法拒否を検証

**検証項目**:
1. 単体 section 画像 `{{##請求.明細.image}}` の埋め込み
2. sharedStrings からの section 画像トークン除去
3. drawing / media / rels の生成
4. 同一プレースホルダ名での複数レコード画像差し込み
5. 旧 `{{%...}}` 記法の拒否

**実行方法**:
```bash
npx tsx tests/verify-image-placeholder.ts
```

---

### テンプレートファイル

#### `test-template.xlsx`
**役割**: テスト用のExcelテンプレートファイル

**内容**:
- プレースホルダー付きのサンプルExcelファイル
- `{{会社名}}`, `{{日付}}`, `{{金額}}`, `{{担当者}}` 等のプレースホルダーを含む
- 印刷設定（ヘッダー、フッター、余白等）を含む

**用途**:
- `test-api.js` のテスト実行時に使用
- APIエンドポイントのテストデータとして使用
- test9メソッドの動作確認

---

### 出力ファイル（自動生成）

#### `output-test.xlsx`
**役割**: Excel生成APIのテスト結果ファイル

**生成方法**:
`test-api.js` 実行時に自動生成されます。

**内容**:
- `test-template.xlsx` のプレースホルダーが置換されたExcelファイル
- 印刷設定が100%保持されている

**確認ポイント**:
1. プレースホルダーが正しく置換されているか
2. 印刷設定（ヘッダー、フッター、余白、ページ設定）が保持されているか
3. セルの書式設定が保持されているか

---

#### `test-output.xlsx`
**役割**: 手動テスト/別のテストで生成されたExcelファイル

**生成方法**:
手動テストやカスタムテストスクリプトによって生成されます。

**用途**:
- 手動での動作確認
- カスタムテストの結果保存

---

#### `output-test.pdf`
**役割**: PDF生成APIのテスト結果ファイル

**生成方法**:
`test-api.js` 実行時に自動生成されます（LibreOfficeが必要）。

**内容**:
- `test-template.xlsx` から生成されたPDFファイル
- プレースホルダーが置換された状態
- Excel印刷設定がPDFに反映されている

**確認ポイント**:
1. プレースホルダーが正しく置換されているか
2. Excel印刷設定（余白、ページサイズ等）がPDFに反映されているか
3. PDFのレイアウトが意図通りか

**注意**:
PDF生成には LibreOffice のインストールが必要です。インストールされていない場合、このテストはスキップされます。

```bash
# LibreOfficeインストール（Ubuntu/Debian）
sudo apt-get install libreoffice

# LibreOfficeインストール（macOS）
brew install --cask libreoffice
```

---

## 🚀 テスト実行ガイド

### 1. 全APIエンドポイントのテスト

```bash
# サーバーを起動（別ターミナル）
cd /home/ubuntu/17_ex-pdf-gen/03_docker-version
npm run dev

# テスト実行
npm test
```

### 2. 印刷設定保存の検証

```bash
npx tsx tests/verify-print-settings.ts
```

### 3. カスタムテンプレートでのテスト

```bash
# 自分のテンプレートファイルを使用
TEMPLATE_PATH=/path/to/your-template.xlsx node tests/test-api.js
```

### 4. 別サーバーへのテスト

```bash
# 本番環境やステージング環境のテスト
API_URL=https://your-production-server.com node tests/test-api.js
```

---

## ⚠️ 注意事項

### `verify-api.ts` について
- 旧Next.js版API（ポート4000）を対象としたスクリプトです
- 現在のPure Express版（ポート3000）には対応していません
- 履歴保存・参考用として残されています
- 最新のAPIテストには `test-api.js` を使用してください

### LibreOfficeの必要性
- PDF生成テストには LibreOffice が必要です
- インストールされていない場合、PDFテストは失敗します
- Excel生成テストは LibreOffice なしでも実行可能です

### テンプレートファイルの要件
- `.xlsx` 形式（Excel 2007以降）である必要があります
- プレースホルダーは `{{name}}` 形式で記述してください
- 印刷設定を保存したい場合は、Excelで事前に設定してください

---

## 📊 テスト結果の見方

### 成功例
```
[1/6] Health Check
      GET /health
      ✅ Success (15ms)
      {
        "status": "ok",
        "timestamp": "2025-12-13T07:30:00.000Z",
        ...
      }
```

### 失敗例
```
[2/6] Placeholder Detection
      POST /template/placeholders
      ❌ Failed (401)
      Error: {
        "error": "Unauthorized",
        "message": "Valid API key required. Provide it in the X-API-Key header."
      }
```

### 認証エラーの場合
APIキーが必要なエンドポイントでは、`X-API-Key` ヘッダーが必要です。
`test-api.js` を修正して APIキーを追加してください：

```javascript
const options = {
  method: test.method,
  headers: {
    'Content-Type': 'application/json',
    'X-API-Key': 'your-api-key-here'  // APIキーを追加
  }
};
```

---

## 🔗 関連ドキュメント

- [API.md](../API.md) - API仕様書
- [README.md](../README.md) - プロジェクト全体のドキュメント
- [src/lib/placeholderReplacer.ts](../src/lib/placeholderReplacer.ts) - test9メソッドの実装

---

## 📝 まとめ

このディレクトリには以下のファイルが含まれています：

| ファイル | 種類 | 用途 | 実行方法 |
|---------|------|------|----------|
| `test-api.js` | テストスクリプト | 全APIの自動テスト | `npm test` |
| `verify-api.ts` | テストスクリプト | 旧API動作確認（参考用） | `npx tsx tests/verify-api.ts` |
| `verify-print-settings.ts` | テストスクリプト | 印刷設定保存の検証 | `npx tsx tests/verify-print-settings.ts` |
| `test-template.xlsx` | テンプレート | テスト用Excelファイル | - |
| `output-test.xlsx` | 出力ファイル | Excel生成結果（自動生成） | - |
| `test-output.xlsx` | 出力ファイル | 手動テスト結果 | - |
| `output-test.pdf` | 出力ファイル | PDF生成結果（自動生成） | - |

**推奨テストフロー**:
1. `npm test` でAPIの全機能をテスト
2. `npx tsx tests/verify-print-settings.ts` で印刷設定保存を検証
3. 出力ファイル（`output-test.xlsx`, `output-test.pdf`）を目視確認
