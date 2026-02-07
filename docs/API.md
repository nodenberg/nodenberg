# Nodenberg API Documentation

Pure Express API Server for Excel & PDF generation.

- Base placeholder replacement: **test9 method** (edit `xl/sharedStrings.xml`)
- Table (array) expansion + multi-page print area: **test13-like behavior**
  - Inserts rows into `xl/worksheets/sheet1.xml` when array data exceeds template rows
  - If output exceeds the first print area by 1+ rows, appends the next print area with the same height
    - Example: `$A$1:$Q$40` → `$A$1:$Q$40,$A$41:$Q$80`

## Base URL

Local (dev):
```
http://localhost:3000
```

Docker (compose default):
```
http://localhost:3200
```

---

## Authentication

API key authentication is **required** for all endpoints except `/health`.

### Setup

1. Create a `.env` file in the project root (copy from `.env.example`)
2. Set the `API_KEY` environment variable:
   ```bash
   API_KEY=your-secret-api-key-here
   ```
3. Restart the server

**Generate a secure API key:**
```bash
# Using openssl
openssl rand -hex 32

# Or use a strong password generator
```

### Usage

Include the API key in the `X-API-Key` header for all requests:

```bash
curl -H "X-API-Key: your-secret-api-key-here" \
  http://localhost:3000/template/placeholders
```

**Important:**
- API key is **required** for all endpoints except `/health`
- If `API_KEY` is not configured, the server will return a 500 error
- Keep your API key secret and never commit it to version control
- Use different keys for development and production

**Browser clients (CORS):**
- Preflight `OPTIONS` requests are allowed without API key.
- Actual API requests still require `X-API-Key` (except `/health`).

### Error Responses

**Unauthorized (401)** - Invalid or missing API key:
```json
{
  "error": "Unauthorized",
  "message": "Valid API key required. Provide it in the X-API-Key header."
}
```

**Server Error (500)** - API key not configured on server:
```json
{
  "error": "Server Configuration Error",
  "message": "API_KEY is not configured on the server. Please set API_KEY in .env file."
}
```

---

## Endpoints

### 1. Health Check

Check server status.

**Endpoint:** `GET /health`

**Request:**
```bash
curl http://localhost:3000/health
```

**Response:**
```json
{
  "status": "ok",
  "timestamp": "2025-12-13T07:30:00.000Z",
  "service": "Nodenberg API Server",
  "version": "0.0.1"
}
```

---

### 2. Detect Placeholders

Detect placeholders in Excel template.

**Endpoint:** `POST /template/placeholders`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg...",
  "detailed": false
}
```

**Parameters:**
- `templateBase64` (required): Base64-encoded Excel file
- `detailed` (optional): Return detailed placeholder information (default: false)

**Simple Response (detailed: false):**
```json
{
  "success": true,
  "placeholders": ["会社名", "請求日", "支払期限", "#明細.番号", "#明細.項目"]
}
```

**Detailed Response (detailed: true):**
```json
{
  "success": true,
  "placeholders": [
    {
      "placeholder": "{{会社名}}",
      "key": "会社名",
      "count": 2
    },
    {
      "placeholder": "{{#明細.番号}}",
      "key": "#明細.番号",
      "count": 9
    }
  ]
}
```

**Example:**
```bash
curl -X POST http://localhost:3000/template/placeholders \
  -H "Content-Type: application/json" \
  -H "X-API-Key: your-api-key-here" \
  -d '{
    "templateBase64": "UEsDBBQABg...",
    "detailed": true
  }'
```

*Note: Omit the `X-API-Key` header if authentication is not enabled.*

---

### 3. Get Template Info

Get template metadata (sheets, cells, etc.).

**Endpoint:** `POST /template/info`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg..."
}
```

**Response:**
```json
{
  "success": true,
  "templateInfo": {
    "sheetCount": 2,
    "sheets": [
      {
        "id": 1,
        "name": "Sheet1",
        "rowCount": 0,
        "columnCount": 0
      },
      {
        "id": 2,
        "name": "Sheet2",
        "rowCount": 0,
        "columnCount": 0
      }
    ]
  }
}
```

**Example:**
```bash
curl -X POST http://localhost:3000/template/info \
  -H "Content-Type: application/json" \
  -d '{
    "templateBase64": "UEsDBBQABg..."
  }'
```

---

### 4. Get Sheet List

Get sheet list with display order, `sheetId`, and sheet name.

**Endpoint:** `POST /template/sheets`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg..."
}
```

**Response:**
```json
{
  "success": true,
  "sheetCount": 3,
  "sheets": [
    {
      "displayOrder": 1,
      "id": 1,
      "name": "スタンダード請求書 単位あり 10％"
    },
    {
      "displayOrder": 2,
      "id": 2,
      "name": "スタンダード請求書 単位あり 10％ _2"
    },
    {
      "displayOrder": 3,
      "id": 3,
      "name": "参照シート"
    }
  ]
}
```

**Example:**
```bash
curl -X POST http://localhost:3000/template/sheets \
  -H "Content-Type: application/json" \
  -d '{
    "templateBase64": "UEsDBBQABg..."
  }'
```

---

### 5. Upload Template

Upload template and optionally generate JSON template.

**Endpoint:** `POST /template/upload`

**Request Body:**
```json
{
  "templateId": "invoice-v1",
  "templateName": "請求書テンプレート",
  "base64Data": "UEsDBBQABg...",
  "generateJsonTemplate": true
}
```

**Parameters:**
- `templateId` (required): Template identifier (string)
- `templateName` (required): Human-readable name (string)
- `base64Data` (required): Base64-encoded Excel file
- `generateJsonTemplate` (optional): Generate a JSON template by reading the workbook via ExcelJS (default: false)

**Response:**
```json
{
  "success": true,
  "template": {
    "id": "invoice-v1",
    "name": "請求書テンプレート",
    "uploadedAt": "2025-12-14T11:00:00.000Z",
    "hasJsonTemplate": true
  }
}
```

**Example:**
```bash
curl -X POST http://localhost:3000/template/upload \
  -H "Content-Type: application/json" \
  -d '{
    "templateBase64": "UEsDBBQABg...",
    "generateJson": true
  }'
```

---

### 6. Generate Excel

Generate Excel file by replacing placeholders with data.

**Endpoint:** `POST /generate/excel`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg...",
  "data": {
    "担当者": "山田太郎",
    "会社名": "テスト株式会社",
    "請求日": "2025年12月1日",
    "支払期限": "2025年12月31日",
    "明細": [
      { "番号": 1, "項目": "Webデザイン一式", "数量": 15, "単位": "個", "単価": 8000 },
      { "番号": 2, "項目": "バナー制作", "数量": 5, "単位": "個", "単価": 6000 }
    ]
  }
}
```

**Parameters:**
- `templateBase64` (required): Base64-encoded Excel template
- `data` (required): Object with placeholder replacements
- `sheetSelectBy` (optional): `"id"` or `"name"` (only effective when `sheetSelectValue` is also provided)
- `sheetSelectValue` (optional): Sheet selector value (integer for `"id"`, string for `"name"`)

**Response:**
```json
{
  "success": true,
  "data": "UEsDBBQABg..."
}
```

**Response Fields:**
- `data`: Base64-encoded generated Excel file

**Example:**
```bash
curl -X POST http://localhost:3000/generate/excel \
  -H "Content-Type: application/json" \
  -d '{
    "templateBase64": "UEsDBBQABg...",
    "data": {
      "会社名": "テスト株式会社",
      "日付": "2025/12/13",
      "金額": "¥1,234,567",
      "担当者": "山田太郎"
    },
    "sheetSelectBy": "name",
    "sheetSelectValue": "スタンダード請求書 単位あり 10％ "
  }' \
  --output response.json

# Extract and decode the file
cat response.json | jq -r '.data' | base64 -d > output.xlsx
```

**Important Notes:**
- Uses **test9 method** for normal placeholders (edits `xl/sharedStrings.xml`)
- If the template uses array placeholders like `{{#明細.項目}}`, the server may also:
  - Insert rows into `xl/worksheets/sheet1.xml` to fit array data (template row style/merge-cells are duplicated)
  - Update `xl/workbook.xml` `_xlnm.Print_Area` to add page ranges when 1+ rows overflow
- Current limitation: array expansion targets `sheet1.xml` and the first detected array only

---

### 7. Generate Excel (By Display Order)

Generate Excel by selecting one sheet using display order (1-based).

**Endpoint:** `POST /generate/excel/by-display-order`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg...",
  "data": {
    "会社名": "テスト株式会社",
    "日付": "2025/12/13",
    "金額": "¥1,234,567",
    "担当者": "山田太郎"
  },
  "displayOrder": 1
}
```

**Parameters:**
- `templateBase64` (required): Base64-encoded Excel template
- `data` (required): Object with placeholder replacements
- `displayOrder` (required): 1-based sheet display order

**Example:**
```bash
curl -X POST http://localhost:3000/generate/excel/by-display-order \
  -H "Content-Type: application/json" \
  -H "X-API-Key: your-api-key-here" \
  -d '{
    "templateBase64": "UEsDBBQABg...",
    "data": {
      "会社名": "テスト株式会社",
      "日付": "2025/12/13",
      "金額": "¥1,234,567",
      "担当者": "山田太郎"
    },
    "displayOrder": 1
  }'
```

---

### 8. Generate PDF

Generate PDF file from Excel template (requires LibreOffice).

**Endpoint:** `POST /generate/pdf`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg...",
  "data": {
    "会社名": "テスト株式会社",
    "日付": "2025/12/13",
    "金額": "¥1,234,567",
    "担当者": "山田太郎"
  }
}
```

**Parameters:**
- `templateBase64` (required): Base64-encoded Excel template
- `data` (required): Object with placeholder replacements
- `sheetSelectBy` (optional): `"id"` or `"name"` (only effective when `sheetSelectValue` is also provided)
- `sheetSelectValue` (optional): Sheet selector value (integer for `"id"`, string for `"name"`)
- `options` (optional): PDF generation options (soffice command, timeout, etc.)

**Response:**
```json
{
  "success": true,
  "data": "JVBERi0xLjQKJeLjz9..."
}
```

**Response Fields:**
- `data`: Base64-encoded generated PDF file

**Example:**
```bash
curl -X POST http://localhost:3000/generate/pdf \
  -H "Content-Type: application/json" \
  -d '{
    "templateBase64": "UEsDBBQABg...",
    "data": {
      "会社名": "テスト株式会社",
      "日付": "2025/12/13",
      "金額": "¥1,234,567",
      "担当者": "山田太郎"
    },
    "sheetSelectBy": "id",
    "sheetSelectValue": 1
  }' \
  --output response.json

# Extract and decode the file
cat response.json | jq -r '.data' | base64 -d > output.pdf
```

---

### 9. Generate PDF (By Display Order)

Generate PDF by selecting one sheet using display order (1-based).

**Endpoint:** `POST /generate/pdf/by-display-order`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg...",
  "data": {
    "会社名": "テスト株式会社",
    "日付": "2025/12/13",
    "金額": "¥1,234,567",
    "担当者": "山田太郎"
  },
  "displayOrder": 1,
  "options": {
    "timeout": 30000
  }
}
```

**Parameters:**
- `templateBase64` (required): Base64-encoded Excel template
- `data` (required): Object with placeholder replacements
- `displayOrder` (required): 1-based sheet display order
- `options` (optional): PDF generation options (soffice command, timeout, etc.)

**Example:**
```bash
curl -X POST http://localhost:3000/generate/pdf/by-display-order \
  -H "Content-Type: application/json" \
  -H "X-API-Key: your-api-key-here" \
  -d '{
    "templateBase64": "UEsDBBQABg...",
    "data": {
      "会社名": "テスト株式会社",
      "日付": "2025/12/13",
      "金額": "¥1,234,567",
      "担当者": "山田太郎"
    },
    "displayOrder": 1
  }'
```

**Requirements:**
- LibreOffice must be installed on the server
- Uses `soffice` command-line tool
- First request may be slower due to LibreOffice initialization

**Error Response (LibreOffice not installed):**
```json
{
  "error": "LibreOffice is required for PDF generation but is not installed",
  "details": "Please install LibreOffice: sudo apt-get install libreoffice"
}
```

---

## Error Responses

All endpoints return errors in the following format:

```json
{
  "error": "Error message",
  "details": "Detailed error description"
}
```

**Common HTTP Status Codes:**
- `200` - Success
- `400` - Bad Request (missing required parameters)
- `500` - Internal Server Error
- `503` - Service Unavailable (LibreOffice not installed)

---

## File Encoding

All file data is transmitted as **Base64-encoded strings**.

### Encoding a file (example in Node.js):
```javascript
const fs = require('fs');
const fileBuffer = fs.readFileSync('template.xlsx');
const templateBase64 = fileBuffer.toString('base64');
```

### Decoding a response (example in Node.js):
```javascript
const outputBuffer = Buffer.from(response.data, 'base64');
fs.writeFileSync('output.xlsx', outputBuffer);
```

### Encoding a file (example in Python):
```python
import base64

with open('template.xlsx', 'rb') as f:
    template_base64 = base64.b64encode(f.read()).decode('utf-8')
```

### Decoding a response (example in Python):
```python
import base64

output_data = base64.b64decode(response['data'])
with open('output.xlsx', 'wb') as f:
    f.write(output_data)
```

---

## Placeholder Format

Placeholders in Excel templates should use the following format:

```
{{placeholder_name}}
```

### Array (Table) placeholders

For table expansion, use:
```
{{#ArrayName.field}}
```

Example:
- `{{#明細.番号}}`
- `{{#明細.項目}}`
- `{{#明細.数量}}`
- `{{#明細.単位}}`
- `{{#明細.単価}}`

**Examples:**
- `{{会社名}}`
- `{{date}}`
- `{{total_amount}}`
- `{{customer_name}}`

**Important:**
- Placeholders are case-sensitive
- Can contain Japanese characters, letters, numbers, and underscores
- Must be wrapped in double curly braces `{{ }}`

---

## test9 Method

The **test9 method** is the core technology that preserves Excel print settings.

### How it works:

1. **Direct XML Manipulation**
   - Unzips Excel file (XLSX is a ZIP archive)
   - Directly edits `xl/sharedStrings.xml`
   - Replaces placeholder text in the XML
   - Re-zips the file

2. **Why it's superior:**
   - **100% preservation** of print settings
   - Headers, footers, margins, page breaks all maintained
   - No library limitations (ExcelJS, xlsx, etc. lose settings)

3. **Implementation:**
   - Located in: `src/lib/placeholderReplacer.ts`
   - Uses: JSZip for ZIP manipulation
   - Pattern: `{{placeholder}}` detection via regex

---

## Multi-page Print Area (test13-like)

If the template has a first print area (via `xl/workbook.xml` `_xlnm.Print_Area`) and content exceeds it:
- The server appends the next print area with the same height.
- Example (page height = 40 rows): `$A$1:$Q$40` → `$A$1:$Q$40,$A$41:$Q$80`

---

## Testing

### CLI Test Suite

Run automated tests:
```bash
npm test
```

**Tests included:**
- Health check
- Placeholder detection (simple)
- Placeholder detection (detailed)
- Template info
- Excel generation
- PDF generation

### GUI Test Client

This project does not ship a Next.js UI.

Use the separate static test client in `04_api-test-client` (served by `http-server`).

**Features:**
- File upload
- Placeholder detection
- Template info display
- JSON data input
- Excel/PDF generation
- File download

---

## Development

### Start development server (hot-reload):
```bash
npm run dev
```

### Build for production:
```bash
npm run build
```

### Start production server:
```bash
npm start
```

### Initialize LibreOffice (for PDF generation):
```bash
npm run init-libreoffice
```

---

## Rate Limits

Currently **no rate limits** are implemented.

For production use, consider adding:
- Rate limiting middleware (express-rate-limit)
- Authentication (JWT)
- Request size limits (already set to 50MB)

---

## CORS

CORS is enabled for all origins:
```typescript
app.use(cors());
```

For production, restrict to specific origins:
```typescript
app.use(cors({
  origin: 'https://yourdomain.com'
}));
```

---

## Support

For issues or questions:
- GitHub: https://github.com/nodenberg/nodenberg
- Documentation: This file (API.md)

---

## Version

**Current Version:** 0.0.1

**Last Updated:** 2025-12-13
