# Nodenberg API Documentation

Pure Express API Server for Excel & PDF generation with test9 method.

## Base URL

```
http://localhost:3000
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
  "placeholders": ["会社名", "日付", "金額", "担当者"]
}
```

**Detailed Response (detailed: true):**
```json
{
  "success": true,
  "placeholders": [
    {
      "placeholder": "{{会社名}}",
      "sheets": ["Sheet1"],
      "occurrences": 2
    },
    {
      "placeholder": "{{日付}}",
      "sheets": ["Sheet1"],
      "occurrences": 1
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
        "name": "Sheet1",
        "rowCount": 10,
        "columnCount": 5
      },
      {
        "name": "Sheet2",
        "rowCount": 15,
        "columnCount": 8
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

### 4. Upload Template

Upload template and optionally generate JSON template.

**Endpoint:** `POST /template/upload`

**Request Body:**
```json
{
  "templateBase64": "UEsDBBQABg...",
  "generateJson": true
}
```

**Parameters:**
- `templateBase64` (required): Base64-encoded Excel file
- `generateJson` (optional): Generate sample JSON data (default: false)

**Response:**
```json
{
  "success": true,
  "message": "Template uploaded successfully",
  "sampleData": {
    "会社名": "",
    "日付": "",
    "金額": "",
    "担当者": ""
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

### 5. Generate Excel

Generate Excel file by replacing placeholders with data.

**Endpoint:** `POST /generate/excel`

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
    }
  }' \
  --output response.json

# Extract and decode the file
cat response.json | jq -r '.data' | base64 -d > output.xlsx
```

**Important Notes:**
- Uses **test9 method** for placeholder replacement
- Preserves **100% of print settings** (headers, footers, margins, page setup)
- Direct XML manipulation via JSZip

---

### 6. Generate PDF

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
    }
  }' \
  --output response.json

# Extract and decode the file
cat response.json | jq -r '.data' | base64 -d > output.pdf
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

Open in browser:
```
http://localhost:3000/
```

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
