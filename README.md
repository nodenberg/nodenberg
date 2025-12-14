# Nodenberg - Excel Report Generator

Node.js-based Excel report generation system with print settings preservation.

- Base placeholder replacement: **test9 method** (edits `xl/sharedStrings.xml`)
- Table (array) expansion + multi-page print area: **test13-like behavior**
  - Inserts rows into `xl/worksheets/sheet1.xml` when array data exceeds template rows
  - If output exceeds the first print area by 1+ rows, appends the next print area with the same height

## Features

- 🎯 **Print Settings Preservation** - keeps template print settings; for table expansion, also updates `sheet1.xml` + `workbook.xml` as needed
- ⚡ **Fast Processing** - Direct XML manipulation (~50-100ms per document)
- 🔒 **Secure** - W3C-compliant XML escaping prevents injection attacks
- 🐳 **Docker Ready** - Includes LibreOffice and Japanese fonts
- 🌐 **REST API** - Easy integration with existing systems
- 🧪 **Test Client** - Separate static client in `04_api-test-client`

## Quick Start

### Using Docker Compose (Recommended)

```bash
# 1. Clone the repository
git clone https://github.com/nodenberg/nodenberg.git
cd nodenberg

# 2. Create .env file
cp .env.example .env

# 3. Generate secure API key
echo "API_KEY=$(openssl rand -base64 32)" >> .env

# 4. Start with Docker Compose
docker compose up -d

# 5. Check status
docker compose ps
```

The application will be available at `http://localhost:3200`.

**詳細なDockerの使用方法は [DOCKER.md](DOCKER.md) を参照してください。**

### Manual Docker Setup

```bash
# Build the image
docker build -t nodenberg-api .

# Run the container
docker run -d \
  -p 3200:3100 \
  -e API_KEY=your-secret-api-key \
  -e PORT=3100 \
  --name nodenberg-api \
  nodenberg-api
```

### Local Development

```bash
# Install dependencies
npm install

# Start development server
npm run dev
```

## API Endpoints

### Express API (default: Port 3000 / Docker: Port 3200)

- **GET  /health** - Health check (no API key required)
- **POST /template/placeholders** - Detect placeholders in template
- **POST /template/info** - Get template information
- **POST /template/upload** - Upload template metadata (optional JSON template generation)
- **POST /generate/excel** - Generate Excel file (supports table expansion)
- **POST /generate/pdf** - Generate PDF file (requires LibreOffice)

See `03_docker-version/docs/API.md` for request/response examples.

### Build for Production

```bash
npm run build

# Start production server (runs LibreOffice init + starts API)
npm start
```

## Troubleshooting

### PDF Print Settings Issue on First Generation

**Problem**: The first PDF generation after server start may have incorrect print settings, but subsequent generations work correctly.

**Cause**: LibreOffice requires initialization on first run to create its user profile. During this initialization, Excel file parsing may be incomplete.

**Solution**: The application automatically initializes LibreOffice on startup (in production mode with `npm start`). For development mode, you can manually initialize:

```bash
# Manually initialize LibreOffice
npm run init-libreoffice

# Then start development server
npm run dev
```

**Docker**: The Docker image pre-initializes LibreOffice during build, so this issue won't occur in containerized deployments.

### Code Changes Not Reflected

If code changes are not reflected:

```bash
# Rebuild TypeScript output
npm run build
```

### Build Errors Not Resolving

If build errors persist even after fixing the code:

```bash
# Clean build artifacts and reinstall
rm -rf node_modules dist
npm install
npm run build
```

### Development vs Production Differences

If the app behaves differently in development (`npm run dev`) vs production (`npm start`):

```bash
# Test with production build locally
npm run build
npm start
```

**Note**: Production runs a LibreOffice initialization step on startup.

### Available Commands

| Command | Description |
|---------|-------------|
| `npm run dev` | Start development server (with cache) |
| `npm run build` | Build TypeScript to `dist/` |
| `npm start` | Initialize LibreOffice + start server |
| `npm run start:no-init` | Start server without LibreOffice init |

## Architecture

### test9 Method - Print Settings Preservation

This application uses the **test9 method** for Excel generation, which:

1. **Reads Excel as ZIP**: .xlsx files are ZIP archives containing XML files
2. **Edits sharedStrings.xml directly**: Placeholders are stored in `xl/sharedStrings.xml`
3. **Preserves print settings**:
   - For normal placeholders: `sheet1.xml` is not modified
   - For table expansion (`{{#...}}`): `sheet1.xml` and `workbook.xml` may be updated (rows inserted + Print_Area paged)
4. **XML escaping**: All user input is automatically escaped using W3C-compliant XML escaping

**Benefits**:
- ✅ Print settings are kept from the template
- ✅ Fast processing (~50-100ms)
- ✅ Secure (automatic XML escaping)
- ✅ No manual validation required

**Comparison**:

| Method | Print Settings | Speed | Security |
|--------|---------------|-------|----------|
| ExcelJS | ❌ Changes values | Medium | ✅ |
| xlsx-populate | ✅ Preserves | Medium | ✅ |
| Direct XML (test8) | ✅ Preserves | ⚡ Fast | ⚠️ Manual escaping |
| **test9 (Current)** | ✅ Preserves | ⚡ Fast | ✅ Auto-escaping |

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Repository

- **GitHub**: [https://github.com/nodenberg/nodenberg](https://github.com/nodenberg/nodenberg)
- **Issues**: [https://github.com/nodenberg/nodenberg/issues](https://github.com/nodenberg/nodenberg/issues)

## Acknowledgments

- Built with Express + TypeScript
- Uses [JSZip](https://stuk.github.io/jszip/) for Excel file manipulation
- PDF generation powered by [LibreOffice](https://www.libreoffice.org/)
