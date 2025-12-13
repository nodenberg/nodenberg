# Nodenberg - Excel Report Generator

Node.js-based Excel report generation system with print settings preservation. This application uses the innovative **test9 method** to maintain 100% print settings integrity while replacing placeholders in Excel templates.

## Features

- 🎯 **100% Print Settings Preservation** - test9 method maintains all Excel print settings
- ⚡ **Fast Processing** - Direct XML manipulation (~50-100ms per document)
- 🔒 **Secure** - W3C-compliant XML escaping prevents injection attacks
- 🐳 **Docker Ready** - Includes LibreOffice and Japanese fonts
- 🌐 **REST API** - Easy integration with existing systems
- 🖥️ **Web UI** - User-friendly interface for template upload and generation

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

### Next.js API Routes (Port 3000)

- **POST /api/template/placeholders** - Detect placeholders in template
- **POST /api/template/info** - Get template information
- **POST /api/generate/excel** - Generate Excel file
- **POST /api/generate/pdf** - Generate PDF file (requires LibreOffice)

### Build for Production

```bash
# Clean build (recommended)
npm run build:clean

# Normal build
npm run build

# Start production server
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

If your code changes are not reflected in the development server, clear the Next.js cache:

```bash
# Clear cache and restart dev server
npm run clean
npm run dev

# Or clean build and start
npm run build:clean
npm run dev
```

### Build Errors Not Resolving

If build errors persist even after fixing the code:

```bash
# Complete clean build
rm -rf node_modules .next node_modules/.cache
npm install
npm run build
```

### Development vs Production Differences

If the app behaves differently in development (`npm run dev`) vs production (`npm start`):

```bash
# Test with production build locally
npm run build:clean
npm start
```

**Note**: Always use `npm run build:clean` before deploying to avoid cache-related issues.

### Available Commands

| Command | Description |
|---------|-------------|
| `npm run dev` | Start development server (with cache) |
| `npm run build` | Build for production (with cache) |
| `npm run build:clean` | Clean build (removes .next and cache) |
| `npm run start` | Start production server |
| `npm run clean` | Remove cache only |
| `npm run lint` | Run ESLint |

## Architecture

### test9 Method - Print Settings Preservation

This application uses the **test9 method** for Excel generation, which:

1. **Reads Excel as ZIP**: .xlsx files are ZIP archives containing XML files
2. **Edits sharedStrings.xml directly**: Placeholders are stored in `xl/sharedStrings.xml`
3. **Preserves print settings**: `xl/worksheets/sheet1.xml` is never modified
4. **XML escaping**: All user input is automatically escaped using W3C-compliant XML escaping

**Benefits**:
- ✅ 100% print settings preservation (no normalization)
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

- Built with [Next.js](https://nextjs.org/)
- Uses [JSZip](https://stuk.github.io/jszip/) for Excel file manipulation
- PDF generation powered by [LibreOffice](https://www.libreoffice.org/)
