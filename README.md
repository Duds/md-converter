# MD Converter

Convert Markdown files to DOCX, XLSX, and PPTX formats with Excel formula support. Includes both CLI and MCP (Model Context Protocol) server interfaces.

## Features

✨ **Multiple Output Formats**
- 📄 **DOCX**: Word documents with proper formatting, headings, tables, lists, and code blocks
- 📊 **XLSX**: Excel spreadsheets with formula support, data type detection, and formatting
- 📽️ **PPTX**: PowerPoint presentations with automatic slide layouts

🔢 **Excel Formula Support**
- Convert `{=FORMULA}` syntax in Markdown tables to actual Excel formulas
- Support for 60+ Excel functions (SUM, AVERAGE, IF, VLOOKUP, etc.)
- Automatic data type detection (numbers, dates, booleans, text)
- Cell reference validation

🤖 **Dual Interface**
- **CLI**: Command-line tool for batch processing
- **MCP Server**: Integration with AI assistants like Claude in Cursor

🎨 **Formatting Options**
- Customisable fonts and sizes
- Auto-width columns in Excel
- Frozen header rows
- Light/dark themes for presentations
- Australian date format (DD/MM/YYYY)

## Installation

```bash
cd /Users/dalerogers/Projects/active/experimental/md_converter
npm install
npm run build
```

## CLI Usage

### Basic Conversion

```bash
# Convert to DOCX
md-convert input.md --format docx

# Convert to XLSX
md-convert input.md --format xlsx

# Convert to PPTX
md-convert input.md --format pptx

# Convert to all formats
md-convert input.md --format all
```

### Batch Processing

```bash
# Convert all markdown files in a directory
md-convert "./docs/**/*.md" --format all --output-dir ./exports
```

### Options

```bash
# XLSX with custom options
md-convert report.md --format xlsx --freeze-headers --auto-width

# PPTX with dark theme
md-convert slides.md --format pptx --theme dark

# Custom fonts
md-convert document.md --format docx --font-family "Arial" --font-size 12
```

### Preview and Validation

```bash
# Preview table extraction
md-convert preview report.md

# Validate formulas
md-convert validate report.md
```

## MCP Server Usage

### Start the Server

```bash
md-convert serve
```

### Configure in Cursor

Add to your Cursor settings (`.cursor/config.json` or user settings):

```json
{
  "mcpServers": {
    "md-converter": {
      "command": "node",
      "args": ["/Users/dalerogers/Projects/active/experimental/md_converter/dist/cli/index.js", "serve"],
      "disabled": false
    }
  }
}
```

### Available MCP Tools

1. **convert_md_to_docx** - Convert Markdown to Word document
2. **convert_md_to_xlsx** - Convert Markdown tables to Excel with formulas
3. **convert_md_to_pptx** - Convert Markdown to PowerPoint presentation
4. **preview_tables** - Preview table extraction and formulas
5. **validate_formulas** - Validate formula syntax

## Formula Syntax

Use `{=FORMULA}` in Markdown tables to create Excel formulas:

```markdown
| Product | Price | Quantity | Total |
|---------|-------|----------|-------|
| Apples  | 2.50  | 10       | {=B2*C2} |
| Oranges | 3.00  | 5        | {=B3*C3} |
| **Total** | | | {=SUM(D2:D3)} |
```

This converts to actual Excel formulas that calculate when the file is opened.

### Supported Formula Features

- **Basic arithmetic**: `{=B2+C2}`, `{=B2*C2}`
- **Functions**: `{=SUM(A1:A10)}`, `{=AVERAGE(B:B)}`
- **Cell references**: `A1`, `$A$1` (absolute)
- **Ranges**: `A1:A10`, `B:B` (entire column), `1:10` (rows)
- **Nested functions**: `{=IF(A1>0, SUM(B1:B10), 0)}`

### Supported Excel Functions

**Math**: SUM, AVERAGE, COUNT, MIN, MAX, ROUND, ABS, SQRT, POWER, MOD, PRODUCT

**Logical**: IF, AND, OR, NOT, XOR, IFERROR, IFNA

**Text**: CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER

**Date**: TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF, DAYS

**Lookup**: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, CHOOSE

**Statistical**: STDEV, VAR, RANK, PERCENTILE

**Financial**: PMT, FV, PV, RATE, NPV, IRR

## PowerPoint Slide Structure

Use headings and horizontal rules to structure presentations:

```markdown
# Main Title
Subtitle text

---

## Section Header

---

### Slide Title

- Bullet point 1
- Bullet point 2

---

### Data Slide

| Header 1 | Header 2 |
|----------|----------|
| Data 1   | Data 2   |
```

**Mapping**:
- `# H1` → Title slide
- `## H2` → Section header slide
- `### H3` → Content slide title
- `---` → Slide separator
- Lists → Bullet points
- Tables → PowerPoint tables

## Examples

See the `examples/` directory:
- **sample.md**: Sales report with formulas
- **presentation.md**: Multi-slide presentation

Try converting them:

```bash
md-convert examples/sample.md --format all
md-convert examples/presentation.md --format pptx
```

## Data Type Detection

The converter automatically detects cell data types:

- **Numbers**: `45000`, `$2,450`, `3.14`
- **Dates**: `08/11/2025` (DD/MM/YYYY - Australian format)
- **Booleans**: `TRUE`, `FALSE`
- **Formulas**: `{=SUM(A1:A10)}`
- **Text**: Everything else

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Run CLI in dev mode
npm run dev -- input.md --format docx

# Start MCP server in dev mode
npm run serve

# Type check
npm run type-check
```

## Project Structure

```
md_converter/
├── src/
│   ├── core/
│   │   ├── parsers/          # Markdown, table, and formula parsing
│   │   └── converters/       # DOCX, XLSX, and PPTX converters
│   ├── mcp/                  # MCP server implementation
│   ├── cli/                  # CLI interface
│   └── index.ts              # Main exports
├── examples/                 # Example markdown files
├── dist/                     # Built JavaScript (generated)
└── tests/                    # Test files

```

## Technical Details

**Markdown Parsing**: Uses `markdown-it` for robust AST-based parsing

**DOCX Generation**: Uses `docx` library with proper styling and formatting

**XLSX Generation**: Uses `exceljs` with native formula support

**PPTX Generation**: Uses `pptxgenjs` with flexible slide layouts

**MCP Integration**: Built on official `@modelcontextprotocol/sdk`

## Limitations

- DOCX: Formulas displayed as text (Word doesn't support Excel formulas)
- PPTX: Complex tables may need manual adjustment
- Images: Not yet supported (coming soon)
- Large files: May take time to process (batch mode recommended)

## Contributing

This is an experimental project. Contributions welcome!

## Licence

MIT

## Author

Dale Rogers

---

**Note**: This project uses Australian English spelling and date formats (DD/MM/YYYY) throughout.

