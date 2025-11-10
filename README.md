# MD Converter

**Version:** 2.1.0  
**Status:** Production Ready

Convert Markdown files to DOCX, XLSX, and PPTX formats with YAML front matter, strict validation, intelligent section breaks, and professional Word styling. Includes both CLI and MCP (Model Context Protocol) server interfaces.

---

## Why This Matters: Consulting by IDE

### The IDE-First Consulting Approach

Modern consulting increasingly happens **inside the IDE** rather than traditional office suites. AI assistants like Claude in Cursor allow consultants to:

- **Think and write in Markdown** - Clean, version-controlled, collaborative
- **Leverage AI assistance** - Real-time document generation and refinement
- **Work in plain text** - Fast, lightweight, Git-friendly
- **Maintain single source of truth** - Markdown is the master, outputs are generated

### The Problem

Clients and stakeholders need documents in **traditional formats** (Word, Excel, PowerPoint):
- Executive briefings require Word templates
- Financial models need Excel workbooks
- Presentations need PowerPoint decks

Converting manually wastes time and breaks the IDE-first workflow.

### The Solution: MD Converter MCP

**Seamless conversion from IDE to client deliverables:**

1. **Consultant works in Markdown** - Full IDE power, AI assistance, version control
2. **Metadata in YAML** - Professional document properties automatically applied
3. **One command conversion** - Instant professional Word/Excel/PowerPoint
4. **Quality assurance** - Built-in validation ensures consistency
5. **Smart exclusions** - System files, notes, and references stay in markdown

**Result:** Consultants stay in flow, clients get professional deliverables.

### Real-World Example: DCCEEW AI Capability Project

**Challenge:** Deliver comprehensive AI capability planning to government client
- 22 planning documents covering architecture, delivery, resourcing, governance
- Executive brief for CIO decision
- Financial model for $8.9M programme
- Professional formatting required (OFFICIAL classification, proper metadata)

**Solution:** 
- All planning done in Markdown with AI assistance
- YAML front matter for classification, versions, authors
- Single command: `md-convert "**/*.md"` → 17 DOCX + 1 XLSX
- Intelligent exclusion of reference materials and system files
- Section breaks at major boundaries for independent formatting
- Professional Word styles for template compatibility

**Outcome:** 
- Planning completed in days, not weeks
- All documents professionally formatted
- Consistent metadata across deliverables
- Ready for executive distribution

---

## Features

### 🚀 Version 2.1.0 Enhancements

**YAML Front Matter Support**
- 14 metadata fields (format, title, author, date, classification, version, status, etc.)
- Automatic mapping to document properties in Word/Excel/PowerPoint
- Format detection (docx, xlsx, pptx, or combinations)
- Document type categorization (document, email, reference, note, system)

**Strict Document Validation**
- Heading hierarchy validation (no skipped levels)
- Table structure consistency checks
- Empty heading detection
- Metadata completeness warnings
- Optional `--strict` mode for production documents

**Intelligent Section Breaks**
- DOCX: `section_breaks: auto` only creates sections before ## H2 (major boundaries)
- PPTX: `slide_breaks: h1|h2|hr` controls slide creation
- Reduces unnecessary section breaks
- Enables independent section formatting

**Smart Exclusion Rules**
- Path-based: Auto-exclude README, /notes/, /references/
- Front matter: `convert: false` or `document_type: email|reference|note|system`
- Cleaner batch conversions
- Clear skip reasons in output

**Professional Word Styles**
- Built-in styles: Normal, Heading 1-6, ListParagraph
- Character styles: Strong (bold), Emphasis (italic)
- Template compatible
- Easy bulk reformatting

**Numbered Lists Fixed**
- Proper Word numbering configuration
- Sequential numbering (1, 2, 3...)
- No more DOCX corruption

**XLSX Rich Text**
- Bold text (`**text**`) properly renders in Excel
- Italic text (`*text*`) properly renders
- Mixed formatting in cells works correctly

### 📄 Multiple Output Formats

- **DOCX**: Word documents with proper formatting, headings, tables, lists, and code blocks
- **XLSX**: Excel spreadsheets with formula support, data type detection, and formatting
- **PPTX**: PowerPoint presentations with automatic slide layouts

### 🔢 Excel Formula Support

- Convert `{=FORMULA}` syntax in Markdown tables to actual Excel formulas
- Support for 60+ Excel functions (SUM, AVERAGE, IF, VLOOKUP, etc.)
- Automatic data type detection (numbers, dates, booleans, text)
- Cell reference validation

### 🤖 Dual Interface

- **CLI**: Command-line tool for batch processing
- **MCP Server**: Integration with AI assistants like Claude in Cursor

### 🎨 Formatting Options

- Customizable fonts and sizes
- Auto-width columns in Excel
- Frozen header rows
- Light/dark themes for presentations
- Australian date format (DD/MM/YYYY)

---

## Installation

```bash
# Clone or navigate to the repository
cd /path/to/md_converter

# Install dependencies
npm install

# Build the project
npm run build
```

---

## Quick Start

### 1. Add Front Matter to Your Markdown

```markdown
---
format: docx
title: "Executive Brief - AI Strategy"
author: "Dale Rogers"
date: "2025-11-10"
classification: "OFFICIAL"
version: "1.0"
status: final
keywords: ["AI", "strategy", "executive"]
section_breaks: auto
---

# Executive Brief

Your content here...
```

### 2. Convert to Word

```bash
md-convert document.md --format docx
```

### 3. Review Output

- Open the generated `.docx` file
- Check File → Info → Properties for metadata
- Verify styles in Home → Styles pane
- Check section breaks (View → Draft mode)

---

## YAML Front Matter Reference

### Required Fields

```yaml
format: docx              # docx, xlsx, pptx, or combinations
title: "Document Title"   # Document title (used in properties)
```

### Extended Fields (Recommended)

```yaml
author: "Your Name"                      # Creator/author name
date: "2025-11-10"                      # Document date (YYYY-MM-DD)
classification: "OFFICIAL"               # Security classification
version: "1.0"                          # Version number
status: final                           # draft|review|approved|final
description: "Brief description"         # Document description (<250 chars)
keywords: ["key1", "key2"]              # Search keywords (array)
subject: "Category"                     # Document subject/category
```

### Document Control Fields

```yaml
convert: false                          # Skip conversion (default: true)
document_type: reference                # document|email|reference|note|system
```

### Format-Specific Fields

```yaml
# DOCX only
section_breaks: auto                    # auto|all|none

# PPTX only
slide_breaks: h2                        # h1|h2|hr
```

---

## Exclusion Rules

### Path-Based Exclusion (Automatic)

These files are automatically skipped:

- `README.md` (any location, case-insensitive)
- Files in `/notes/` directory
- Files in `/reference/` or `/references/` directory

### Front Matter Exclusion

**Method 1: Explicit flag**
```yaml
convert: false
```

**Method 2: Document type (automatic)**
```yaml
document_type: email     # Auto-skips: email, reference, note, system
```

### Example Output

```
Found 24 file(s) to convert

Skipping: README.md (excluded by path pattern)
Skipping: notes/meeting-notes.md (excluded by path pattern)
Converting: EXECUTIVE_BRIEF.md
  ✓ DOCX: EXECUTIVE_BRIEF.docx
Converting: budget.md
  ⊘ Skipped (document_type: reference)

Conversion Summary:
  ✓ Success: 18
  ⊘ Skipped: 6
```

---

## Section Break Rules (DOCX)

Control how `---` horizontal rules are handled:

### `section_breaks: auto` (Recommended)

**Rule:** Section break ONLY if `---` is followed by `## H2`

**Example:**
```markdown
## Major Section A

Content here...

---

## Major Section B

New section with independent formatting
```

**Result:** Section break created (can have different headers/footers)

```markdown
Content here...

---

More content in same section (just visual divider)
```

**Result:** Visual divider only (no section break)

### `section_breaks: all`

Every `---` creates a Word section break.

**Use when:** Need maximum control over section formatting

### `section_breaks: none`

All `---` are visual dividers only.

**Use when:** Simple document, no section independence needed

---

## Slide Break Rules (PPTX)

Control when new slides are created:

### `slide_breaks: h2` (Default)

New slide at every `# H1` and `## H2`

**Best for:** Standard presentations with clear sections

### `slide_breaks: h1`

New slide only at `# H1`

**Best for:** Dense presentations with more content per slide

### `slide_breaks: hr`

New slide at every `---` horizontal rule

**Best for:** Manual control over slide breaks

---

## CLI Usage

### Basic Conversion

```bash
# Convert to DOCX (Word)
md-convert input.md --format docx

# Convert to XLSX (Excel) - tables only
md-convert input.md --format xlsx

# Convert to PPTX (PowerPoint)
md-convert input.md --format pptx

# Convert to all formats
md-convert input.md --format all
```

### Batch Processing

```bash
# Convert all markdown files
md-convert "**/*.md" --format docx

# Convert with output directory
md-convert "./docs/**/*.md" --format all --output-dir ./exports
```

### Validation

```bash
# Standard mode (shows warnings, continues)
md-convert document.md --format docx

# Strict mode (fails on warnings)
md-convert document.md --format docx --strict

# Skip validation (faster, not recommended)
md-convert document.md --format docx --no-validate
```

### Options

```bash
# XLSX options
md-convert report.md --format xlsx --freeze-headers --auto-width

# PPTX options
md-convert slides.md --format pptx --theme dark

# Font options
md-convert document.md --format docx --font-family "Calibri" --font-size 11
```

### Commands

```bash
# Preview table extraction
md-convert preview report.md

# Validate formulas
md-convert validate report.md

# Start MCP server
md-convert serve
```

---

## MCP Server Usage

### Configure in Cursor

Add to `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "md-converter": {
      "command": "node",
      "args": [
        "/Users/dalerogers/Projects/active/experimental/md_converter/dist/cli/index.js",
        "serve"
      ],
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

### Using in Cursor with Claude

```
You: "Convert this markdown to Word"
Claude: [Uses convert_md_to_docx tool]

You: "Create Excel from these budget tables"
Claude: [Uses convert_md_to_xlsx tool]
```

**Benefit:** Seamless conversion without leaving your IDE workflow.

---

## Word Document Features

### Styles Applied

**Paragraph Styles:**
- `Normal` - Body text (base style for all content)
- `Heading 1` through `Heading 6` - Structural headings
- `ListParagraph` - All list items (bullets and numbered)

**Character Styles:**
- `Strong` - Bold text (`**text**`)
- `Emphasis` - Italic text (`*text*`)

**Benefits:**
- ✅ Compatible with corporate Word templates
- ✅ Easy bulk reformatting (change Normal style → all text updates)
- ✅ Follows Microsoft Word best practices
- ✅ Professional appearance

### Section Breaks

Documents are split into Word sections at major boundaries:

**Features per section:**
- Independent headers and footers
- Different page numbering (restart or different style)
- Different page orientation (portrait/landscape mix)
- Different margins
- Different page sizes

**Viewing section breaks:**
- View → Draft mode → Shows "Section Break (Next Page)"
- Status bar shows "Section X of Y"

---

## Excel Workbook Features

### Table Conversion

Each markdown table becomes a separate worksheet:

**Features:**
- Headers with grey background and bold text
- Frozen header row (scroll data, headers stay visible)
- Auto-width columns (fits content automatically)
- Proper borders and gridlines

### Formula Support

Use `{=FORMULA}` in markdown tables:

```markdown
| Product | Price | Qty | Total |
|---------|-------|-----|-------|
| Apples  | 2.50  | 10  | {=B2*C2} |
| Oranges | 3.00  | 5   | {=B3*C3} |
| **Total** | | | {=SUM(D2:D3)} |
```

Converts to actual Excel formulas that calculate automatically.

### Rich Text Formatting

Bold and italic formatting from markdown:

```markdown
| Item | Value |
|------|-------|
| **Total Cost** | $1,500 |
| *Estimated* | $1,200 |
```

**Result:** Bold and italic text display correctly in Excel cells.

---

## PowerPoint Presentation Features

### Slide Structure

Headings and horizontal rules control slide breaks:

```markdown
# Main Title
Presentation subtitle

---

## Section Header Slide

---

### Content Slide

- Bullet point 1
- Bullet point 2

---
```

**Mapping:**
- `# H1` → Title slide
- `## H2` → Section header slide
- `### H3` → Content slide with title
- `---` → Slide separator (if `slide_breaks: hr`)

### Themes

```bash
# Light theme (default)
md-convert slides.md --format pptx --theme light

# Dark theme
md-convert slides.md --format pptx --theme dark
```

---

## Validation

### What's Validated

**YAML Front Matter:**
- Valid YAML syntax
- Required fields present (format, title)
- Valid enum values (format, status, document_type)
- Date format (YYYY-MM-DD recommended)

**Content Structure:**
- Heading hierarchy (no skipped levels: H1 → H2 → H3)
- Empty headings
- Table column consistency
- Orphaned lists

**Metadata Completeness:**
- Warns if missing recommended fields (author, date, classification, version, keywords)
- Warns if description > 250 characters

### Validation Modes

**Standard (default):**
```bash
md-convert document.md --format docx
```
- Shows warnings in yellow
- Continues conversion
- Recommended for most use

**Strict mode:**
```bash
md-convert document.md --format docx --strict
```
- Treats warnings as errors
- Fails conversion on any warning
- Use for production/client documents

**No validation:**
```bash
md-convert document.md --format docx --no-validate
```
- Skips all validation
- Faster for trusted documents
- Not recommended for client deliverables

---

## Complete Front Matter Example

```yaml
---
# Required fields
format: docx                                    # Output format
title: "Executive Brief - AI Strategy"          # Document title

# Extended metadata (recommended)
author: "Dale Rogers"                           # Author name
date: "2025-11-10"                             # Date (YYYY-MM-DD)
classification: "OFFICIAL"                      # Security level
version: "1.0"                                 # Version number
status: final                                  # draft|review|approved|final
description: "Executive decision brief"         # Brief description
keywords: ["AI", "strategy", "executive"]      # Search keywords
subject: "Executive Briefing"                  # Category

# Document control (optional)
convert: true                                  # Enable conversion (default)
document_type: document                        # document|email|reference|note|system

# Format-specific (optional)
section_breaks: auto                           # auto|all|none (DOCX)
slide_breaks: h2                               # h1|h2|hr (PPTX)
---

# Your Document Starts Here

Content goes here...
```

**See:** `FRONTMATTER.md` for complete specification

---

## Formula Syntax

### Basic Usage

```markdown
| Item | Amount | Tax | Total |
|------|--------|-----|-------|
| Product A | 100 | {=B2*0.1} | {=B2+C2} |
| Product B | 200 | {=B3*0.1} | {=B3+C3} |
| **Total** | {=SUM(B2:B3)} | {=SUM(C2:C3)} | {=SUM(D2:D3)} |
```

### Supported Functions

**Math:** SUM, AVERAGE, COUNT, MIN, MAX, ROUND, ABS, SQRT, POWER, MOD, PRODUCT

**Logical:** IF, AND, OR, NOT, XOR, IFERROR, IFNA

**Text:** CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER

**Date:** TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF, DAYS

**Lookup:** VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, CHOOSE

**Statistical:** STDEV, VAR, RANK, PERCENTILE

**Financial:** PMT, FV, PV, RATE, NPV, IRR

### Cell References

```markdown
{=A1}           # Simple reference
{=$A$1}         # Absolute reference
{=A1:A10}       # Range
{=B:B}          # Entire column
{=1:10}         # Row range
```

---

## Best Practices

### For Consulting Deliverables

1. **Always add YAML front matter**
   - Provides professional metadata
   - Enables format detection
   - Shows document ownership and version

2. **Use `section_breaks: auto` for DOCX**
   - Clean section structure
   - Major sections get independence
   - Avoids over-sectioning

3. **Set proper classification**
   - OFFICIAL, PROTECTED, etc.
   - Shows in document properties
   - Important for government/enterprise

4. **Include version and status**
   - Tracks document maturity
   - `status: draft` for work-in-progress
   - `status: final` for client delivery

5. **Use keywords for searchability**
   - 3-6 keywords per document
   - Helps clients find documents later
   - Good for document management systems

### For Multi-Format Output

**Strategy documents** → DOCX only
```yaml
format: docx
```

**Financial models** → XLSX only (for formulas)
```yaml
format: xlsx
```

**Executive presentations** → DOCX + PPTX
```yaml
format: docx,pptx
```

**Comprehensive deliverable** → All formats
```yaml
format: all
```

### For Project Organization

**Deliverables** → `format: docx`, `status: final`

**References** → `document_type: reference` (auto-excluded)

**Working notes** → In `/notes/` directory (auto-excluded)

**System docs** → `convert: false` or `document_type: system`

---

## Examples

### Government Executive Brief

```yaml
---
format: docx
title: "Executive Brief – AI Capability Enablement"
author: "Dale Rogers"
date: "2025-11-07"
classification: "OFFICIAL"
version: "1.0"
status: final
description: "Executive decision brief presenting three delivery options"
keywords: ["AI", "capability", "executive brief", "decision"]
subject: "Executive Briefing"
section_breaks: auto
---

# Executive Brief

## Situation
...
```

### Financial Model

```yaml
---
format: xlsx
title: "ROM Costing Model"
author: "Dale Rogers"
date: "2025-11-07"
classification: "OFFICIAL"
version: "1.0"
status: final
description: "Budget model with detailed cost breakdown"
keywords: ["ROM", "budget", "costing", "financial"]
subject: "Financial Planning"
---

# Budget Model

## Phase 1 Costs

| Item | Cost |
|------|------|
| Labour | {=SUM(B2:B10)} |
```

### Reference Material

```yaml
---
format: docx
title: "Background Research"
author: "Team"
date: "2025-11-07"
classification: "OFFICIAL"
version: "0.5"
status: draft
document_type: reference        # Auto-excluded from batch conversion
description: "Background research and analysis"
keywords: ["research", "background"]
subject: "Reference"
---
```

---

## Validation Reference

### Common Errors (Fail Conversion)

❌ Missing required field: `format` or `title`  
❌ Invalid format value  
❌ Heading hierarchy violation (H1 → H3, skipping H2)  
❌ Empty heading  
❌ Table column count mismatch

### Common Warnings (Continue with Warning)

⚠️ Missing recommended field (author, date, version)  
⚠️ Date not in YYYY-MM-DD format  
⚠️ Description > 250 characters  
⚠️ No H1 title in document  
⚠️ Code block without language tag  
⚠️ Empty list

### Viewing Validation Output

```
Converting: document.md
  ⚠ Warnings:
    • Recommended field missing: author
    • Recommended field missing: version
  📄 Executive Brief (docx)
     v1.0 - final
  ✓ DOCX: document.docx
```

---

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Run CLI in dev mode (with TypeScript)
npm run dev -- input.md --format docx

# Start MCP server in dev mode
npm run serve

# Type check (without building)
npm run type-check

# Watch mode (auto-rebuild on changes)
tsc --watch
```

---

## Project Structure

```
md_converter/
├── src/
│   ├── core/
│   │   ├── parsers/
│   │   │   ├── markdown.ts              # Markdown parsing
│   │   │   ├── frontmatter-parser.ts    # YAML front matter
│   │   │   ├── table-parser.ts          # Table processing
│   │   │   └── formula-parser.ts        # Formula validation
│   │   ├── converters/
│   │   │   ├── docx-converter.ts        # Word generation
│   │   │   ├── xlsx-converter.ts        # Excel generation
│   │   │   ├── pptx-converter.ts        # PowerPoint generation
│   │   │   └── section-rules.ts         # Section/slide break logic
│   │   └── validators/
│   │       └── document-validator.ts     # Document validation
│   ├── mcp/
│   │   ├── server.ts                    # MCP server
│   │   └── tools.ts                     # MCP tool definitions
│   ├── cli/
│   │   └── index.ts                     # CLI interface
│   └── index.ts                         # Main exports
├── examples/
│   ├── frontmatter-template.md          # Template for new docs
│   ├── sample.md                        # Formula example
│   └── presentation.md                  # Presentation example
├── FRONTMATTER.md                       # Complete specification
├── CHANGELOG.md                         # Version history
├── README.md                            # This file
├── package.json                         # Dependencies
└── tsconfig.json                        # TypeScript config
```

---

## Technical Details

**Markdown Parsing:** `markdown-it` - Robust AST-based parsing

**DOCX Generation:** `docx` library - Professional Word documents with:
- Built-in Word styles (Normal, Heading 1-6, ListParagraph, Strong, Emphasis)
- Proper numbering configuration for lists
- Section management for independent formatting
- Document properties (title, author, subject, keywords, classification)

**XLSX Generation:** `exceljs` - Excel workbooks with:
- Native formula support (actual Excel formulas, not text)
- Rich text formatting (bold, italic in cells)
- Multiple worksheets (one per markdown table)
- Auto-width columns and frozen headers
- Document properties

**PPTX Generation:** `pptxgenjs` - PowerPoint presentations with:
- Flexible slide layouts
- Theme support (light/dark)
- Tables and bullet points
- Document properties

**YAML Parsing:** `js-yaml` - Industry standard YAML parser

**MCP Integration:** `@modelcontextprotocol/sdk` - Official Model Context Protocol SDK

---

## Limitations

### Formula Support by Format

| Format | Formula Support | Description |
|--------|----------------|-------------|
| **XLSX** | ✅ **Full Support** | Formulas converted to actual Excel formulas that calculate automatically |
| **DOCX** | ❌ Text Only | Formulas displayed as plain text (Word doesn't support cell calculations) |
| **PPTX** | ❌ Text Only | Formulas displayed as plain text in table cells |

**Important:** For working formulas, always use `format: xlsx`.

### Other Limitations

- **Images:** Not yet supported (planned for v2.2)
- **Complex nested lists:** Limited to single-level lists
- **Very large tables:** May need manual adjustment in PPTX
- **Custom Word templates:** Not yet supported (uses built-in styles only)
- **Mermaid diagrams:** Not supported (render as code blocks)

---

## Troubleshooting

### "No tables found" error

**Solution:** XLSX format requires markdown tables. Add at least one table to your document.

### Bold text not showing in Excel

**Solution:** Use v2.1.0 or later. Earlier versions had a rich text formatting bug (now fixed).

### Numbered lists not working in Word

**Solution:** Use v2.1.0 or later. Earlier versions were missing numbering configuration (now fixed).

### Section breaks everywhere in Word

**Solution:** Use `section_breaks: auto` instead of `section_breaks: all`. Auto mode only creates sections at major boundaries (before ## H2).

### Files being skipped during conversion

**Solution:** Check if file matches exclusion rules:
- Is it README.md?
- Is it in /notes/ or /references/ directory?
- Does front matter have `convert: false`?
- Does front matter have `document_type: email|reference|note|system`?

---

## Contributing

This is an active project. Contributions welcome!

**Areas for contribution:**
- Image support
- Custom Word templates
- Multi-level list support
- Mermaid diagram rendering
- Additional formula functions
- Enhanced validation rules

---

## Changelog

See `CHANGELOG.md` for version history.

---

## License

MIT

---

## Author

Dale Rogers  
Service Design Lead  
2025

---

## Acknowledgments

Built for real-world consulting needs, refined through production use on Australian Government projects.

**Special thanks to the DCCEEW AI Capability Enablement project** for providing the demanding use case that drove these enhancements.

---

**Note:** This project uses Australian English spelling and date formats (DD/MM/YYYY) throughout.
