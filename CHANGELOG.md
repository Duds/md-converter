# Changelog

All notable changes to MD Converter will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.1.0] - 2025-11-10

### Major Enhancements

#### YAML Front Matter Support
- Added comprehensive YAML front matter parsing with 14 metadata fields
- Automatic metadata mapping to document properties (Word/Excel/PowerPoint)
- Format detection from front matter (docx, xlsx, pptx, or combinations)
- Document type categorization for automatic exclusion

#### Strict Document Validation
- Heading hierarchy validation (no skipped levels)
- Table structure consistency checks
- Empty heading detection
- Metadata completeness warnings
- New `--strict` CLI flag for production use

#### Intelligent Section Break Handling
- Smart section breaks in DOCX (auto mode: only before ## H2)
- Configurable modes: auto, all, none
- Slide break rules for PPTX (h1, h2, hr modes)
- Reduces unnecessary section breaks

#### Smart Exclusion Rules
- Path-based auto-exclusion (README, /notes/, /references/)
- Front matter exclusion (convert: false or document_type)
- Clear skip reasons in CLI output
- Cleaner batch conversions

#### Professional Word Styles
- Proper built-in styles: Normal, Heading 1-6, ListParagraph
- Character styles: Strong (bold), Emphasis (italic)
- Template compatible output
- Follows Microsoft Word best practices

### Added

- `src/core/parsers/frontmatter-parser.ts` - YAML parsing and validation
- `src/core/validators/document-validator.ts` - Comprehensive validation engine
- `src/core/converters/section-rules.ts` - Section/slide break logic
- `FRONTMATTER.md` - Complete front matter specification
- `examples/frontmatter-template.md` - Ready-to-use template
- `CHANGELOG.md` - This file
- `convert` field in front matter for explicit enable/disable
- `document_type` field for automatic exclusion by category
- `section_breaks` field for DOCX section control
- `slide_breaks` field for PPTX slide control
- Validation output in CLI (errors in red, warnings in yellow)
- Metadata summary display in CLI output
- Skipped file count in conversion summary

### Fixed

- **Critical:** DOCX numbered lists bug - Added proper `default-numbering` configuration
  - Numbered lists were corrupting DOCX files and preventing Word from opening them
  - Now works perfectly with sequential numbering (1, 2, 3...)
  
- **Critical:** XLSX bold text rendering - Fixed rich text formatting
  - Bold text (`**text**`) in markdown tables was not showing as bold in Excel
  - Header cell font override was destroying rich text formatting
  - Data cell conflicting logic was disabling bold
  - Now renders correctly with proper font weight

### Changed

- CLI now validates documents before conversion (can be disabled with `--no-validate`)
- Section breaks in DOCX now intelligent (default: auto mode, only at major boundaries)
- Format detection now uses front matter `format` field when present
- CLI shows detailed validation results for each file
- Exclusion rules apply during batch conversion

### Dependencies

- Added `js-yaml` (^4.1.0) for YAML parsing
- Added `@types/js-yaml` (^4.0.9) for TypeScript support

---

## [1.0.0] - 2025-11-09

### Initial Release

#### Features

- Convert Markdown to DOCX (Word documents)
- Convert Markdown to XLSX (Excel spreadsheets)
- Convert Markdown to PPTX (PowerPoint presentations)
- Excel formula support with `{=FORMULA}` syntax
- 60+ Excel functions supported
- Automatic data type detection (numbers, dates, booleans)
- CLI interface for batch processing
- MCP server interface for AI assistant integration
- Preview and validation commands
- Customizable formatting options
- Australian date format support (DD/MM/YYYY)

#### CLI Commands

- `md-convert <file>` - Convert markdown files
- `md-convert preview <file>` - Preview table extraction
- `md-convert validate <file>` - Validate formulas
- `md-convert serve` - Start MCP server

#### MCP Tools

- `convert_md_to_docx` - Word conversion
- `convert_md_to_xlsx` - Excel conversion with formulas
- `convert_md_to_pptx` - PowerPoint conversion
- `preview_tables` - Table preview
- `validate_formulas` - Formula validation

---

## [Unreleased]

### Planned for 2.2.0

- Image support in all formats
- Custom Word template support
- Multi-level nested list support
- Mermaid diagram rendering
- Table of Contents generation
- Cross-reference support
- Enhanced PPTX layouts
- Git integration (auto-populate metadata from git)
- Batch validation report command

---

## Version History Summary

| Version | Date | Key Features |
|---------|------|--------------|
| 2.1.0 | 2025-11-10 | YAML front matter, validation, intelligent sections, exclusion rules, bug fixes |
| 1.0.0 | 2025-11-09 | Initial release with DOCX, XLSX, PPTX support and MCP integration |

---

## Breaking Changes

### 2.1.0

**None.** All changes are backward compatible.

- Documents without front matter still convert (with warnings)
- Existing CLI commands work unchanged
- Default behavior preserved for all options

**Migration:** Add YAML front matter to get new features. Not required.

---

## Bug Fixes by Version

### 2.1.0

**Critical Bugs Fixed:**

1. **DOCX Numbered Lists Corruption**
   - **Issue:** Numbered lists caused DOCX files to fail opening in Word
   - **Cause:** Missing `default-numbering` configuration in Document creation
   - **Fix:** Added proper numbering scheme configuration
   - **Impact:** All numbered lists now work correctly

2. **XLSX Bold Text Not Rendering**
   - **Issue:** Bold text (`**text**`) in tables didn't show as bold in Excel
   - **Cause:** Font override destroying rich text formatting
   - **Fix:** Conditional font application, let ExcelJS handle rich text
   - **Impact:** Bold and italic text now display correctly in Excel

---

## Credits

**Primary Developer:** Dale Rogers

**Tested on:** DCCEEW AI Capability Enablement project (Australian Government)
- 22 planning documents
- Executive briefings, architecture documents, financial models
- Validated with Microsoft Word 365, Excel 365, PowerPoint 365

---

**Last Updated:** 2025-11-10

