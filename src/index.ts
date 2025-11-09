/**
 * MD Converter - Main exports
 * Convert Markdown to DOCX, XLSX, and PPTX formats
 */

export { parseMarkdown, extractTables } from './core/parsers/markdown.js';
export { parseFormula, validateFormula } from './core/parsers/formula-parser.js';
export { convertToDocx } from './core/converters/docx-converter.js';
export { convertToXlsx } from './core/converters/xlsx-converter.js';
export { convertToPptx } from './core/converters/pptx-converter.js';

export interface ConversionOptions {
  freezeHeaders?: boolean;
  autoWidth?: boolean;
  theme?: 'light' | 'dark';
  layout?: 'title-content' | 'blank' | 'section';
}

export interface ConversionResult {
  success: boolean;
  outputPath: string;
  message?: string;
  metadata?: Record<string, unknown>;
}

