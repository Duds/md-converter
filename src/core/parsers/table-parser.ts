/**
 * Table Parser
 * Extract and process tables with formula detection
 */

import type { TableData } from './markdown.js';

export interface ProcessedTable {
  headers: string[];
  rows: ProcessedRow[];
  hasFormulas: boolean;
}

export interface ProcessedRow {
  cells: ProcessedCell[];
}

export interface ProcessedCell {
  rawValue: string;
  displayValue: string;
  isFormula: boolean;
  formula?: string;
  dataType: 'string' | 'number' | 'boolean' | 'date' | 'formula';
  numericValue?: number;
  dateValue?: Date;
}

const FORMULA_PATTERN = /\{=([^}]+)\}/g;

/**
 * Process a table and detect formulas
 */
export function processTable(table: TableData): ProcessedTable {
  const processedRows: ProcessedRow[] = [];
  let hasFormulas = false;

  for (const row of table.rows) {
    const processedCells: ProcessedCell[] = [];
    
    for (const cellValue of row) {
      const processed = processCell(cellValue);
      if (processed.isFormula) {
        hasFormulas = true;
      }
      processedCells.push(processed);
    }
    
    processedRows.push({ cells: processedCells });
  }

  return {
    headers: table.headers,
    rows: processedRows,
    hasFormulas,
  };
}

/**
 * Process a single cell and detect its type
 */
export function processCell(value: string): ProcessedCell {
  const trimmedValue = value.trim();

  // Check for formula
  const formulaMatch = trimmedValue.match(FORMULA_PATTERN);
  if (formulaMatch) {
    const formula = formulaMatch[0].substring(2, formulaMatch[0].length - 1); // Remove {= and }
    return {
      rawValue: value,
      displayValue: formula,
      isFormula: true,
      formula,
      dataType: 'formula',
    };
  }

  // Check for boolean
  if (trimmedValue.toLowerCase() === 'true' || trimmedValue.toLowerCase() === 'false') {
    return {
      rawValue: value,
      displayValue: trimmedValue,
      isFormula: false,
      dataType: 'boolean',
    };
  }

  // Check for number (including currency)
  const numericValue = parseNumeric(trimmedValue);
  if (numericValue !== null) {
    return {
      rawValue: value,
      displayValue: trimmedValue,
      isFormula: false,
      dataType: 'number',
      numericValue,
    };
  }

  // Check for date (DD/MM/YYYY format - Australian)
  const dateValue = parseAustralianDate(trimmedValue);
  if (dateValue) {
    return {
      rawValue: value,
      displayValue: trimmedValue,
      isFormula: false,
      dataType: 'date',
      dateValue,
    };
  }

  // Default to string
  return {
    rawValue: value,
    displayValue: trimmedValue,
    isFormula: false,
    dataType: 'string',
  };
}

/**
 * Parse numeric values including currency symbols
 */
function parseNumeric(value: string): number | null {
  // Remove common currency symbols and formatting
  const cleaned = value
    .replace(/[$€£¥AUD]/g, '')
    .replace(/,/g, '')
    .replace(/\s/g, '')
    .trim();

  if (cleaned === '') {
    return null;
  }

  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? null : parsed;
}

/**
 * Parse Australian date format (DD/MM/YYYY)
 */
function parseAustralianDate(value: string): Date | null {
  const datePattern = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
  const match = value.match(datePattern);

  if (!match) {
    return null;
  }

  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1; // JS months are 0-indexed
  const year = parseInt(match[3], 10);

  const date = new Date(year, month, day);

  // Validate the date is valid
  if (
    date.getFullYear() === year &&
    date.getMonth() === month &&
    date.getDate() === day
  ) {
    return date;
  }

  return null;
}

/**
 * Get all formulas from a table
 */
export function extractFormulas(table: ProcessedTable): string[] {
  const formulas: string[] = [];

  for (const row of table.rows) {
    for (const cell of row.cells) {
      if (cell.isFormula && cell.formula) {
        formulas.push(cell.formula);
      }
    }
  }

  return formulas;
}

/**
 * Get table dimensions for Excel cell references
 */
export interface TableDimensions {
  rows: number;
  cols: number;
  hasHeaders: boolean;
}

export function getTableDimensions(table: ProcessedTable): TableDimensions {
  return {
    rows: table.rows.length,
    cols: table.headers.length,
    hasHeaders: table.headers.length > 0,
  };
}

/**
 * Convert column index to Excel column letter (0 -> A, 1 -> B, 25 -> Z, 26 -> AA)
 */
export function columnIndexToLetter(index: number): string {
  let letter = '';
  let num = index;

  while (num >= 0) {
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26) - 1;
  }

  return letter;
}

/**
 * Convert Excel column letter to index (A -> 0, B -> 1, AA -> 26)
 */
export function columnLetterToIndex(letter: string): number {
  let index = 0;
  
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  
  return index - 1;
}

