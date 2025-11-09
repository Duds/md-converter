/**
 * Formula Parser
 * Parse and validate Excel formulas
 */

export interface FormulaValidation {
  isValid: boolean;
  formula: string;
  errors: string[];
  warnings: string[];
  cellReferences: string[];
  functions: string[];
}

// Common Excel functions whitelist
const EXCEL_FUNCTIONS = [
  // Math functions
  'SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS',
  'MIN', 'MAX', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'ABS',
  'SQRT', 'POWER', 'MOD', 'PRODUCT', 'SUMIF', 'SUMIFS',
  'AVERAGEIF', 'AVERAGEIFS', 'MEDIAN', 'MODE',
  
  // Logical functions
  'IF', 'AND', 'OR', 'NOT', 'XOR', 'TRUE', 'FALSE', 'IFERROR', 'IFNA',
  
  // Text functions
  'CONCATENATE', 'CONCAT', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM',
  'UPPER', 'LOWER', 'PROPER', 'SUBSTITUTE', 'REPLACE', 'TEXT',
  
  // Date functions
  'TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY',
  'DATEDIF', 'DAYS', 'NETWORKDAYS', 'EOMONTH',
  
  // Lookup functions
  'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'CHOOSE',
  
  // Statistical functions
  'STDEV', 'STDEVP', 'VAR', 'VARP', 'RANK', 'PERCENTILE',
  
  // Financial functions
  'PMT', 'FV', 'PV', 'RATE', 'NPV', 'IRR',
];

const CELL_REFERENCE_PATTERN = /\$?[A-Z]+\$?\d+/g;
const RANGE_PATTERN = /\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+/g;
const COLUMN_RANGE_PATTERN = /\$?[A-Z]+:\$?[A-Z]+/g;
const ROW_RANGE_PATTERN = /\$?\d+:\$?\d+/g;
const FUNCTION_PATTERN = /\b([A-Z]+)\s*\(/g;

/**
 * Parse and extract information from an Excel formula
 */
export function parseFormula(formula: string): FormulaValidation {
  const errors: string[] = [];
  const warnings: string[] = [];
  const cellReferences: string[] = [];
  const functions: string[] = [];

  // Remove leading = if present
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

  // Extract cell references
  const cellMatches = cleanFormula.match(CELL_REFERENCE_PATTERN);
  if (cellMatches) {
    cellReferences.push(...cellMatches);
  }

  // Extract range references
  const rangeMatches = cleanFormula.match(RANGE_PATTERN);
  if (rangeMatches) {
    cellReferences.push(...rangeMatches);
  }

  // Extract column ranges (e.g., A:A)
  const columnMatches = cleanFormula.match(COLUMN_RANGE_PATTERN);
  if (columnMatches) {
    cellReferences.push(...columnMatches);
  }

  // Extract row ranges (e.g., 1:10)
  const rowMatches = cleanFormula.match(ROW_RANGE_PATTERN);
  if (rowMatches) {
    cellReferences.push(...rowMatches);
  }

  // Extract and validate functions
  const functionMatches = cleanFormula.matchAll(FUNCTION_PATTERN);
  for (const match of functionMatches) {
    const funcName = match[1];
    functions.push(funcName);

    if (!EXCEL_FUNCTIONS.includes(funcName)) {
      warnings.push(`Unknown function: ${funcName}`);
    }
  }

  // Basic syntax validation
  const openParens = (cleanFormula.match(/\(/g) || []).length;
  const closeParens = (cleanFormula.match(/\)/g) || []).length;
  if (openParens !== closeParens) {
    errors.push('Mismatched parentheses');
  }

  // Check for empty formula
  if (cleanFormula.trim() === '') {
    errors.push('Empty formula');
  }

  return {
    isValid: errors.length === 0,
    formula: cleanFormula,
    errors,
    warnings,
    cellReferences,
    functions,
  };
}

/**
 * Validate a formula and return detailed results
 */
export function validateFormula(formula: string): FormulaValidation {
  return parseFormula(formula);
}

/**
 * Check if a string contains a formula pattern
 */
export function isFormula(value: string): boolean {
  const pattern = /\{=([^}]+)\}/;
  return pattern.test(value);
}

/**
 * Extract formula from markdown syntax {=FORMULA}
 */
export function extractFormula(value: string): string | null {
  const pattern = /\{=([^}]+)\}/;
  const match = value.match(pattern);
  return match ? match[1] : null;
}

/**
 * Validate cell reference format
 */
export function isValidCellReference(ref: string): boolean {
  const pattern = /^\$?[A-Z]+\$?\d+$/;
  return pattern.test(ref);
}

/**
 * Validate range reference format
 */
export function isValidRangeReference(ref: string): boolean {
  const cellRange = /^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/;
  const columnRange = /^\$?[A-Z]+:\$?[A-Z]+$/;
  const rowRange = /^\$?\d+:\$?\d+$/;
  
  return cellRange.test(ref) || columnRange.test(ref) || rowRange.test(ref);
}

/**
 * Parse cell reference into components
 */
export interface CellReference {
  column: string;
  row: number;
  absoluteColumn: boolean;
  absoluteRow: boolean;
}

export function parseCellReference(ref: string): CellReference | null {
  const pattern = /^(\$)?([A-Z]+)(\$)?(\d+)$/;
  const match = ref.match(pattern);

  if (!match) {
    return null;
  }

  return {
    column: match[2],
    row: parseInt(match[4], 10),
    absoluteColumn: match[1] === '$',
    absoluteRow: match[3] === '$',
  };
}

/**
 * Format cell reference with optional absolute markers
 */
export function formatCellReference(
  column: string,
  row: number,
  absoluteColumn = false,
  absoluteRow = false
): string {
  const colPrefix = absoluteColumn ? '$' : '';
  const rowPrefix = absoluteRow ? '$' : '';
  return `${colPrefix}${column}${rowPrefix}${row}`;
}

/**
 * Get all unique cell references from a formula
 */
export function getAllReferences(formula: string): string[] {
  const validation = parseFormula(formula);
  return Array.from(new Set(validation.cellReferences));
}

/**
 * Check if formula contains circular references (basic check)
 */
export function hasCircularReference(
  formula: string,
  currentCell: string
): boolean {
  const references = getAllReferences(formula);
  return references.includes(currentCell);
}

/**
 * Sanitise formula for safe execution
 */
export function sanitiseFormula(formula: string): string {
  // Remove any potentially dangerous characters
  let cleaned = formula.trim();
  
  // Remove leading = if present
  if (cleaned.startsWith('=')) {
    cleaned = cleaned.substring(1);
  }
  
  // Remove any script tags or HTML
  cleaned = cleaned.replace(/<[^>]*>/g, '');
  
  return cleaned;
}

/**
 * Convert relative references to absolute based on offset
 */
export function adjustReferences(
  formula: string,
  rowOffset: number,
  colOffset: number
): string {
  let adjusted = formula;

  // Find all cell references
  const references = formula.match(CELL_REFERENCE_PATTERN);
  if (!references) {
    return formula;
  }

  for (const ref of references) {
    const parsed = parseCellReference(ref);
    if (!parsed) continue;

    // Only adjust relative references
    let newCol = parsed.column;
    let newRow = parsed.row;

    if (!parsed.absoluteColumn && colOffset !== 0) {
      // Adjust column (convert to index, add offset, convert back)
      const colIndex = columnToIndex(parsed.column);
      newCol = indexToColumn(colIndex + colOffset);
    }

    if (!parsed.absoluteRow && rowOffset !== 0) {
      newRow = parsed.row + rowOffset;
    }

    const newRef = formatCellReference(
      newCol,
      newRow,
      parsed.absoluteColumn,
      parsed.absoluteRow
    );

    adjusted = adjusted.replace(ref, newRef);
  }

  return adjusted;
}

/**
 * Convert column letter to index (A=0, B=1, Z=25, AA=26)
 */
function columnToIndex(column: string): number {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Convert index to column letter (0=A, 1=B, 25=Z, 26=AA)
 */
function indexToColumn(index: number): string {
  let column = '';
  let num = index + 1;

  while (num > 0) {
    const remainder = (num - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    num = Math.floor((num - 1) / 26);
  }

  return column;
}

