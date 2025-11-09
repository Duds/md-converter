/**
 * XLSX Converter
 * Convert Markdown tables to Excel with formula support
 */

import ExcelJS from 'exceljs';
import { parseMarkdown, type TableData } from '../parsers/markdown.js';
import { processTable, type ProcessedTable } from '../parsers/table-parser.js';
import { validateFormula } from '../parsers/formula-parser.js';
import * as fs from 'fs/promises';
import * as path from 'path';

export interface XlsxConversionOptions {
  freezeHeaders?: boolean;
  autoWidth?: boolean;
  addBorders?: boolean;
  headerStyle?: {
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
  };
}

export interface XlsxConversionResult {
  success: boolean;
  outputPath: string;
  worksheetNames: string[];
  tableCount: number;
  formulaCount: number;
  warnings: string[];
}

/**
 * Convert Markdown file to XLSX with formula support
 */
export async function convertToXlsx(
  inputPath: string,
  outputPath?: string,
  options: XlsxConversionOptions = {}
): Promise<XlsxConversionResult> {
  // Read markdown file
  const markdownContent = await fs.readFile(inputPath, 'utf-8');
  
  // Parse markdown
  const parsed = parseMarkdown(markdownContent);
  
  if (parsed.tables.length === 0) {
    throw new Error('No tables found in the markdown file');
  }

  // Determine output path
  const output = outputPath || inputPath.replace(/\.md$/, '.xlsx');
  
  // Create workbook
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MD Converter';
  workbook.created = new Date();
  
  const worksheetNames: string[] = [];
  let totalFormulas = 0;
  const warnings: string[] = [];

  // Process each table
  for (let i = 0; i < parsed.tables.length; i++) {
    const table = parsed.tables[i];
    const processedTable = processTable(table);
    
    // Create worksheet
    const sheetName = `Table ${i + 1}`;
    const worksheet = workbook.addWorksheet(sheetName);
    worksheetNames.push(sheetName);

    // Add table to worksheet
    const formulaCount = await addTableToWorksheet(
      worksheet,
      processedTable,
      options,
      warnings
    );
    
    totalFormulas += formulaCount;
  }

  // Ensure output directory exists
  await fs.mkdir(path.dirname(output), { recursive: true });
  
  // Save workbook
  await workbook.xlsx.writeFile(output);

  return {
    success: true,
    outputPath: output,
    worksheetNames,
    tableCount: parsed.tables.length,
    formulaCount: totalFormulas,
    warnings,
  };
}

/**
 * Add a processed table to a worksheet
 */
async function addTableToWorksheet(
  worksheet: ExcelJS.Worksheet,
  table: ProcessedTable,
  options: XlsxConversionOptions,
  warnings: string[]
): Promise<number> {
  let formulaCount = 0;
  const startRow = 1;
  const startCol = 1;

  // Add headers
  if (table.headers.length > 0) {
    for (let col = 0; col < table.headers.length; col++) {
      const cell = worksheet.getCell(startRow, startCol + col);
      cell.value = table.headers[col];
      
      // Apply header styling
      cell.font = { 
        bold: options.headerStyle?.bold !== false,
        size: options.headerStyle?.fontSize || 12,
      };
      
      if (options.headerStyle?.fillColor) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: options.headerStyle.fillColor },
        };
      } else {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' }, // Light grey
        };
      }
      
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
    }

    // Freeze header row if requested
    if (options.freezeHeaders !== false) {
      worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    }
  }

  // Add data rows
  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
    const row = table.rows[rowIndex];
    const excelRow = startRow + rowIndex + 1; // +1 for header row

    for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
      const cellData = row.cells[colIndex];
      const cell = worksheet.getCell(excelRow, startCol + colIndex);

      if (cellData.isFormula && cellData.formula) {
        // Validate formula
        const validation = validateFormula(cellData.formula);
        
        if (!validation.isValid) {
          warnings.push(
            `Invalid formula in row ${excelRow}, col ${startCol + colIndex}: ${validation.errors.join(', ')}`
          );
          // Fall back to text
          cell.value = cellData.rawValue;
        } else {
          // Set as formula (ExcelJS expects formulas without leading =)
          const formulaWithoutEquals = cellData.formula.startsWith('=') 
            ? cellData.formula.substring(1) 
            : cellData.formula;
          
          cell.value = { formula: formulaWithoutEquals };
          formulaCount++;
          
          // Add warnings if any
          if (validation.warnings.length > 0) {
            warnings.push(
              `Formula warnings in row ${excelRow}, col ${startCol + colIndex}: ${validation.warnings.join(', ')}`
            );
          }
        }
      } else if (cellData.dataType === 'number' && cellData.numericValue !== undefined) {
        // Set as number
        cell.value = cellData.numericValue;
        cell.numFmt = '#,##0.00';
      } else if (cellData.dataType === 'date' && cellData.dateValue) {
        // Set as date (Australian format DD/MM/YYYY)
        cell.value = cellData.dateValue;
        cell.numFmt = 'dd/mm/yyyy';
      } else if (cellData.dataType === 'boolean') {
        // Set as boolean
        cell.value = cellData.displayValue.toLowerCase() === 'true';
      } else {
        // Set as text
        cell.value = cellData.displayValue;
      }

      // Apply basic alignment
      if (cellData.dataType === 'number' || cellData.dataType === 'formula') {
        cell.alignment = { horizontal: 'right' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
    }
  }

  // Add borders if requested
  if (options.addBorders !== false) {
    const borderStyle: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF000000' } };
    
    for (let row = startRow; row <= startRow + table.rows.length; row++) {
      for (let col = startCol; col < startCol + table.headers.length; col++) {
        const cell = worksheet.getCell(row, col);
        cell.border = {
          top: borderStyle,
          left: borderStyle,
          bottom: borderStyle,
          right: borderStyle,
        };
      }
    }
  }

  // Auto-size columns if requested
  if (options.autoWidth !== false) {
    worksheet.columns.forEach((column) => {
      if (column) {
        let maxLength = 10;
        
        // Check each cell in the column
        column.eachCell?.({ includeEmpty: false }, (cell) => {
          const cellValue = cell.value?.toString() || '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
        
        // Set width with some padding (max 50 characters)
        column.width = Math.min(maxLength + 2, 50);
      }
    });
  }

  return formulaCount;
}

/**
 * Convert a single table to XLSX
 */
export async function convertTableToXlsx(
  table: TableData,
  outputPath: string,
  options: XlsxConversionOptions = {}
): Promise<XlsxConversionResult> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MD Converter';
  workbook.created = new Date();

  const processedTable = processTable(table);
  const worksheet = workbook.addWorksheet('Data');
  const warnings: string[] = [];

  const formulaCount = await addTableToWorksheet(
    worksheet,
    processedTable,
    options,
    warnings
  );

  await workbook.xlsx.writeFile(outputPath);

  return {
    success: true,
    outputPath,
    worksheetNames: ['Data'],
    tableCount: 1,
    formulaCount,
    warnings,
  };
}

/**
 * Preview how a table will be converted (without saving)
 */
export function previewTableConversion(table: TableData): {
  headers: string[];
  rowCount: number;
  columnCount: number;
  formulas: Array<{ row: number; col: number; formula: string }>;
  dataTypes: Record<string, number>;
} {
  const processedTable = processTable(table);
  const formulas: Array<{ row: number; col: number; formula: string }> = [];
  const dataTypes: Record<string, number> = {
    string: 0,
    number: 0,
    boolean: 0,
    date: 0,
    formula: 0,
  };

  for (let rowIndex = 0; rowIndex < processedTable.rows.length; rowIndex++) {
    const row = processedTable.rows[rowIndex];
    
    for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
      const cell = row.cells[colIndex];
      
      dataTypes[cell.dataType]++;
      
      if (cell.isFormula && cell.formula) {
        formulas.push({
          row: rowIndex + 2, // +2 for 1-indexed and header row
          col: colIndex + 1, // +1 for 1-indexed
          formula: cell.formula,
        });
      }
    }
  }

  return {
    headers: processedTable.headers,
    rowCount: processedTable.rows.length,
    columnCount: processedTable.headers.length,
    formulas,
    dataTypes,
  };
}

