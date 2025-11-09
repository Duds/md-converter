#!/usr/bin/env node

/**
 * MD Converter CLI
 * Command-line interface for converting Markdown to DOCX, XLSX, and PPTX
 */

import { Command } from 'commander';
import chalk from 'chalk';
import { glob } from 'glob';
import * as path from 'path';
import * as fs from 'fs/promises';
import { convertToDocx } from '../core/converters/docx-converter.js';
import { convertToXlsx } from '../core/converters/xlsx-converter.js';
import { convertToPptx } from '../core/converters/pptx-converter.js';
import { parseMarkdown } from '../core/parsers/markdown.js';
import { processTable, extractFormulas } from '../core/parsers/table-parser.js';
import { validateFormula } from '../core/parsers/formula-parser.js';

const program = new Command();

program
  .name('md-convert')
  .description('Convert Markdown files to DOCX, XLSX, and PPTX formats with Excel formula support')
  .version('1.0.0');

// Convert command
program
  .argument('<input>', 'Input markdown file or glob pattern')
  .option('-f, --format <format>', 'Output format: docx, xlsx, pptx, or all', 'all')
  .option('-o, --output <path>', 'Output file path (for single file)')
  .option('-d, --output-dir <dir>', 'Output directory (for batch processing)')
  .option('--freeze-headers', 'Freeze header row in Excel (XLSX only)', true)
  .option('--auto-width', 'Auto-size columns in Excel (XLSX only)', true)
  .option('--theme <theme>', 'Presentation theme: light or dark (PPTX only)', 'light')
  .option('--font-family <family>', 'Font family to use', 'Arial')
  .option('--font-size <size>', 'Font size in points', '12')
  .action(async (input, options) => {
    try {
      await handleConvert(input, options);
    } catch (error) {
      console.error(chalk.red('Error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  });

// Preview command
program
  .command('preview')
  .description('Preview table extraction and formulas from a markdown file')
  .argument('<input>', 'Input markdown file')
  .action(async (input) => {
    try {
      await handlePreview(input);
    } catch (error) {
      console.error(chalk.red('Error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  });

// Validate command
program
  .command('validate')
  .description('Validate formulas in a markdown file')
  .argument('<input>', 'Input markdown file')
  .action(async (input) => {
    try {
      await handleValidate(input);
    } catch (error) {
      console.error(chalk.red('Error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  });

// Serve command (MCP server)
program
  .command('serve')
  .description('Start MCP server (STDIO mode)')
  .action(async () => {
    try {
      // Dynamic import to avoid loading MCP server code unless needed
      const { startMcpServer } = await import('../mcp/server.js');
      await startMcpServer();
    } catch (error) {
      console.error(chalk.red('Error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  });

program.parse();

/**
 * Handle convert command
 */
async function handleConvert(input: string, options: any): Promise<void> {
  const format = options.format.toLowerCase();
  const formats = format === 'all' ? ['docx', 'xlsx', 'pptx'] : [format];

  // Validate format
  for (const fmt of formats) {
    if (!['docx', 'xlsx', 'pptx'].includes(fmt)) {
      throw new Error(`Invalid format: ${fmt}. Must be docx, xlsx, pptx, or all`);
    }
  }

  // Find input files
  const files = await findInputFiles(input);
  
  if (files.length === 0) {
    throw new Error(`No markdown files found matching: ${input}`);
  }

  console.log(chalk.blue(`Found ${files.length} file(s) to convert`));
  console.log();

  let successCount = 0;
  let errorCount = 0;

  // Convert each file
  for (const file of files) {
    console.log(chalk.cyan(`Converting: ${file}`));
    
    try {
      for (const fmt of formats) {
        const outputPath = await getOutputPath(file, fmt, options);
        
        await convertFile(file, outputPath, fmt, options);
        
        console.log(chalk.green(`  ✓ ${fmt.toUpperCase()}: ${outputPath}`));
        successCount++;
      }
    } catch (error) {
      console.error(chalk.red(`  ✗ Error: ${error instanceof Error ? error.message : String(error)}`));
      errorCount++;
    }
    
    console.log();
  }

  // Summary
  console.log(chalk.bold('Conversion Summary:'));
  console.log(chalk.green(`  ✓ Success: ${successCount}`));
  if (errorCount > 0) {
    console.log(chalk.red(`  ✗ Failed: ${errorCount}`));
  }
}

/**
 * Handle preview command
 */
async function handlePreview(input: string): Promise<void> {
  const content = await fs.readFile(input, 'utf-8');
  const parsed = parseMarkdown(content);

  console.log(chalk.bold.blue('Markdown Preview'));
  console.log(chalk.grey('─'.repeat(50)));
  console.log();

  console.log(chalk.yellow(`Tables found: ${parsed.tables.length}`));
  console.log();

  for (let i = 0; i < parsed.tables.length; i++) {
    const table = parsed.tables[i];
    const processed = processTable(table);
    const formulas = extractFormulas(processed);

    console.log(chalk.cyan(`Table ${i + 1}:`));
    console.log(chalk.grey(`  Headers: ${table.headers.join(', ')}`));
    console.log(chalk.grey(`  Rows: ${table.rows.length}`));
    console.log(chalk.grey(`  Columns: ${table.headers.length}`));
    
    if (formulas.length > 0) {
      console.log(chalk.green(`  Formulas: ${formulas.length}`));
      formulas.forEach((formula, idx) => {
        console.log(chalk.grey(`    ${idx + 1}. ${formula}`));
      });
    } else {
      console.log(chalk.grey(`  Formulas: None`));
    }
    
    console.log();
  }
}

/**
 * Handle validate command
 */
async function handleValidate(input: string): Promise<void> {
  const content = await fs.readFile(input, 'utf-8');
  const parsed = parseMarkdown(content);

  console.log(chalk.bold.blue('Formula Validation'));
  console.log(chalk.grey('─'.repeat(50)));
  console.log();

  let totalFormulas = 0;
  let validFormulas = 0;
  let invalidFormulas = 0;

  for (let i = 0; i < parsed.tables.length; i++) {
    const table = parsed.tables[i];
    const processed = processTable(table);
    const formulas = extractFormulas(processed);

    if (formulas.length > 0) {
      console.log(chalk.cyan(`Table ${i + 1}:`));
      
      for (const formula of formulas) {
        totalFormulas++;
        const validation = validateFormula(formula);

        if (validation.isValid) {
          validFormulas++;
          console.log(chalk.green(`  ✓ ${formula}`));
          
          if (validation.warnings.length > 0) {
            validation.warnings.forEach((warning) => {
              console.log(chalk.yellow(`    ⚠ ${warning}`));
            });
          }
        } else {
          invalidFormulas++;
          console.log(chalk.red(`  ✗ ${formula}`));
          
          validation.errors.forEach((error) => {
            console.log(chalk.red(`    ✗ ${error}`));
          });
        }
      }
      
      console.log();
    }
  }

  console.log(chalk.bold('Validation Summary:'));
  console.log(chalk.grey(`  Total formulas: ${totalFormulas}`));
  console.log(chalk.green(`  Valid: ${validFormulas}`));
  if (invalidFormulas > 0) {
    console.log(chalk.red(`  Invalid: ${invalidFormulas}`));
  }
}

/**
 * Find input files matching pattern
 */
async function findInputFiles(pattern: string): Promise<string[]> {
  // If it's a single file, return it
  try {
    const stats = await fs.stat(pattern);
    if (stats.isFile()) {
      return [pattern];
    }
  } catch {
    // Not a single file, try glob pattern
  }

  // Use glob to find files
  const files = await glob(pattern, {
    ignore: ['node_modules/**', 'dist/**', '.git/**'],
  });

  return files;
}

/**
 * Get output path for a file
 */
async function getOutputPath(
  inputPath: string,
  format: string,
  options: any
): Promise<string> {
  if (options.output && !options.outputDir) {
    // Single file output specified
    return options.output;
  }

  const dir = options.outputDir || path.dirname(inputPath);
  const basename = path.basename(inputPath, path.extname(inputPath));
  const outputFilename = `${basename}.${format}`;

  return path.join(dir, outputFilename);
}

/**
 * Convert a single file
 */
async function convertFile(
  inputPath: string,
  outputPath: string,
  format: string,
  options: any
): Promise<void> {
  const conversionOptions = {
    freezeHeaders: options.freezeHeaders,
    autoWidth: options.autoWidth,
    theme: options.theme,
    fontFamily: options.fontFamily,
    fontSize: parseInt(options.fontSize, 10),
  };

  switch (format) {
    case 'docx':
      await convertToDocx(inputPath, outputPath, conversionOptions);
      break;
    
    case 'xlsx':
      await convertToXlsx(inputPath, outputPath, conversionOptions);
      break;
    
    case 'pptx':
      await convertToPptx(inputPath, outputPath, conversionOptions);
      break;
    
    default:
      throw new Error(`Unsupported format: ${format}`);
  }
}

