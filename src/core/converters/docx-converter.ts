/**
 * DOCX Converter
 * Convert Markdown to Word document with formatting
 */

// @ts-nocheck
import * as docx from 'docx';
import * as fs from 'fs/promises';
import * as path from 'path';
import { parseMarkdown, type ContentBlock, type TableData } from '../parsers/markdown.js';

export interface DocxConversionOptions {
  pageSize?: 'A4' | 'LETTER';
  margins?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };
  fontSize?: number;
  fontFamily?: string;
}

export interface DocxConversionResult {
  success: boolean;
  outputPath: string;
  pageCount?: number;
  warnings: string[];
}

const HEADING_LEVELS: Record<number, docx.HeadingLevel> = {
  1: docx.HeadingLevel.HEADING_1,
  2: docx.HeadingLevel.HEADING_2,
  3: docx.HeadingLevel.HEADING_3,
  4: docx.HeadingLevel.HEADING_4,
  5: docx.HeadingLevel.HEADING_5,
  6: docx.HeadingLevel.HEADING_6,
};

/**
 * Convert Markdown file to DOCX
 */
export async function convertToDocx(
  inputPath: string,
  outputPath?: string,
  options: DocxConversionOptions = {}
): Promise<DocxConversionResult> {
  // Read markdown file
  const markdownContent = await fs.readFile(inputPath, 'utf-8');
  
  // Parse markdown
  const parsed = parseMarkdown(markdownContent);
  
  // Determine output path
  const output = outputPath || inputPath.replace(/\.md$/, '.docx');
  
  // Create document
  const doc = new docx.Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: (options.margins?.top || 1440) * 20, // Convert to twips (1/20 of a point)
            right: (options.margins?.right || 1440) * 20,
            bottom: (options.margins?.bottom || 1440) * 20,
            left: (options.margins?.left || 1440) * 20,
          },
        },
      },
      children: convertContentToDocx(parsed.content, options),
    }],
  });

  // Ensure output directory exists
  await fs.mkdir(path.dirname(output), { recursive: true });
  
  // Write the document
  const buffer = await docx.Packer.toBuffer(doc);
  await fs.writeFile(output, buffer);

  return {
    success: true,
    outputPath: output,
    warnings: [],
  };
}

/**
 * Convert content blocks to DOCX paragraphs and tables
 */
function convertContentToDocx(
  content: ContentBlock[],
  options: DocxConversionOptions
): Array<docx.Paragraph | docx.Table> {
  const elements: Array<docx.Paragraph | docx.Table> = [];

  for (const block of content) {
    switch (block.type) {
      case 'heading':
        elements.push(createHeading(block, options));
        break;
      
      case 'paragraph':
        elements.push(createParagraph(block, options));
        break;
      
      case 'list':
        elements.push(...createList(block, options));
        break;
      
      case 'table':
        if (typeof block.content === 'object' && 'headers' in block.content) {
          elements.push(createTable(block.content as TableData, options));
        }
        break;
      
      case 'code':
        elements.push(createCodeBlock(block, options));
        break;
      
      case 'hr':
        elements.push(createHorizontalRule());
        break;
    }
  }

  return elements;
}

/**
 * Create a heading paragraph
 */
function createHeading(block: ContentBlock, _options: DocxConversionOptions): docx.Paragraph {
  const level = block.level || 1;
  const text = typeof block.content === 'string' ? block.content : '';
  
  return new docx.Paragraph({
    text,
    heading: HEADING_LEVELS[level] || docx.HeadingLevel.HEADING_1,
    spacing: {
      before: 240,
      after: 120,
    },
  });
}

/**
 * Create a paragraph with inline formatting
 */
function createParagraph(block: ContentBlock, options: DocxConversionOptions): docx.Paragraph {
  const text = typeof block.content === 'string' ? block.content : '';
  const runs = parseInlineFormatting(text, options);
  
  return new docx.Paragraph({
    children: runs,
    spacing: {
      after: 120,
    },
  });
}

/**
 * Parse inline formatting (bold, italic, code)
 */
function parseInlineFormatting(text: string, options: DocxConversionOptions): docx.TextRun[] {
  const runs: docx.TextRun[] = [];
  
  // Simple parsing for bold (**text**), italic (*text*), and inline code (`text`)
  // This is a simplified version - a full implementation would use a proper parser
  
  const segments = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`|\[([^\]]+)\]\(([^)]+)\))/);
  
  for (const segment of segments) {
    if (!segment) continue;
    
    if (segment.startsWith('**') && segment.endsWith('**')) {
      // Bold
      runs.push(new docx.TextRun({
        text: segment.slice(2, -2),
        bold: true,
        size: (options.fontSize || 11) * 2, // Size in half-points
        font: options.fontFamily || 'Calibri',
      }));
    } else if (segment.startsWith('*') && segment.endsWith('*')) {
      // Italic
      runs.push(new docx.TextRun({
        text: segment.slice(1, -1),
        italics: true,
        size: (options.fontSize || 11) * 2,
        font: options.fontFamily || 'Calibri',
      }));
    } else if (segment.startsWith('`') && segment.endsWith('`')) {
      // Inline code
      runs.push(new docx.TextRun({
        text: segment.slice(1, -1),
        font: 'Courier New',
        size: (options.fontSize || 11) * 2,
      }));
    } else {
      // Regular text
      runs.push(new docx.TextRun({
        text: segment,
        size: (options.fontSize || 11) * 2,
        font: options.fontFamily || 'Calibri',
      }));
    }
  }
  
  return runs.length > 0 ? runs : [new docx.TextRun({ 
    text, 
    size: (options.fontSize || 11) * 2,
    font: options.fontFamily || 'Calibri',
  })];
}

/**
 * Create list paragraphs
 */
function createList(block: ContentBlock, options: DocxConversionOptions): docx.Paragraph[] {
  const items = Array.isArray(block.content) ? block.content : [];
  
  return items.map((item) => {
    const runs = parseInlineFormatting(item, options);
    
    return new docx.Paragraph({
      children: runs,
      bullet: block.ordered ? undefined : { level: 0 },
      numbering: block.ordered ? { reference: 'default-numbering', level: 0 } : undefined,
      spacing: {
        after: 60,
      },
    });
  });
}

/**
 * Create a table
 */
function createTable(tableData: TableData, options: DocxConversionOptions): docx.Table {
  const rows: docx.TableRow[] = [];

  // Create header row
  if (tableData.headers.length > 0) {
    const headerCells = tableData.headers.map((header) =>
      new docx.TableCell({
        children: [
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: header,
                bold: true,
                size: (options.fontSize || 11) * 2,
                font: options.fontFamily || 'Calibri',
              }),
            ],
          }),
        ],
        shading: {
          fill: 'D3D3D3', // Light grey background
        },
      })
    );
    
    rows.push(new docx.TableRow({ children: headerCells }));
  }

  // Create data rows
  for (const row of tableData.rows) {
    const cells = row.map((cellValue) => {
      const runs = parseInlineFormatting(cellValue, options);
      
      return new docx.TableCell({
        children: [
          new docx.Paragraph({
            children: runs,
          }),
        ],
      });
    });
    
    rows.push(new docx.TableRow({ children: cells }));
  }

  return new docx.Table({
    rows,
    width: {
      size: 100,
      type: docx.WidthType.PERCENTAGE,
    },
    borders: {
      top: { style: docx.BorderStyle.SINGLE, size: 1 },
      bottom: { style: docx.BorderStyle.SINGLE, size: 1 },
      left: { style: docx.BorderStyle.SINGLE, size: 1 },
      right: { style: docx.BorderStyle.SINGLE, size: 1 },
      insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1 },
      insideVertical: { style: docx.BorderStyle.SINGLE, size: 1 },
    },
  });
}

/**
 * Create a code block
 */
function createCodeBlock(block: ContentBlock, options: DocxConversionOptions): docx.Paragraph {
  const code = typeof block.content === 'string' ? block.content : '';
  
  return new docx.Paragraph({
    children: [
      new docx.TextRun({
        text: code,
        font: 'Courier New',
        size: (options.fontSize || 10) * 2,
      }),
    ],
    shading: {
      fill: 'F5F5F5', // Light grey background
    },
    spacing: {
      before: 120,
      after: 120,
    },
  });
}

/**
 * Create a horizontal rule
 */
function createHorizontalRule(): docx.Paragraph {
  return new docx.Paragraph({
    border: {
      bottom: {
        color: '000000',
        space: 1,
        style: docx.BorderStyle.SINGLE,
        size: 6,
      },
    },
    spacing: {
      before: 120,
      after: 120,
    },
  });
}

/**
 * Convert content directly from markdown string
 */
export async function convertMarkdownToDocx(
  markdown: string,
  outputPath: string,
  options: DocxConversionOptions = {}
): Promise<DocxConversionResult> {
  const parsed = parseMarkdown(markdown);

  const doc = new docx.Document({
    sections: [{
      children: convertContentToDocx(parsed.content, options),
    }],
  });

  const buffer = await docx.Packer.toBuffer(doc);
  await fs.writeFile(outputPath, buffer);

  return {
    success: true,
    outputPath,
    warnings: [],
  };
}

