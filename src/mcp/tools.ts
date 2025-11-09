/**
 * MCP Tools Definitions
 * Define tools exposed by the MCP server
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';

export const TOOLS: Tool[] = [
  {
    name: 'convert_md_to_docx',
    description: 'Convert a Markdown file to DOCX (Word document) format with proper formatting',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the input Markdown file',
        },
        output_path: {
          type: 'string',
          description: 'Optional output path for the DOCX file. If not specified, uses the same name as input with .docx extension',
        },
        options: {
          type: 'object',
          description: 'Optional conversion options',
          properties: {
            fontSize: {
              type: 'number',
              description: 'Font size in points (default: 11)',
            },
            fontFamily: {
              type: 'string',
              description: 'Font family (default: Calibri)',
            },
          },
        },
      },
      required: ['file_path'],
    },
  },
  {
    name: 'convert_md_to_xlsx',
    description: 'Convert Markdown tables to XLSX (Excel) format with formula support. Formulas in tables using {=FORMULA} syntax will be converted to actual Excel formulas',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the input Markdown file',
        },
        output_path: {
          type: 'string',
          description: 'Optional output path for the XLSX file. If not specified, uses the same name as input with .xlsx extension',
        },
        options: {
          type: 'object',
          description: 'Optional conversion options',
          properties: {
            freezeHeaders: {
              type: 'boolean',
              description: 'Freeze the header row (default: true)',
            },
            autoWidth: {
              type: 'boolean',
              description: 'Auto-size columns to fit content (default: true)',
            },
          },
        },
      },
      required: ['file_path'],
    },
  },
  {
    name: 'convert_md_to_pptx',
    description: 'Convert Markdown to PPTX (PowerPoint) presentation. H1=Title slide, H2=Section header, H3=Content slide. Use --- (horizontal rule) to separate slides',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the input Markdown file',
        },
        output_path: {
          type: 'string',
          description: 'Optional output path for the PPTX file. If not specified, uses the same name as input with .pptx extension',
        },
        options: {
          type: 'object',
          description: 'Optional conversion options',
          properties: {
            theme: {
              type: 'string',
              enum: ['light', 'dark'],
              description: 'Presentation theme (default: light)',
            },
            fontSize: {
              type: 'number',
              description: 'Font size for content in points (default: 18)',
            },
          },
        },
      },
      required: ['file_path'],
    },
  },
  {
    name: 'preview_tables',
    description: 'Preview table extraction from a Markdown file, showing headers, row/column counts, and formula locations',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Markdown file',
        },
      },
      required: ['file_path'],
    },
  },
  {
    name: 'validate_formulas',
    description: 'Validate all formulas in a Markdown file, checking syntax and Excel function names',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Markdown file',
        },
      },
      required: ['file_path'],
    },
  },
];

