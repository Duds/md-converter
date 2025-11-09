/**
 * MCP Server
 * Model Context Protocol server for MD Converter
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import * as fs from 'fs/promises';
import * as path from 'path';
import { convertToDocx } from '../core/converters/docx-converter.js';
import { convertToXlsx, previewTableConversion } from '../core/converters/xlsx-converter.js';
import { convertToPptx } from '../core/converters/pptx-converter.js';
import { parseMarkdown } from '../core/parsers/markdown.js';
import { processTable, extractFormulas } from '../core/parsers/table-parser.js';
import { validateFormula } from '../core/parsers/formula-parser.js';
import { TOOLS } from './tools.js';

/**
 * Start the MCP server
 */
export async function startMcpServer(): Promise<void> {
  const server = new Server(
    {
      name: 'md-converter',
      version: '1.0.0',
    },
    {
      capabilities: {
        tools: {},
        resources: {},
      },
    }
  );

  // List available tools
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: TOOLS,
    };
  });

  // Handle tool calls
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      switch (name) {
        case 'convert_md_to_docx':
          return await handleConvertToDocx(args);
        
        case 'convert_md_to_xlsx':
          return await handleConvertToXlsx(args);
        
        case 'convert_md_to_pptx':
          return await handleConvertToPptx(args);
        
        case 'preview_tables':
          return await handlePreviewTables(args);
        
        case 'validate_formulas':
          return await handleValidateFormulas(args);
        
        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error: ${error instanceof Error ? error.message : String(error)}`,
          },
        ],
        isError: true,
      };
    }
  });

  // List available resources
  server.setRequestHandler(ListResourcesRequestSchema, async () => {
    const examplesDir = path.join(process.cwd(), 'examples');
    
    try {
      const files = await fs.readdir(examplesDir);
      const mdFiles = files.filter(f => f.endsWith('.md'));
      
      return {
        resources: mdFiles.map(file => ({
          uri: `example://${file}`,
          name: file,
          description: `Example Markdown file: ${file}`,
          mimeType: 'text/markdown',
        })),
      };
    } catch {
      return { resources: [] };
    }
  });

  // Read resources
  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    const { uri } = request.params;
    
    if (uri.startsWith('example://')) {
      const filename = uri.replace('example://', '');
      const filepath = path.join(process.cwd(), 'examples', filename);
      
      try {
        const content = await fs.readFile(filepath, 'utf-8');
        return {
          contents: [
            {
              uri,
              mimeType: 'text/markdown',
              text: content,
            },
          ],
        };
      } catch (error) {
        throw new Error(`Failed to read resource: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
    
    throw new Error(`Unknown resource URI: ${uri}`);
  });

  // Start server with stdio transport
  const transport = new StdioServerTransport();
  await server.connect(transport);
  
  console.error('MD Converter MCP server running on stdio');
}

/**
 * Handle convert to DOCX tool
 */
async function handleConvertToDocx(args: any) {
  const { file_path, output_path, options = {} } = args;
  
  const result = await convertToDocx(file_path, output_path, options);
  
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          success: result.success,
          outputPath: result.outputPath,
          warnings: result.warnings,
        }, null, 2),
      },
    ],
  };
}

/**
 * Handle convert to XLSX tool
 */
async function handleConvertToXlsx(args: any) {
  const { file_path, output_path, options = {} } = args;
  
  const result = await convertToXlsx(file_path, output_path, options);
  
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          success: result.success,
          outputPath: result.outputPath,
          worksheetNames: result.worksheetNames,
          tableCount: result.tableCount,
          formulaCount: result.formulaCount,
          warnings: result.warnings,
        }, null, 2),
      },
    ],
  };
}

/**
 * Handle convert to PPTX tool
 */
async function handleConvertToPptx(args: any) {
  const { file_path, output_path, options = {} } = args;
  
  const result = await convertToPptx(file_path, output_path, options);
  
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          success: result.success,
          outputPath: result.outputPath,
          slideCount: result.slideCount,
          warnings: result.warnings,
        }, null, 2),
      },
    ],
  };
}

/**
 * Handle preview tables tool
 */
async function handlePreviewTables(args: any) {
  const { file_path } = args;
  
  const content = await fs.readFile(file_path, 'utf-8');
  const parsed = parseMarkdown(content);
  
  const previews = parsed.tables.map((table, index) => {
    const preview = previewTableConversion(table);
    return {
      tableNumber: index + 1,
      ...preview,
    };
  });
  
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          tableCount: parsed.tables.length,
          tables: previews,
        }, null, 2),
      },
    ],
  };
}

/**
 * Handle validate formulas tool
 */
async function handleValidateFormulas(args: any) {
  const { file_path } = args;
  
  const content = await fs.readFile(file_path, 'utf-8');
  const parsed = parseMarkdown(content);
  
  const validations: any[] = [];
  let totalFormulas = 0;
  let validCount = 0;
  let invalidCount = 0;
  
  for (let i = 0; i < parsed.tables.length; i++) {
    const table = parsed.tables[i];
    const processed = processTable(table);
    const formulas = extractFormulas(processed);
    
    for (const formula of formulas) {
      totalFormulas++;
      const validation = validateFormula(formula);
      
      if (validation.isValid) {
        validCount++;
      } else {
        invalidCount++;
      }
      
      validations.push({
        tableNumber: i + 1,
        formula,
        isValid: validation.isValid,
        errors: validation.errors,
        warnings: validation.warnings,
        cellReferences: validation.cellReferences,
        functions: validation.functions,
      });
    }
  }
  
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          summary: {
            totalFormulas,
            validCount,
            invalidCount,
          },
          validations,
        }, null, 2),
      },
    ],
  };
}

