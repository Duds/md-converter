/**
 * Markdown Parser
 * Parse MD to AST and extract structured content
 */

import MarkdownIt from 'markdown-it';
import type Token from 'markdown-it/lib/token.js';

export interface ParsedMarkdown {
  tokens: Token[];
  tables: TableData[];
  headings: HeadingData[];
  content: ContentBlock[];
}

export interface TableData {
  headers: string[];
  rows: string[][];
  caption?: string;
  startLine?: number;
  endLine?: number;
}

export interface HeadingData {
  level: number;
  text: string;
  line?: number;
}

export interface ContentBlock {
  type: 'heading' | 'paragraph' | 'list' | 'table' | 'code' | 'hr';
  content: string | TableData | string[];
  level?: number;
  language?: string;
  ordered?: boolean;
}

/**
 * Parse Markdown content to structured data
 */
export function parseMarkdown(markdown: string): ParsedMarkdown {
  const md = new MarkdownIt({
    html: true,
    linkify: true,
    typographer: true,
  });

  const tokens = md.parse(markdown, {});
  const tables = extractTables(tokens);
  const headings = extractHeadings(tokens);
  const content = extractContent(tokens);

  return {
    tokens,
    tables,
    headings,
    content,
  };
}

/**
 * Extract tables from markdown tokens
 */
export function extractTables(tokens: Token[]): TableData[] {
  const tables: TableData[] = [];
  let currentTable: Partial<TableData> | null = null;
  let currentRow: string[] = [];
  let isHeader = false;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    if (token.type === 'table_open') {
      currentTable = {
        headers: [],
        rows: [],
        startLine: token.map?.[0],
      };
    } else if (token.type === 'table_close') {
      if (currentTable && currentTable.headers && currentTable.rows) {
        currentTable.endLine = token.map?.[1];
        tables.push(currentTable as TableData);
      }
      currentTable = null;
    } else if (token.type === 'thead_open') {
      isHeader = true;
    } else if (token.type === 'thead_close') {
      isHeader = false;
    } else if (token.type === 'tr_open') {
      currentRow = [];
    } else if (token.type === 'tr_close') {
      if (currentTable) {
        if (isHeader) {
          currentTable.headers = [...currentRow];
        } else {
          currentTable.rows!.push([...currentRow]);
        }
      }
      currentRow = [];
    } else if (token.type === 'td_open' || token.type === 'th_open') {
      // Next token will be inline with cell content
      const nextToken = tokens[i + 1];
      if (nextToken && nextToken.type === 'inline') {
        currentRow.push(nextToken.content);
      }
    }
  }

  return tables;
}

/**
 * Extract headings from markdown tokens
 */
export function extractHeadings(tokens: Token[]): HeadingData[] {
  const headings: HeadingData[] = [];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    
    if (token.type === 'heading_open') {
      const level = parseInt(token.tag.substring(1), 10); // h1 -> 1, h2 -> 2, etc.
      const nextToken = tokens[i + 1];
      
      if (nextToken && nextToken.type === 'inline') {
        headings.push({
          level,
          text: nextToken.content,
          line: token.map?.[0],
        });
      }
    }
  }

  return headings;
}

/**
 * Extract all content blocks from markdown tokens
 */
export function extractContent(tokens: Token[]): ContentBlock[] {
  const content: ContentBlock[] = [];
  let i = 0;

  while (i < tokens.length) {
    const token = tokens[i];

    if (token.type === 'heading_open') {
      const level = parseInt(token.tag.substring(1), 10);
      const nextToken = tokens[i + 1];
      if (nextToken && nextToken.type === 'inline') {
        content.push({
          type: 'heading',
          content: nextToken.content,
          level,
        });
      }
      i += 3; // Skip heading_open, inline, heading_close
    } else if (token.type === 'paragraph_open') {
      const nextToken = tokens[i + 1];
      if (nextToken && nextToken.type === 'inline') {
        content.push({
          type: 'paragraph',
          content: nextToken.content,
        });
      }
      i += 3; // Skip paragraph_open, inline, paragraph_close
    } else if (token.type === 'bullet_list_open' || token.type === 'ordered_list_open') {
      const items = extractListItems(tokens, i);
      content.push({
        type: 'list',
        content: items,
        ordered: token.type === 'ordered_list_open',
      });
      // Skip all list tokens
      while (i < tokens.length && tokens[i].type !== 'bullet_list_close' && tokens[i].type !== 'ordered_list_close') {
        i++;
      }
      i++; // Skip the close token
    } else if (token.type === 'table_open') {
      const tableData = extractSingleTable(tokens, i);
      if (tableData) {
        content.push({
          type: 'table',
          content: tableData,
        });
      }
      // Skip all table tokens
      while (i < tokens.length && tokens[i].type !== 'table_close') {
        i++;
      }
      i++; // Skip the close token
    } else if (token.type === 'fence' || token.type === 'code_block') {
      content.push({
        type: 'code',
        content: token.content,
        language: token.info || undefined,
      });
      i++;
    } else if (token.type === 'hr') {
      content.push({
        type: 'hr',
        content: '',
      });
      i++;
    } else {
      i++;
    }
  }

  return content;
}

/**
 * Extract list items from tokens starting at a specific index
 */
function extractListItems(tokens: Token[], startIndex: number): string[] {
  const items: string[] = [];
  let i = startIndex + 1; // Skip list_open

  while (i < tokens.length) {
    const token = tokens[i];
    
    if (token.type === 'list_item_open') {
      // Look for inline content in the list item
      const inlineToken = tokens[i + 2]; // list_item_open, paragraph_open, inline
      if (inlineToken && inlineToken.type === 'inline') {
        items.push(inlineToken.content);
      }
    } else if (token.type === 'bullet_list_close' || token.type === 'ordered_list_close') {
      break;
    }
    
    i++;
  }

  return items;
}

/**
 * Extract a single table starting at a specific token index
 */
function extractSingleTable(tokens: Token[], startIndex: number): TableData | null {
  let currentTable: Partial<TableData> = {
    headers: [],
    rows: [],
  };
  let currentRow: string[] = [];
  let isHeader = false;
  let i = startIndex;

  while (i < tokens.length) {
    const token = tokens[i];

    if (token.type === 'table_close') {
      break;
    } else if (token.type === 'thead_open') {
      isHeader = true;
    } else if (token.type === 'thead_close') {
      isHeader = false;
    } else if (token.type === 'tr_open') {
      currentRow = [];
    } else if (token.type === 'tr_close') {
      if (isHeader) {
        currentTable.headers = [...currentRow];
      } else {
        currentTable.rows!.push([...currentRow]);
      }
      currentRow = [];
    } else if (token.type === 'td_open' || token.type === 'th_open') {
      const nextToken = tokens[i + 1];
      if (nextToken && nextToken.type === 'inline') {
        currentRow.push(nextToken.content);
      }
    }

    i++;
  }

  if (currentTable.headers && currentTable.rows) {
    return currentTable as TableData;
  }

  return null;
}

/**
 * Check if content contains horizontal rule (slide separator for PPTX)
 */
export function hasHorizontalRules(tokens: Token[]): boolean {
  return tokens.some(token => token.type === 'hr');
}

/**
 * Split content by horizontal rules for slide generation
 */
export function splitByHorizontalRules(tokens: Token[]): Token[][] {
  const slides: Token[][] = [];
  let currentSlide: Token[] = [];

  for (const token of tokens) {
    if (token.type === 'hr') {
      if (currentSlide.length > 0) {
        slides.push(currentSlide);
        currentSlide = [];
      }
    } else {
      currentSlide.push(token);
    }
  }

  if (currentSlide.length > 0) {
    slides.push(currentSlide);
  }

  return slides;
}

