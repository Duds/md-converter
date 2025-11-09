/**
 * PPTX Converter
 * Convert Markdown to PowerPoint presentation
 */

// @ts-nocheck
import * as PptxGenJS from 'pptxgenjs';
import * as fs from 'fs/promises';
import * as path from 'path';
import { parseMarkdown, splitByHorizontalRules, extractContent, type ContentBlock, type TableData } from '../parsers/markdown.js';

export interface PptxConversionOptions {
  theme?: 'light' | 'dark';
  layout?: 'title-content' | 'blank' | 'section';
  fontSize?: number;
  fontFamily?: string;
}

export interface PptxConversionResult {
  success: boolean;
  outputPath: string;
  slideCount: number;
  warnings: string[];
}

/**
 * Convert Markdown file to PPTX
 */
export async function convertToPptx(
  inputPath: string,
  outputPath?: string,
  options: PptxConversionOptions = {}
): Promise<PptxConversionResult> {
  // Read markdown file
  const markdownContent = await fs.readFile(inputPath, 'utf-8');
  
  // Parse markdown
  const parsed = parseMarkdown(markdownContent);
  
  // Determine output path
  const output = outputPath || inputPath.replace(/\.md$/, '.pptx');
  
  const warnings: string[] = [];

  // Create presentation
  // @ts-ignore
  const pres = new (PptxGenJS as any).default();
  
  // Set presentation properties
  pres.author = 'MD Converter';
  pres.company = '';
  pres.subject = 'Converted from Markdown';
  pres.title = 'Presentation';

  // Apply theme
  const themeColors = getThemeColors(options.theme || 'light');
  
  // Split content by horizontal rules (slides)
  const slideGroups = splitByHorizontalRules(parsed.tokens);
  
  if (slideGroups.length === 0) {
    // No horizontal rules, treat as single content flow
    createSlidesFromContent(pres, parsed.content, options, themeColors, warnings);
  } else {
    // Create slides from groups
    for (const group of slideGroups) {
      const content = extractContent(group);
      createSlidesFromContent(pres, content, options, themeColors, warnings);
    }
  }

  // Ensure output directory exists
  await fs.mkdir(path.dirname(output), { recursive: true });
  
  // Save presentation
  await pres.writeFile({ fileName: output });

  return {
    success: true,
    outputPath: output,
    slideCount: pres.slides.length,
    warnings,
  };
}

/**
 * Get theme colors
 */
function getThemeColors(theme: 'light' | 'dark'): {
  background: string;
  text: string;
  accent: string;
} {
  if (theme === 'dark') {
    return {
      background: '1F1F1F',
      text: 'FFFFFF',
      accent: '4A9EFF',
    };
  }
  
  return {
    background: 'FFFFFF',
    text: '000000',
    accent: '0066CC',
  };
}

/**
 * Create slides from content blocks
 */
function createSlidesFromContent(
  pres: any,
  content: ContentBlock[],
  options: PptxConversionOptions,
  themeColors: { background: string; text: string; accent: string },
  _warnings: string[]
): void {
  let i = 0;
  
  while (i < content.length) {
    const block = content[i];
    
    if (block.type === 'heading' && block.level === 1) {
      // H1 = Title slide
      const nextBlock = content[i + 1];
      const subtitle = nextBlock?.type === 'paragraph' 
        ? (typeof nextBlock.content === 'string' ? nextBlock.content as string : '')
        : '';
      
      createTitleSlide(pres, block, subtitle, options, themeColors);
      i += subtitle ? 2 : 1;
    } else if (block.type === 'heading' && block.level === 2) {
      // H2 = Section header
      createSectionSlide(pres, block, options, themeColors);
      i++;
    } else if (block.type === 'heading' && block.level === 3) {
      // H3 = Content slide title
      const slideContent = collectSlideContent(content, i + 1);
      createContentSlide(pres, block, slideContent, options, themeColors);
      i += 1 + slideContent.length;
    } else {
      i++;
    }
  }
}

/**
 * Create a title slide
 */
function createTitleSlide(
  pres: any,
  titleBlock: ContentBlock,
  subtitle: string,
  options: PptxConversionOptions,
  themeColors: { background: string; text: string; accent: string }
): void {
  const slide = pres.addSlide();
  slide.background = { color: themeColors.background };
  
  const title = typeof titleBlock.content === 'string' ? titleBlock.content : '';
  
  // Add title
  slide.addText(title, {
    x: 0.5,
    y: 1.5,
    w: 9,
    h: 1.5,
    fontSize: 44,
    bold: true,
    color: themeColors.text,
    align: 'center',
    fontFace: options.fontFamily || 'Arial',
  });
  
  // Add subtitle if present
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5,
      y: 3.2,
      w: 9,
      h: 1,
      fontSize: 24,
      color: themeColors.text,
      align: 'center',
      fontFace: options.fontFamily || 'Arial',
    });
  }
}

/**
 * Create a section header slide
 */
function createSectionSlide(
  pres: any,
  headingBlock: ContentBlock,
  options: PptxConversionOptions,
  themeColors: { background: string; text: string; accent: string }
): void {
  const slide = pres.addSlide();
  slide.background = { color: themeColors.accent };
  
  const heading = typeof headingBlock.content === 'string' ? headingBlock.content : '';
  
  slide.addText(heading, {
    x: 0.5,
    y: 2.5,
    w: 9,
    h: 1.5,
    fontSize: 36,
    bold: true,
    color: 'FFFFFF',
    align: 'center',
    fontFace: options.fontFamily || 'Arial',
  });
}

/**
 * Create a content slide
 */
function createContentSlide(
  pres: any,
  titleBlock: ContentBlock,
  content: ContentBlock[],
  options: PptxConversionOptions,
  themeColors: { background: string; text: string; accent: string }
): void {
  const slide = pres.addSlide();
  slide.background = { color: themeColors.background };
  
  const title = typeof titleBlock.content === 'string' ? titleBlock.content : '';
  
  // Add title
  slide.addText(title, {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.8,
    fontSize: 32,
    bold: true,
    color: themeColors.text,
    fontFace: options.fontFamily || 'Arial',
  });
  
  // Add content
  let yPosition = 1.5;
  
  for (const block of content) {
    if (yPosition > 6.5) break; // Stop if we run out of space
    
    switch (block.type) {
      case 'paragraph': {
        const text = typeof block.content === 'string' ? block.content : '';
        slide.addText(text, {
          x: 0.5,
          y: yPosition,
          w: 9,
          h: 0.6,
          fontSize: options.fontSize || 18,
          color: themeColors.text,
          fontFace: options.fontFamily || 'Arial',
        });
        yPosition += 0.8;
        break;
      }
      
      case 'list': {
        const items = Array.isArray(block.content) ? block.content : [];
        const bulletPoints = items.map((item) => ({ text: item }));
        
        slide.addText(bulletPoints, {
          x: 0.8,
          y: yPosition,
          w: 8.5,
          h: Math.min(items.length * 0.5, 4),
          fontSize: options.fontSize || 18,
          bullet: block.ordered ? { type: 'number' } : true,
          color: themeColors.text,
          fontFace: options.fontFamily || 'Arial',
        });
        yPosition += Math.min(items.length * 0.5 + 0.3, 4.5);
        break;
      }
      
      case 'table': {
        if (typeof block.content === 'object' && 'headers' in block.content) {
          addTableToSlide(slide, block.content as TableData, yPosition, options, themeColors);
          yPosition += 3; // Tables take up more space
        }
        break;
      }
      
      case 'code': {
        const code = typeof block.content === 'string' ? block.content : '';
        slide.addText(code, {
          x: 0.5,
          y: yPosition,
          w: 9,
          h: Math.min(2, 6.5 - yPosition),
          fontSize: 14,
          fontFace: 'Courier New',
          color: themeColors.text,
          fill: { color: 'F5F5F5' },
        });
        yPosition += 2.2;
        break;
      }
    }
  }
}

/**
 * Collect content for a single slide
 */
function collectSlideContent(content: ContentBlock[], startIndex: number): ContentBlock[] {
  const slideContent: ContentBlock[] = [];
  
  for (let i = startIndex; i < content.length; i++) {
    const block = content[i];
    
    // Stop at next heading or HR
    if (block.type === 'heading' || block.type === 'hr') {
      break;
    }
    
    slideContent.push(block);
  }
  
  return slideContent;
}

/**
 * Add a table to a slide
 */
function addTableToSlide(
  slide: any,
  tableData: TableData,
  yPosition: number,
  options: PptxConversionOptions,
  themeColors: { background: string; text: string; accent: string }
): void {
  const rows: any[][] = [];
  
  // Add header row
  if (tableData.headers.length > 0) {
    rows.push(
      tableData.headers.map((header) => ({
        text: header,
        options: {
          bold: true,
          fontSize: options.fontSize || 14,
          color: 'FFFFFF',
          fill: { color: themeColors.accent },
        },
      }))
    );
  }
  
  // Add data rows
  for (const row of tableData.rows) {
    rows.push(
      row.map((cell) => ({
        text: cell,
        options: {
          fontSize: options.fontSize || 12,
          color: themeColors.text,
        },
      }))
    );
  }
  
  slide.addTable(rows, {
    x: 0.5,
    y: yPosition,
    w: 9,
    border: { type: 'solid', pt: 1, color: '999999' },
  });
}

/**
 * Convert markdown string directly to PPTX
 */
export async function convertMarkdownToPptx(
  markdown: string,
  outputPath: string,
  options: PptxConversionOptions = {}
): Promise<PptxConversionResult> {
  const parsed = parseMarkdown(markdown);
  const warnings: string[] = [];

  // @ts-ignore
  const pres = new (PptxGenJS as any).default();
  pres.author = 'MD Converter';
  
  const themeColors = getThemeColors(options.theme || 'light');
  createSlidesFromContent(pres, parsed.content, options, themeColors, warnings);
  
  await pres.writeFile({ fileName: outputPath });

  return {
    success: true,
    outputPath,
    slideCount: pres.slides.length,
    warnings,
  };
}

