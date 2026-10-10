/**
 * @fileoverview Excel 模板读写辅助工具函数。
 * 包含解析 Excel 模板文件、往指定单元格回写数据、保留样式并导出等功能。
 * 遵循 Google TypeScript 编码标准。
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";

/**
 * 单元格映射项定义
 */
export interface CellMapping {
  /** 单元格坐标，例如 "B3" */
  cellRef: string;
  /** 字段的唯一标识符，例如 "employeeName" */
  fieldName: string;
  /** 字段的可读名称，例如 "员工姓名" */
  label: string;
  /** 字段类型：input（用户输入）还是 ai（AI 自动生成） */
  type: "input" | "ai";
  /** 针对 AI 字段的提示说明，指导 AI 该如何生成此项内容 */
  aiInstruction?: string;
}

/**
 * 加载 Excel 文件并返回 ExcelJS Workbook 对象
 *
 * @param file 浏览器上传的文件对象
 * @returns Promise<ExcelJS.Workbook> ExcelJS 工作簿对象
 */
export async function loadWorkbook(file: File): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

/**
 * 将数据（键值对）写入 Excel 模板并返回生成的 Excel 文件的 ArrayBuffer。
 * 该操作会在浏览器本地内存中复制并写入，原样保留 Excel 的所有样式、边框、字体及公式。
 *
 * @param templateBuffer 原始 Excel 模板的 ArrayBuffer 二进制流
 * @param data 需要写入的数据，键为单元格坐标（如 "B3"），值为写入的内容（字符串、数字等）
 * @returns Promise<ArrayBuffer> 写入数据后生成的 Excel 文件 ArrayBuffer
 */
export async function writeDataToTemplate(templateBuffer: ArrayBuffer, data: Record<string, string | number>): Promise<ArrayBuffer> {
  // 创建一个新的 Workbook 对象并加载模板
  const workbook = new ExcelJS.Workbook();
  // 必须克隆或直接从 buffer 加载，避免污染原始的 templateBuffer
  await workbook.xlsx.load(templateBuffer.slice(0));

  // 默认对第一个工作表（Sheet）进行操作
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("未能在上传的 Excel 模板中找到有效的工作表 (Worksheet)。");
  }

  // 遍历数据并回写单元格
  for (const [cellRef, value] of Object.entries(data)) {
    const cell = worksheet.getCell(cellRef);

    // 如果值是数字，尝试将其转换为 Number 类型，以便 Excel 能够正确识别为数值并应用公式计算
    if (typeof value === "string" && value.trim() !== "" && !isNaN(Number(value))) {
      cell.value = Number(value);
    } else {
      cell.value = value;
    }
  }

  // 重新计算计算公式（ExcelJS 会在导出时重置公式的缓存值，由 Excel 打开时重新计算）
  // 写入 Buffer 并返回
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}

/**
 * 在浏览器中触发下载生成的 Excel 文件
 *
 * @param buffer Excel 文件的二进制 ArrayBuffer
 * @param fileName 导出的文件名，例如 "张三_2026年6月绩效表.xlsx"
 */
export function downloadExcelFile(buffer: ArrayBuffer, fileName: string): void {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;
  document.body.appendChild(link);
  link.click();

  // 清理 URL 对象
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * 正式员工绩效表单任务字段定义
 */
export interface PerformanceTask {
  type: string; // 指标类型，如 KPI, CPI
  level: string; // 指标等级，如 重要关键任务
  weight: number | string; // 权重，如 0.3 或 "扣分项"
  category: string; // 所属板块
  description: string; // 解释说明
  time_target: string; // 时间目标
  count_target: string; // 数量目标
  quality_target: string; // 质量目标
  time_standard: string; // 时间标准
  count_standard: string; // 数量标准
  quality_standard: string; // 质量标准
}

/**
 * 正式员工绩效表头部元数据字段定义
 */
export interface FormalHeaderInfo {
  /** 所属公司 (回写到 D2) */
  company?: string;
  /** 被考核人姓名 (回写到 D3) */
  name: string;
  /** 所属部门 (回写到 H3) */
  department?: string;
  /** 岗位名称 (回写到 L3) */
  position: string;
  /** 考核人姓名 (回写到 D4) */
  evaluator?: string;
  /** 考核人所属部门 (回写到 H4) */
  evaluatorDepartment?: string;
  /** 考核人岗位 (回写到 L4) */
  evaluatorPosition?: string;
}

// ============================================================
// 内部辅助：XML 特殊字符转义
// ============================================================
function escapeXml(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

/**
 * 内部辅助：将一个字符串添加到 sharedStrings.xml 中（若已存在则复用），
 * 返回该字符串的索引（从 0 开始）
 */
function upsertSharedString(ssXml: string, value: string): { index: number; newSsXml: string } {
  // 找到所有 <si>...</si> 条目
  const siMatches = [...ssXml.matchAll(/<si>[\s\S]*?<\/si>/g)];
  for (let i = 0; i < siMatches.length; i++) {
    const si = siMatches[i][0];
    // 提取所有 <t>...</t> 中的文本（合并多段富文本）
    const tParts = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
    const existing = tParts
      .map((m) =>
        m[1]
          .replace(/&amp;/g, "&")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'"),
      )
      .join("");
    if (existing === value) {
      return { index: i, newSsXml: ssXml };
    }
  }

  // 字符串不存在：追加新条目到 </sst> 前
  const newSi = `<si><t xml:space="preserve">${escapeXml(value)}</t></si>`;
  let newSsXml = ssXml.replace(/<\/sst>/, `${newSi}</sst>`);

  // 同步更新 count 和 uniqueCount 属性
  const newIndex = siMatches.length;
  newSsXml = newSsXml.replace(/(<sst[^>]* count=")(\d+)("[^>]* uniqueCount=")(\d+)(")/, (_, p1, _c, p3, _u, p5) => `${p1}${newIndex + 1}${p3}${newIndex + 1}${p5}`);
  return { index: newIndex, newSsXml };
}

/**
 * 默认固定的正式员工绩效表头部元数据常量（与公司正式绩效模板保持一致）
 */
export const DEFAULT_FORMAL_COMPANY = "四川久宏川科技有限公司";
export const DEFAULT_FORMAL_DEPARTMENT = "软件研发部";
export const DEFAULT_FORMAL_EVALUATOR = "李杰、张剑锋";
export const DEFAULT_FORMAL_EVALUATOR_DEPARTMENT = "软件研发部";
export const DEFAULT_FORMAL_EVALUATOR_POSITION = "软件研发部经理、主管";

/**
 * 任务样式解析结果接口
 */
interface TaskResolvedStyles {
  /** 任务行居中对齐单元格样式索引（包含 10号黑色字体、全封闭细边框、居中、自动换行） */
  centerStyleId: number;
  /** 任务行居左对齐单元格样式索引（质量目标与质量标准专用，包含 10号黑色字体、全封闭细边框、居左、自动换行） */
  leftStyleId: number;
  /** 头部信息填写区域样式索引（包含 12号黑色常规字体、全封闭细边框、居中对齐） */
  headerStyleId: number;
  /** 若向 cellXfs 追加了新样式，返回更新后的 styles.xml 文本 */
  newStylesXml?: string;
}

/**
 * 内部辅助：动态解析并确保工作表中的工作考核项与头部信息样式。
 * 针对用户上传的任意模板：
 * 1. 严格筛选非红色、非彩色、10-11pt 的宋体黑色字体 (Regular，不加粗) 用于考核任务项；
 * 2. 严格筛选非红色、12pt 的宋体黑色常规字体 (Regular，不加粗) 用于头部填写项（姓名、部门、岗位、考核人等）；
 * 3. 严格筛选上、下、左、右四边均具备完整细边框（border style="thin"）的边框配置，根除缺右边框问题；
 * 4. 分别精确定位垂直居中且开启自动换行 (wrapText) 的“居中对齐”与“居左对齐”样式；
 * 5. 【关键架构防护】：严格限定在 <cellXfs> 作用域内解析，杜绝与外部 <cellStyleXfs> 混淆，
 *    确保解析出的 styleId 与工作表中单元格的 s 属性 1:1 严格对齐；
 * 6. 若模板缺失对应样式，则动态、无侵入地在 <cellXfs> 追加规范样式，保证极端情况下亦 100% 合规。
 *
 * @param stylesXml xl/styles.xml 原始文本
 * @returns TaskResolvedStyles 包含解析出的居中样式 ID、居左样式 ID、头部12号字样式 ID 与可能更新的 styles.xml
 */
function resolveTaskStyles(stylesXml: string): TaskResolvedStyles {
  const fonts = [...stylesXml.matchAll(/<font>([\s\S]*?)<\/font>/g)].map((m) => m[1]);
  const borders = [...stylesXml.matchAll(/<border[^>]*>([\s\S]*?)<\/border>/g)].map((m) => m[1]);

  // ★ 核心修复：仅提取 <cellXfs> 作用域内的单元格格式，绝不匹配外部的 <cellStyleXfs>
  // Excel/WPS 中的单元格样式属性 s 严格对应 <cellXfs> 的 0-based 下标。若匹配到 <cellStyleXfs>，
  // 会导致样式下标偏移（如偏移 49），使原本居左的样式错误指向百分比居中样式。
  const cellXfsMatch = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
  const cellXfsInner = cellXfsMatch ? cellXfsMatch[1] : "";
  const cellXfs = [...cellXfsInner.matchAll(/<xf\s+([^>]+?)(?:\/>|>([\s\S]*?)<\/xf>)/gs)];

  let centerStyleId = -1;
  let leftStyleId = -1;
  let headerStyleId = -1;

  for (let i = 0; i < cellXfs.length; i++) {
    const attrs = cellXfs[i][1];
    const inner = cellXfs[i][2] || "";

    const fontIdMatch = attrs.match(/fontId="(\d+)"/);
    const borderIdMatch = attrs.match(/borderId="(\d+)"/);
    if (!fontIdMatch || !borderIdMatch) continue;

    const fontId = parseInt(fontIdMatch[1], 10);
    const borderId = parseInt(borderIdMatch[1], 10);
    const font = fonts[fontId] || "";
    const border = borders[borderId] || "";

    // 1. 字体颜色安全审计：严禁红色或醒目提示色（如 FFFF0000、FFC00000、FF9C0006 等），且常规内容必须不加粗
    const isRed = font.includes("FFFF0000") || font.includes("FFC00000") || font.includes("FF9C0006");
    const isBold = font.includes("<b/>") || font.includes("<b>");
    if (isRed || isBold) continue;

    // 2. 边框完整性审计：四边（left, right, top, bottom）必须均显式声明了 style 边框线，杜绝缺边框
    const hasLeft = border.includes("<left style=");
    const hasRight = border.includes("<right style=");
    const hasTop = border.includes("<top style=");
    const hasBottom = border.includes("<bottom style=");
    if (!hasLeft || !hasRight || !hasTop || !hasBottom) continue;

    // 3. 数字格式审计：优先匹配通用常规格式 (numFmtId="0")，避免误用百分比或特定日期格式
    const isGeneralFmt = attrs.includes('numFmtId="0"');
    if (!isGeneralFmt) continue;

    // 4. 任务项基础排版审计：必须开启文字自动换行与垂直居中（任务项为 10号常规字）
    const hasWrap = inner.includes('wrapText="1"');
    const isVcenter = inner.includes('vertical="center"');
    if (hasWrap && isVcenter) {
      if (inner.includes('horizontal="center"')) {
        if (centerStyleId === -1 || fontId === 19) {
          centerStyleId = i;
        }
      }
      if (inner.includes('horizontal="left"')) {
        if (leftStyleId === -1 || fontId === 19) {
          leftStyleId = i;
        }
      }
    }

    // 5. 头部填写项样式匹配（必须为 12号字，黑色宋体常规不加粗，居中对齐）
    const isSz12 = font.includes('sz val="12"') || font.includes('sz val="12.0"');
    const isCenter = inner.includes('horizontal="center"');
    if (isSz12 && isCenter) {
      if (headerStyleId === -1 || fontId === 21) {
        headerStyleId = i;
      }
    }
  }

  // 6. 兜底保护：若用户上传的精简模板未预置对应对齐样式，安全无损追加到 <cellXfs>
  let newStylesXml: string | undefined = undefined;
  if (centerStyleId === -1 || leftStyleId === -1 || headerStyleId === -1) {
    let workingXml = stylesXml;
    let blackFontId = -1;
    let blackFont12Id = -1;
    for (let f = 0; f < fonts.length; f++) {
      const font = fonts[f];
      const isRed = font.includes("FFFF0000") || font.includes("FFC00000") || font.includes("FF9C0006");
      const isBold = font.includes("<b/>") || font.includes("<b>");
      if (!isRed && !isBold) {
        if (blackFontId === -1) blackFontId = f;
        if (f === 19) blackFontId = f;
        const isSz12 = font.includes('sz val="12"') || font.includes('sz val="12.0"');
        if (isSz12) {
          if (blackFont12Id === -1 || f === 21) blackFont12Id = f;
        }
      }
    }
    if (blackFontId === -1) blackFontId = 0;
    if (blackFont12Id === -1) blackFont12Id = blackFontId;

    let thinBorderId = -1;
    for (let b = 0; b < borders.length; b++) {
      const border = borders[b];
      if (border.includes("<left style=") && border.includes("<right style=") && border.includes("<top style=") && border.includes("<bottom style=")) {
        thinBorderId = b;
        if (b === 5) break;
      }
    }
    if (thinBorderId === -1) thinBorderId = 0;

    let currentCellXfsCount = cellXfs.length;

    if (centerStyleId === -1) {
      centerStyleId = currentCellXfsCount++;
      const centerXml = `<xf numFmtId="0" fontId="${blackFontId}" fillId="0" borderId="${thinBorderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`;
      workingXml = workingXml.replace(/<\/cellXfs>/, `${centerXml}</cellXfs>`);
      workingXml = workingXml.replace(/(<cellXfs\s+count=")(\d+)(")/, (_, p1, c, p3) => `${p1}${parseInt(c, 10) + 1}${p3}`);
    }

    if (leftStyleId === -1) {
      leftStyleId = currentCellXfsCount++;
      const leftXml = `<xf numFmtId="0" fontId="${blackFontId}" fillId="0" borderId="${thinBorderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>`;
      workingXml = workingXml.replace(/<\/cellXfs>/, `${leftXml}</cellXfs>`);
      workingXml = workingXml.replace(/(<cellXfs\s+count=")(\d+)(")/, (_, p1, c, p3) => `${p1}${parseInt(c, 10) + 1}${p3}`);
    }

    if (headerStyleId === -1) {
      headerStyleId = currentCellXfsCount++;
      const headerXml = `<xf numFmtId="0" fontId="${blackFont12Id}" fillId="0" borderId="${thinBorderId}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>`;
      workingXml = workingXml.replace(/<\/cellXfs>/, `${headerXml}</cellXfs>`);
      workingXml = workingXml.replace(/(<cellXfs\s+count=")(\d+)(")/, (_, p1, c, p3) => `${p1}${parseInt(c, 10) + 1}${p3}`);
    }
    newStylesXml = workingXml;
  }

  return { centerStyleId, leftStyleId, headerStyleId, newStylesXml };
}

/**
 * 内部辅助：根据任务各项文本内容估算行高。
 * 质量目标、质量标准及解释说明常包含分条换行或较长说明文本，
 * 通过段落换行与估算列宽换行综合计算总行数，确保所有内容完整展示，杜绝上下截断。
 *
 * @param task 单项绩效考核任务
 * @returns 适配的行高数值（单位：pt）
 */
function estimateTaskRowHeight(task: PerformanceTask): number {
  const getLineCount = (text: string | undefined, charsPerLine: number): number => {
    if (!text) return 1;
    const paragraphs = text.split(/\r?\n/);
    let totalLines = 0;
    for (const p of paragraphs) {
      const len = p.trim().length;
      totalLines += Math.max(1, Math.ceil(len / charsPerLine));
    }
    return totalLines;
  };

  // F列 解释说明 (列宽约 31.5, 容纳约 16 个中文字符)
  const linesF = getLineCount(task.description, 16);
  // I列 质量目标 (列宽约 24.1, 容纳约 12 个中文字符)
  const linesI = getLineCount(task.quality_target, 12);
  // L列 质量标准 (列宽约 35.2, 容纳约 18 个中文字符)
  const linesL = getLineCount(task.quality_standard, 18);

  const maxLines = Math.max(1, linesF, linesI, linesL);
  if (maxLines <= 1) return 27;
  if (maxLines === 2) return 45;
  if (maxLines === 3) return 60;
  return Math.min(120, 20 + maxLines * 16);
}

/**
 * 内部辅助：在 sheet XML 中更新指定行的行高并设置 customHeight="1"
 *
 * @param sheetXml sheet 工作表 XML 内容
 * @param rn 行号，如 10
 * @param height 目标行高
 * @returns 更新后的 sheet XML
 */
function updateRowHeight(sheetXml: string, rn: number, height: number): string {
  const rowRegex = new RegExp(`<row r="${rn}"([^>]*)>`);
  return sheetXml.replace(rowRegex, (_match, attrs) => {
    // 关键修复：必须使用单词边界 \b 防止误伤 customHeight 内部的 ht=" 字符串导致生成畸形 customHeig 属性损坏 Excel
    const cleanAttrs = attrs
      .replace(/\s*\bht="[^"]*"/g, "")
      .replace(/\s*\bcustomHeight="[^"]*"/g, "")
      .trim();
    const prefix = cleanAttrs ? ` ${cleanAttrs}` : "";
    return `<row r="${rn}"${prefix} ht="${height}" customHeight="1">`;
  });
}

/**
 * 内部辅助：在 sheet XML 中将指定单元格设置为字符串（sharedStrings 索引）
 * 并可指定样式索引 s。若未传入 styleId，则完整保留单元格原有的 s 属性，绝不丢失原有格式。
 *
 * @param sheetXml sheet 工作表 XML 内容
 * @param cellRef 单元格坐标，如 "D3"
 * @param strIndex sharedStrings.xml 中的字符串索引
 * @param styleId 可选指定的样式索引，若未传则保留单元格已有样式
 * @returns 替换更新后的 sheet XML 内容
 */
function patchCellStrWithStyle(sheetXml: string, cellRef: string, strIndex: number, styleId?: number): string {
  const re = new RegExp(`(<c r="${cellRef}")(?: s="(\\d+)")?(?: t="[^"]*")?(>(?:<v>[^<]*<\\/v>)?<\\/c>|\\/>)`, "g");
  let matched = false;
  const result = sheetXml.replace(re, (_, prefix, originalStyleId) => {
    matched = true;
    // 优先使用显式指定的 styleId，若未指定则安全保留原本单元格携带的样式 ID
    const finalStyleId = styleId !== undefined ? styleId : originalStyleId;
    const sAttr = finalStyleId !== undefined ? ` s="${finalStyleId}"` : "";
    return `${prefix}${sAttr} t="s"><v>${strIndex}</v></c>`;
  });
  if (matched) return result;

  // 兜底防御：若单元格在对应行中完全缺失，安全插入到对应行内
  const colMatch = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!colMatch) return sheetXml;
  const col = colMatch[1];
  const rn = colMatch[2];
  const rowRegex = new RegExp(`(<row r="${rn}"[^>]*>)([\\s\\S]*?)(<\\/row>)`);
  return sheetXml.replace(rowRegex, (_m, openTag, inner, closeTag) => {
    const sAttr = styleId !== undefined ? ` s="${styleId}"` : "";
    const newCellXml = `<c r="${cellRef}"${sAttr} t="s"><v>${strIndex}</v></c>`;
    const cells = [...inner.matchAll(/<c r="([A-Z]+)\d+"[^>]*>[\s\S]*?<\/c>|<c r="([A-Z]+)\d+"[^>]*\/>/g)];
    for (const c of cells) {
      const cCol = c[1] || "";
      if (cCol.length > col.length || (cCol.length === col.length && cCol > col)) {
        const idx = inner.indexOf(c[0]);
        return `${openTag}${inner.slice(0, idx)}${newCellXml}${inner.slice(idx)}${closeTag}`;
      }
    }
    return `${openTag}${inner}${newCellXml}${closeTag}`;
  });
}

/**
 * 内部辅助：在 sheet XML 中将指定单元格设置为数字并可指定样式索引 s。
 * 若未传入 styleId，则完整保留单元格原有的 s 属性。
 *
 * @param sheetXml sheet 工作表 XML 内容
 * @param cellRef 单元格坐标，如 "A10"
 * @param num 需要写入的数值
 * @param styleId 可选指定的样式索引
 * @returns 替换更新后的 sheet XML 内容
 */
function patchCellNumWithStyle(sheetXml: string, cellRef: string, num: number, styleId?: number): string {
  const re = new RegExp(`(<c r="${cellRef}")(?: s="(\\d+)")?(?: t="[^"]*")?(>(?:<v>[^<]*<\\/v>)?<\\/c>|\\/>)`, "g");
  let matched = false;
  const result = sheetXml.replace(re, (_, prefix, originalStyleId) => {
    matched = true;
    const finalStyleId = styleId !== undefined ? styleId : originalStyleId;
    const sAttr = finalStyleId !== undefined ? ` s="${finalStyleId}"` : "";
    return `${prefix}${sAttr}><v>${num}</v></c>`;
  });
  if (matched) return result;

  const colMatch = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!colMatch) return sheetXml;
  const col = colMatch[1];
  const rn = colMatch[2];
  const rowRegex = new RegExp(`(<row r="${rn}"[^>]*>)([\\s\\S]*?)(<\\/row>)`);
  return sheetXml.replace(rowRegex, (_m, openTag, inner, closeTag) => {
    const sAttr = styleId !== undefined ? ` s="${styleId}"` : "";
    const newCellXml = `<c r="${cellRef}"${sAttr}><v>${num}</v></c>`;
    const cells = [...inner.matchAll(/<c r="([A-Z]+)\d+"[^>]*>[\s\S]*?<\/c>|<c r="([A-Z]+)\d+"[^>]*\/>/g)];
    for (const c of cells) {
      const cCol = c[1] || "";
      if (cCol.length > col.length || (cCol.length === col.length && cCol > col)) {
        const idx = inner.indexOf(c[0]);
        return `${openTag}${inner.slice(0, idx)}${newCellXml}${inner.slice(idx)}${closeTag}`;
      }
    }
    return `${openTag}${inner}${newCellXml}${closeTag}`;
  });
}

/**
 * 内部辅助：清空单元格内容（变为空单元格，保留样式）
 *
 * @param sheetXml sheet 工作表 XML 内容
 * @param cellRef 单元格坐标，如 "A12"
 * @param styleId 可选指定的样式索引，若未传则保留原有样式
 * @returns 替换更新后的 sheet XML 内容
 */
function patchCellEmpty(sheetXml: string, cellRef: string, styleId?: number): string {
  const re = new RegExp(`(<c r="${cellRef}")(?: s="(\\d+)")?(?: t="[^"]*")?(>(?:<v>[^<]*<\\/v>)?<\\/c>|\\/>)`, "g");
  return sheetXml.replace(re, (_, prefix, originalStyleId) => {
    const finalStyleId = styleId !== undefined ? styleId : originalStyleId;
    const sAttr = finalStyleId !== undefined ? ` s="${finalStyleId}"` : "";
    return `${prefix}${sAttr}/>`;
  });
}

/**
 * 将正式员工月度绩效任务与头部信息写入用户上传的 Excel 模板。
 *
 * 【核心架构设计准则】：
 * 1. 100% 保留模板原生样式与色彩：不修改 xl/styles.xml，完整保留用户模板的全部主题色、填充与边框；
 * 2. 动态样式采样与继承：动态从模板第 10 行采样各列原始样式 (s 属性)，确保后续所有任务行（含动态插入行与清空行）
 *    各列样式严格统一，杜绝样式硬编码与错位乱序；
 * 3. 头部信息无损回写：回写 D2、D3、H3、L3、D4、H4、L4、D5 时完整继承模板单元格自带样式；
 * 4. 纯净定向修改：仅定位并修改目标工作表与共享字符串表，其余工作表（管理层、说明表等）一概不做任何改动。
 *
 * @param templateBuffer 原始模板文件的 ArrayBuffer
 * @param nameOrInfo 员工姓名或头部信息对象
 * @param positionOrTasks 岗位名称或任务列表
 * @param tasksOrMonth 任务列表或考核月份
 * @param targetMonthOrExtra 考核月份或附加信息
 * @param extraInfo 附加头部信息 (公司、部门、考核人等)
 * @returns Promise<ArrayBuffer> 修改后的 Excel 文件二进制 buffer
 */
export async function writePerformanceToTemplate(templateBuffer: ArrayBuffer, nameOrInfo: string | FormalHeaderInfo, positionOrTasks: string | PerformanceTask[], tasksOrMonth?: PerformanceTask[] | string, targetMonthOrExtra?: string | Partial<FormalHeaderInfo>, extraInfo?: Partial<FormalHeaderInfo>): Promise<ArrayBuffer> {
  // ── 1. 参数归一化 ──
  let name = "";
  let position = "";
  let company = "";
  let department = "";
  let evaluator = "";
  let evaluatorDepartment = "";
  let evaluatorPosition = "";
  let tasks: PerformanceTask[] = [];
  let targetMonth = "";

  if (typeof nameOrInfo === "object" && nameOrInfo !== null) {
    name = nameOrInfo.name || "";
    position = nameOrInfo.position || "";
    company = nameOrInfo.company || "";
    department = nameOrInfo.department || "";
    evaluator = nameOrInfo.evaluator || "";
    evaluatorDepartment = nameOrInfo.evaluatorDepartment || "";
    evaluatorPosition = nameOrInfo.evaluatorPosition || "";
    tasks = (positionOrTasks as PerformanceTask[]) || [];
    targetMonth = (tasksOrMonth as string) || "";
  } else {
    name = nameOrInfo || "";
    position = (positionOrTasks as string) || "";
    tasks = (tasksOrMonth as PerformanceTask[]) || [];
    targetMonth = (targetMonthOrExtra as string) || "";
    if (extraInfo) {
      company = extraInfo.company || "";
      department = extraInfo.department || "";
      evaluator = extraInfo.evaluator || "";
      evaluatorDepartment = extraInfo.evaluatorDepartment || "";
      evaluatorPosition = extraInfo.evaluatorPosition || "";
    }
  }

  // ── 2. 解析年月 ──
  let year = new Date().getFullYear();
  let month = new Date().getMonth() + 1;
  let lastDay = new Date(year, month, 0).getDate();
  if (targetMonth && typeof targetMonth === "string") {
    const parts = targetMonth.split("-");
    if (parts.length === 2) {
      const py = parseInt(parts[0], 10);
      const pm = parseInt(parts[1], 10);
      if (!isNaN(py) && !isNaN(pm)) {
        year = py;
        month = pm;
        lastDay = new Date(year, month, 0).getDate();
      }
    }
  }

  // ── 3. 用 JSZip 加载 xlsx ──
  // 注意：严禁修改 xl/styles.xml，完整保留用户上传模板的所有主题填充色、表格边框与多 Sheet 样式
  const zip = await JSZip.loadAsync(templateBuffer.slice(0));

  // 从 workbook.xml 查找"绩效计划表（基层员工）"对应的 sheet 路径
  const workbookXmlFile = zip.file("xl/workbook.xml");
  if (!workbookXmlFile) throw new Error("模板 Excel 格式异常：找不到 workbook.xml");
  const workbookXml = await workbookXmlFile.async("text");

  const targetSheetName = "绩效计划表（基层员工）";
  let sheetPath = "xl/worksheets/sheet1.xml"; // 默认路径

  // 查找 sheet 的 r:id
  const sheetNodeMatch = workbookXml.match(new RegExp(`<sheet[^>]+name="[^"]*${targetSheetName.replace(/[()（）]/g, '[^"]*')}[^"]*"[^>]+r:id="(rId\\d+)"`));
  if (sheetNodeMatch) {
    const rId = sheetNodeMatch[1];
    const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
    if (wbRelsFile) {
      const wbRels = await wbRelsFile.async("text");
      const relMatch = wbRels.match(new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"`));
      if (relMatch) {
        const target = relMatch[1];
        if (target.startsWith("/xl/")) sheetPath = target.slice(1);
        else if (target.startsWith("worksheets/")) sheetPath = `xl/${target}`;
        else sheetPath = `xl/worksheets/${target}`;
      }
    }
  }

  // ── 4. 读取 XML ──
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`找不到工作表：${sheetPath}`);
  let sheetXml = await sheetFile.async("text");

  const ssFile = zip.file("xl/sharedStrings.xml");
  if (!ssFile) throw new Error("找不到 sharedStrings.xml");
  let ssXml = await ssFile.async("text");

  // ── 4.1 动态解析工作考核项与头部信息样式（确保 10/12 号黑色字体、四周边框完整、以及质量目标/标准严格居左） ──
  const stylesFile = zip.file("xl/styles.xml");
  let stylesXml = stylesFile ? await stylesFile.async("text") : "";
  const { centerStyleId, leftStyleId, headerStyleId, newStylesXml } = resolveTaskStyles(stylesXml);
  if (newStylesXml) {
    stylesXml = newStylesXml;
  }

  // 建立列名到样式索引的映射字典：
  // 序号(A)、指标类型(B)、指标等级(C)、权重(D)、所属板块(E)、时间目标(G)、数量目标(H)、时间标准(J)、数量标准(K) 采用居中样式
  // 解释说明(F)、质量目标(I)、质量标准(L) 严格采用居左样式
  const taskColStyles: Record<string, number> = {
    A: centerStyleId,
    B: centerStyleId,
    C: centerStyleId,
    D: centerStyleId,
    E: centerStyleId,
    F: leftStyleId,
    G: centerStyleId,
    H: centerStyleId,
    I: leftStyleId, // 质量目标，严格居左对齐
    J: centerStyleId,
    K: centerStyleId,
    L: leftStyleId, // 质量标准，严格居左对齐
  };

  // ── 5. 写入字符串到单元格的统一辅助函数 ──
  const writeStr = (cellRef: string, value: string, styleId?: number) => {
    if (!value && value !== "0") return;
    const r = upsertSharedString(ssXml, value);
    ssXml = r.newSsXml;
    sheetXml = patchCellStrWithStyle(sheetXml, cellRef, r.index, styleId);
  };

  // ── 6. 填写头部区域 ──
  // 企业固定信息默认兜底：公司、部门、考核人、考核人部门、考核人岗位
  const finalCompany = (company || "").trim() || DEFAULT_FORMAL_COMPANY;
  const finalDepartment = (department || "").trim() || DEFAULT_FORMAL_DEPARTMENT;
  const finalEvaluator = (evaluator || "").trim() || DEFAULT_FORMAL_EVALUATOR;
  const finalEvaluatorDept = (evaluatorDepartment || "").trim() || DEFAULT_FORMAL_EVALUATOR_DEPARTMENT;
  const finalEvaluatorPos = (evaluatorPosition || "").trim() || DEFAULT_FORMAL_EVALUATOR_POSITION;

  // 动态字段：被考核人姓名与岗位
  const finalName = (name || "").trim() || "杨砚翔";
  const finalPosition = (position || "").trim() || "APP开发工程师";

  // A1：在 sharedStrings 中直接替换月份文字（不改 XML 结构与居中样式）
  {
    const siList = [...ssXml.matchAll(/<si>[\s\S]*?<\/si>/g)];
    if (siList.length > 0) {
      const targetIdx = siList.findIndex((si) => si[0].includes("绩效计划表") || si[0].includes("月"));
      const idxToUpdate = targetIdx >= 0 ? targetIdx : 0;
      const oldSi = siList[idxToUpdate][0];
      const newSi = oldSi.replace(/(<t[^>]*>)([\s\S]*?)(<\/t>)/, (_m, open, text, close) => {
        const newText = text.replace(/^.*?月/, `${month}月`);
        return `${open}${newText}${close}`;
      });
      ssXml = ssXml.replace(oldSi, newSi);
    }
  }

  // 写入头部单元格：
  // D2、D5 继承模板原生样式（模板原生自带 12号字）；
  // D3、H3、L3、D4、H4、L4 安全写入 12号字黑色常规居中规范样式 headerStyleId，确保填写的头部信息内容严格为 12 号字
  writeStr("D2", finalCompany);
  writeStr("D3", finalName, headerStyleId);
  writeStr("H3", finalDepartment, headerStyleId);
  writeStr("L3", finalPosition, headerStyleId);
  writeStr("D4", finalEvaluator, headerStyleId);
  writeStr("H4", finalEvaluatorDept, headerStyleId);
  writeStr("L4", finalEvaluatorPos, headerStyleId);

  // D5：考核周期（模板 D5 原生即为 12号字加粗样式）
  const periodStr = `${year} 年   ${month}   月   1 日 至 ${year} 年   ${month}   月 ${lastDay} 日`;
  writeStr("D5", periodStr);

  // ── 7. 处理任务行 ──
  const N = tasks.length;

  // 7a. 若任务数 > 4，在第 13 行之后插入 N-4 个新行（动态继承统一列样式与统一边框）
  if (N > 4) {
    const insertCount = N - 4;

    const insertedRows: string[] = [];
    for (let k = 0; k < insertCount; k++) {
      const newRn = 14 + k;
      let cellsXml = "";
      for (const col of ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]) {
        const sAttr = taskColStyles[col] !== undefined ? ` s="${taskColStyles[col]}"` : "";
        cellsXml += `<c r="${col}${newRn}"${sAttr}/>`;
      }
      const rowOpenTag = `<row r="${newRn}" customFormat="1" ht="27" customHeight="1" spans="1:12">`;
      insertedRows.push(`${rowOpenTag}${cellsXml}</row>`);
    }

    // 将行 14 及以上的行号全部向后偏移 insertCount（先处理 mergeCell，再处理单元格和行）
    sheetXml = sheetXml.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (_match, c1, r1s, c2, r2s) => {
      const r1 = parseInt(r1s, 10);
      const r2 = parseInt(r2s, 10);
      const nr1 = r1 >= 14 ? r1 + insertCount : r1;
      const nr2 = r2 >= 14 ? r2 + insertCount : r2;
      return `<mergeCell ref="${c1}${nr1}:${c2}${nr2}"`;
    });
    sheetXml = sheetXml.replace(/<c r="([A-Z]+)(\d+)"/g, (match, col, rns) => {
      const rn = parseInt(rns, 10);
      return rn >= 14 ? `<c r="${col}${rn + insertCount}"` : match;
    });
    sheetXml = sheetXml.replace(/<row r="(\d+)"/g, (match, rns) => {
      const rn = parseInt(rns, 10);
      return rn >= 14 ? `<row r="${rn + insertCount}"` : match;
    });

    // 在第 13 行 </row> 之后插入新行
    sheetXml = sheetXml.replace(/(<row r="13"[\s\S]*?<\/row>)/, `$1${insertedRows.join("")}`);

    // 同步更新 dimension 坐标范围
    sheetXml = sheetXml.replace(/(<dimension ref="[A-Z]+1:[A-Z]+)(\d+)(")/, (_, p1, lastRow, p3) => {
      const newLastRow = parseInt(lastRow, 10) + insertCount;
      return `${p1}${newLastRow}${p3}`;
    });
  }

  // 7b. 写入各行任务数据（各列严格采用统一黑色样式、全封闭细边框，质量目标与质量标准居左对齐）
  for (let i = 0; i < N; i++) {
    const rn = 10 + i;
    const task = tasks[i];

    // 自适应多行内容计算行高，确保无论是在前 4 行还是超过 4 行的新增行，排版视觉完全一致且不遮挡文字
    const rowHt = estimateTaskRowHeight(task);
    sheetXml = updateRowHeight(sheetXml, rn, rowHt);

    // 序号（A列，数字，居中）
    sheetXml = patchCellNumWithStyle(sheetXml, `A${rn}`, i + 1, taskColStyles["A"]);

    // 指标类型（B列，居中）
    {
      const r = upsertSharedString(ssXml, task.type || "KPI");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `B${rn}`, r.index, taskColStyles["B"]);
    }

    // 指标等级（C列，居中）
    {
      const r = upsertSharedString(ssXml, task.level || "重要关键任务");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `C${rn}`, r.index, taskColStyles["C"]);
    }

    // 权重（D列）：扣分项写入文字，普通百分比写入规范百分比字符串，居中
    const isDeduction = task.weight === "扣分项" || (typeof task.weight === "string" && task.weight.includes("扣分")) || task.category === "市场侧临时新增开发任务" || (task.description && task.description.includes("市场侧临时新增"));

    let weightStr = "25%";
    if (isDeduction) {
      weightStr = "扣分项";
    } else if (typeof task.weight === "string") {
      const clean = task.weight.trim();
      if (clean.includes("%")) {
        weightStr = clean;
      } else {
        const p = parseFloat(clean);
        weightStr = !isNaN(p) ? (p > 1 ? `${p}%` : `${Math.round(p * 100)}%`) : clean;
      }
    } else if (typeof task.weight === "number") {
      const p = task.weight > 1 ? task.weight : Math.round(task.weight * 100);
      weightStr = `${p}%`;
    }
    {
      const r = upsertSharedString(ssXml, weightStr);
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `D${rn}`, r.index, taskColStyles["D"]);
    }

    // 所属板块（E列，居中）
    {
      const r = upsertSharedString(ssXml, task.category || "/");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `E${rn}`, r.index, taskColStyles["E"]);
    }

    // 解释说明（F列，居左）
    {
      const r = upsertSharedString(ssXml, task.description || "");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `F${rn}`, r.index, taskColStyles["F"]);
    }

    // 时间目标（G列，居中）
    {
      const r = upsertSharedString(ssXml, task.time_target || "");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `G${rn}`, r.index, taskColStyles["G"]);
    }

    // 数量目标（H列，居中）
    {
      const r = upsertSharedString(ssXml, task.count_target || "/");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `H${rn}`, r.index, taskColStyles["H"]);
    }

    // ★ 质量目标（I列，严格居左对齐）
    {
      const r = upsertSharedString(ssXml, task.quality_target || "");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `I${rn}`, r.index, taskColStyles["I"]);
    }

    // 时间标准（J列，居中）
    {
      const r = upsertSharedString(ssXml, task.time_standard || "");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `J${rn}`, r.index, taskColStyles["J"]);
    }

    // 数量标准（K列，居中）
    {
      const r = upsertSharedString(ssXml, task.count_standard || "/");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `K${rn}`, r.index, taskColStyles["K"]);
    }

    // ★ 质量标准（L列，严格居左对齐）
    {
      const r = upsertSharedString(ssXml, task.quality_standard || "");
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `L${rn}`, r.index, taskColStyles["L"]);
    }
  }

  // 7c. 若任务数 < 4，清空剩余行（使用 taskColStyles 保持与任务行 100% 统一的排版与边框）
  if (N < 4) {
    for (let r = 10 + N; r <= 13; r++) {
      sheetXml = updateRowHeight(sheetXml, r, 27);
      for (const col of ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]) {
        sheetXml = patchCellEmpty(sheetXml, `${col}${r}`, taskColStyles[col]);
      }
    }
  }

  // ── 8. 写回 ZIP，生成 ArrayBuffer ──
  zip.file(sheetPath, sheetXml);
  zip.file("xl/sharedStrings.xml", ssXml);
  if (newStylesXml) {
    zip.file("xl/styles.xml", stylesXml);
  }

  const outputBuffer = await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  return outputBuffer;
}

/**
 * 从已填写的正式员工绩效 Excel 中读取姓名、岗位及工作考核任务项
 *
 * @param file 历史绩效 Excel 文件
 * @returns Promise<{ name: string; position: string; company?: string; department?: string; evaluator?: string; evaluatorDepartment?: string; evaluatorPosition?: string; tasks: PerformanceTask[] }>
 */
export async function readPerformanceFromExcel(file: File): Promise<{
  name: string;
  position: string;
  company?: string;
  department?: string;
  evaluator?: string;
  evaluatorDepartment?: string;
  evaluatorPosition?: string;
  tasks: PerformanceTask[];
}> {
  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();
  await workbook.xlsx.load(arrayBuffer);

  const sheetName = "绩效计划表（基层员工）";
  const worksheet = workbook.getWorksheet(sheetName) || workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("未能在导入的 Excel 中找到有效的工作表 (Worksheet)。");
  }

  // 读取头部字段：所属公司(D2)、被考核人(D3)、所属部门(H3)、岗位(L3)、考核人(D4)、考核人部门(H4)、考核人岗位(L4)
  const company = String(worksheet.getCell("D2").value || "").trim();
  const name = String(worksheet.getCell("D3").value || "").trim();
  const department = String(worksheet.getCell("H3").value || "").trim();
  const position = String(worksheet.getCell("L3").value || "").trim();
  const evaluator = String(worksheet.getCell("D4").value || "").trim();
  const evaluatorDepartment = String(worksheet.getCell("H4").value || "").trim();
  const evaluatorPosition = String(worksheet.getCell("L4").value || "").trim();

  const tasks: PerformanceTask[] = [];
  let rowNum = 10;

  while (true) {
    const row = worksheet.getRow(rowNum);
    const seqVal = row.getCell(1).value;
    const typeVal = String(row.getCell(2).value || "").trim();

    // 如果序号为空，或者指标类型为"固定项"，说明工作考核项已读完，跳出循环
    if (seqVal === null || seqVal === undefined || seqVal === "" || typeVal === "固定项") {
      break;
    }

    // 尝试读取权重值并统一化
    let weightVal: number | string = 0.25;
    const rawWeight = row.getCell(4).value;
    if (typeof rawWeight === "string") {
      const clean = rawWeight.trim();
      if (clean === "扣分项" || clean.includes("扣分")) {
        weightVal = "扣分项";
      } else {
        const parsed = parseFloat(clean.replace("%", ""));
        if (!isNaN(parsed)) {
          weightVal = parsed > 1 ? parsed / 100 : parsed;
        }
      }
    } else if (typeof rawWeight === "number") {
      weightVal = rawWeight > 1 ? rawWeight / 100 : rawWeight;
    } else if (rawWeight && typeof rawWeight === "object" && "result" in rawWeight) {
      // 针对有公式计算的情况，如果是 ExcelJS 包含的计算公式结果
      const res = (rawWeight as any).result;
      if (typeof res === "number") weightVal = res;
    }

    tasks.push({
      type: typeVal,
      level: String(row.getCell(3).value || "").trim(),
      weight: weightVal,
      category: String(row.getCell(5).value || "").trim(),
      description: String(row.getCell(6).value || "").trim(),
      time_target: String(row.getCell(7).value || "").trim(),
      count_target: String(row.getCell(8).value || "").trim(),
      quality_target: String(row.getCell(9).value || "").trim(),
      time_standard: String(row.getCell(10).value || "").trim(),
      count_standard: String(row.getCell(11).value || "").trim(),
      quality_standard: String(row.getCell(12).value || "").trim(),
    });

    rowNum++;
  }

  return {
    name,
    position,
    company: company || undefined,
    department: department || undefined,
    evaluator: evaluator || undefined,
    evaluatorDepartment: evaluatorDepartment || undefined,
    evaluatorPosition: evaluatorPosition || undefined,
    tasks,
  };
}

/**
 * 绩效质量标准库条目定义
 */
export interface QualityStandardItem {
  /** 任务/业务类型，如 "功能开发", "测试线上问题处理" */
  categoryType: string;
  /** 对应的质量目标（分条规范） */
  qualityTarget: string;
  /** 对应的质量标准（违规扣分细则） */
  qualityStandard: string;
}

/**
 * 解析绩效质量标准库 Excel 文件
 *
 * @param buffer 质量标准库 Excel 二进制 ArrayBuffer
 * @returns Promise<QualityStandardItem[]> 解析出的标准条目列表
 */
export async function parseQualityStandards(buffer: ArrayBuffer): Promise<QualityStandardItem[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer.slice(0));

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    return [];
  }

  const standards: QualityStandardItem[] = [];
  const rowCount = worksheet.rowCount;

  // 从第 2 行开始读取（第 1 行为表头）
  for (let r = 2; r <= rowCount; r++) {
    const row = worksheet.getRow(r);
    const categoryType = String(row.getCell(1).value || "").trim();
    const qualityTarget = String(row.getCell(2).value || "").trim();
    const qualityStandard = String(row.getCell(3).value || "").trim();

    if (categoryType && (qualityTarget || qualityStandard)) {
      standards.push({
        categoryType,
        qualityTarget,
        qualityStandard,
      });
    }
  }

  return standards;
}
