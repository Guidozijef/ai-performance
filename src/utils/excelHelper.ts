/**
 * @fileoverview Excel 模板读写辅助工具函数。
 * 包含解析 Excel 模板文件、往指定单元格回写数据、保留样式并导出等功能。
 * 遵循 Google TypeScript 编码标准。
 */

import ExcelJS from 'exceljs';
import JSZip from 'jszip';

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
  type: 'input' | 'ai';
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
export async function writeDataToTemplate(
  templateBuffer: ArrayBuffer,
  data: Record<string, string | number>
): Promise<ArrayBuffer> {
  // 创建一个新的 Workbook 对象并加载模板
  const workbook = new ExcelJS.Workbook();
  // 必须克隆或直接从 buffer 加载，避免污染原始的 templateBuffer
  await workbook.xlsx.load(templateBuffer.slice(0));

  // 默认对第一个工作表（Sheet）进行操作
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('未能在上传的 Excel 模板中找到有效的工作表 (Worksheet)。');
  }

  // 遍历数据并回写单元格
  for (const [cellRef, value] of Object.entries(data)) {
    const cell = worksheet.getCell(cellRef);
    
    // 如果值是数字，尝试将其转换为 Number 类型，以便 Excel 能够正确识别为数值并应用公式计算
    if (typeof value === 'string' && value.trim() !== '' && !isNaN(Number(value))) {
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
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
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
  type: string;        // 指标类型，如 KPI, CPI
  level: string;       // 指标等级，如 重要关键任务
  weight: number | string; // 权重，如 0.3 或 "扣分项"
  category: string;    // 所属板块
  description: string; // 解释说明
  time_target: string; // 时间目标
  count_target: string;// 数量目标
  quality_target: string; // 质量目标
  time_standard: string;  // 时间标准
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
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * 内部辅助：将一个字符串添加到 sharedStrings.xml 中（若已存在则复用），
 * 返回该字符串的索引（从 0 开始）
 */
function upsertSharedString(
  ssXml: string,
  value: string
): { index: number; newSsXml: string } {
  // 找到所有 <si>...</si> 条目
  const siMatches = [...ssXml.matchAll(/<si>[\s\S]*?<\/si>/g)];
  for (let i = 0; i < siMatches.length; i++) {
    const si = siMatches[i][0];
    // 提取所有 <t>...</t> 中的文本（合并多段富文本）
    const tParts = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
    const existing = tParts
      .map(m => m[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
      )
      .join('');
    if (existing === value) {
      return { index: i, newSsXml: ssXml };
    }
  }

  // 字符串不存在：追加新条目到 </sst> 前
  const newSi = `<si><t xml:space="preserve">${escapeXml(value)}</t></si>`;
  let newSsXml = ssXml.replace(/<\/sst>/, `${newSi}</sst>`);

  // 同步更新 count 和 uniqueCount 属性
  const newIndex = siMatches.length;
  newSsXml = newSsXml.replace(
    /(<sst[^>]* count=")(\d+)("[^>]* uniqueCount=")(\d+)(")/,
    (_, p1, _c, p3, _u, p5) => `${p1}${newIndex + 1}${p3}${newIndex + 1}${p5}`
  );
  return { index: newIndex, newSsXml };
}

/**
 * 默认固定的正式员工绩效表头部元数据常量（与公司正式绩效模板保持一致）
 */
export const DEFAULT_FORMAL_COMPANY = '四川久宏川科技有限公司';
export const DEFAULT_FORMAL_DEPARTMENT = '软件研发部';
export const DEFAULT_FORMAL_EVALUATOR = '李杰、张剑锋';
export const DEFAULT_FORMAL_EVALUATOR_DEPARTMENT = '软件研发部';
export const DEFAULT_FORMAL_EVALUATOR_POSITION = '软件研发部经理、主管';

/**
 * 内部辅助：在 sheet XML 中将指定单元格设置为字符串（sharedStrings 索引）
 * 并可指定样式索引 s。若未传入 styleId，则完整保留单元格原有的 s 属性，绝不丢失原有居中格式。
 *
 * @param sheetXml sheet 工作表 XML 内容
 * @param cellRef 单元格坐标，如 "D3"
 * @param strIndex sharedStrings.xml 中的字符串索引
 * @param styleId 可选指定的样式索引，如 276 (居中 Bold)、279 (居中 Regular)
 * @returns 替换更新后的 sheet XML 内容
 */
function patchCellStrWithStyle(
  sheetXml: string,
  cellRef: string,
  strIndex: number,
  styleId?: number
): string {
  const re = new RegExp(
    `(<c r="${cellRef}")(?: s="(\\d+)")?(?: t="[^"]*")?(>(?:<v>[^<]*<\\/v>)?<\\/c>|\\/>)`,
    'g'
  );
  return sheetXml.replace(re, (_, prefix, originalStyleId) => {
    // 优先使用显式指定的 styleId，若未指定则安全保留原本单元格携带的样式 ID
    const finalStyleId = styleId !== undefined ? styleId : originalStyleId;
    const sAttr = finalStyleId !== undefined ? ` s="${finalStyleId}"` : '';
    return `${prefix}${sAttr} t="s"><v>${strIndex}</v></c>`;
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
function patchCellNumWithStyle(
  sheetXml: string,
  cellRef: string,
  num: number,
  styleId?: number
): string {
  const re = new RegExp(
    `(<c r="${cellRef}")(?: s="(\\d+)")?(?: t="[^"]*")?(>(?:<v>[^<]*<\\/v>)?<\\/c>|\\/>)`,
    'g'
  );
  return sheetXml.replace(re, (_, prefix, originalStyleId) => {
    const finalStyleId = styleId !== undefined ? styleId : originalStyleId;
    const sAttr = finalStyleId !== undefined ? ` s="${finalStyleId}"` : '';
    return `${prefix}${sAttr}><v>${num}</v></c>`;
  });
}

/**
 * 内部辅助：清空单元格内容（变为空单元格，保留样式）
 */
function patchCellEmpty(sheetXml: string, cellRef: string, styleId?: number): string {
  const re = new RegExp(
    `(<c r="${cellRef}")(?: s="(\\d+)")?(?: t="[^"]*")?(>(?:<v>[^<]*<\\/v>)?<\\/c>|\\/>)`,
    'g'
  );
  return sheetXml.replace(re, (_, prefix, originalStyleId) => {
    const finalStyleId = styleId !== undefined ? styleId : originalStyleId;
    const sAttr = finalStyleId !== undefined ? ` s="${finalStyleId}"` : '';
    return `${prefix}${sAttr}/>`;
  });
}

/**
 * 将正式员工月度绩效任务与头部信息写入特定模板。
 *
 * 【核心策略】：使用 JSZip 直接在 xlsx 原始 XML 字节层面操作，
 * 1. 样式纯白化：在 styles.xml 中将所有彩色背景填充重置为 none，确保导出的文件 100% 白底黑字、无任何彩色背景；
 * 2. 字体与排版统一：任务行各字段统一使用 10pt 宋体 Regular（字号、字重严格一致）；
 * 3. 对齐方式优化：质量目标（I 列）与质量标准（L 列）设为居左对齐并自动换行；
 * 4. 保持模板结构：行高、列宽、合并单元格及其他固定行 100% 保持原有模板结构。
 *
 * @param templateBuffer 原始模板文件的 ArrayBuffer
 * @param nameOrInfo 员工姓名或头部信息对象
 * @param positionOrTasks 岗位名称或任务列表
 * @param tasksOrMonth 任务列表或考核月份
 * @param targetMonthOrExtra 考核月份或附加信息
 * @param extraInfo 附加头部信息 (公司、部门、考核人等)
 * @returns Promise<ArrayBuffer> 修改后的 Excel 文件二进制 buffer
 */
export async function writePerformanceToTemplate(
  templateBuffer: ArrayBuffer,
  nameOrInfo: string | FormalHeaderInfo,
  positionOrTasks: string | PerformanceTask[],
  tasksOrMonth?: PerformanceTask[] | string,
  targetMonthOrExtra?: string | Partial<FormalHeaderInfo>,
  extraInfo?: Partial<FormalHeaderInfo>
): Promise<ArrayBuffer> {
  // ── 1. 参数归一化 ──
  let name = '';
  let position = '';
  let company = '';
  let department = '';
  let evaluator = '';
  let evaluatorDepartment = '';
  let evaluatorPosition = '';
  let tasks: PerformanceTask[] = [];
  let targetMonth = '';

  if (typeof nameOrInfo === 'object' && nameOrInfo !== null) {
    name = nameOrInfo.name || '';
    position = nameOrInfo.position || '';
    company = nameOrInfo.company || '';
    department = nameOrInfo.department || '';
    evaluator = nameOrInfo.evaluator || '';
    evaluatorDepartment = nameOrInfo.evaluatorDepartment || '';
    evaluatorPosition = nameOrInfo.evaluatorPosition || '';
    tasks = (positionOrTasks as PerformanceTask[]) || [];
    targetMonth = (tasksOrMonth as string) || '';
  } else {
    name = nameOrInfo || '';
    position = (positionOrTasks as string) || '';
    tasks = (tasksOrMonth as PerformanceTask[]) || [];
    targetMonth = (targetMonthOrExtra as string) || '';
    if (extraInfo) {
      company = extraInfo.company || '';
      department = extraInfo.department || '';
      evaluator = extraInfo.evaluator || '';
      evaluatorDepartment = extraInfo.evaluatorDepartment || '';
      evaluatorPosition = extraInfo.evaluatorPosition || '';
    }
  }

  // ── 2. 解析年月 ──
  let year = new Date().getFullYear();
  let month = new Date().getMonth() + 1;
  let lastDay = new Date(year, month, 0).getDate();
  if (targetMonth && typeof targetMonth === 'string') {
    const parts = targetMonth.split('-');
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
  const zip = await JSZip.loadAsync(templateBuffer.slice(0));

  // ── 3.1 全局样式纯白化：将 styles.xml 中的所有填充颜色重置为 none，确保整个文档白底黑字 ──
  const stylesFile = zip.file('xl/styles.xml');
  if (stylesFile) {
    let stylesXml = await stylesFile.async('text');
    // 将 <fills> 区域内除索引 1 (gray125) 外的所有 patternFill 统一替换为 patternType="none"，彻底消除彩色与黄色背景
    stylesXml = stylesXml.replace(/<fills[^>]*>([\s\S]*?)<\/fills>/, (match, inner) => {
      let count = 0;
      const cleanFills = inner.replace(/<fill>[\s\S]*?<\/fill>/g, () => {
        count++;
        // 保留第 2 个 (index 1) 为 gray125，其余全部为 none
        if (count === 2) {
          return '<fill><patternFill patternType="gray125"/></fill>';
        }
        return '<fill><patternFill patternType="none"/></fill>';
      });
      return `<fills count="${count}">${cleanFills}</fills>`;
    });
    zip.file('xl/styles.xml', stylesXml);
  }

  // 从 workbook.xml 查找"绩效计划表（基层员工）"对应的 sheet 路径
  const workbookXmlFile = zip.file('xl/workbook.xml');
  if (!workbookXmlFile) throw new Error('模板 Excel 格式异常：找不到 workbook.xml');
  const workbookXml = await workbookXmlFile.async('text');

  const targetSheetName = '绩效计划表（基层员工）';
  let sheetPath = 'xl/worksheets/sheet1.xml'; // 默认路径

  // 查找 sheet 的 r:id
  const sheetNodeMatch = workbookXml.match(
    new RegExp(`<sheet[^>]+name="[^"]*${targetSheetName.replace(/[()（）]/g, '[^"]*')}[^"]*"[^>]+r:id="(rId\\d+)"`)
  );
  if (sheetNodeMatch) {
    const rId = sheetNodeMatch[1];
    const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (wbRelsFile) {
      const wbRels = await wbRelsFile.async('text');
      const relMatch = wbRels.match(
        new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"`)
      );
      if (relMatch) {
        const target = relMatch[1];
        if (target.startsWith('/xl/')) sheetPath = target.slice(1);
        else if (target.startsWith('worksheets/')) sheetPath = `xl/${target}`;
        else sheetPath = `xl/worksheets/${target}`;
      }
    }
  }

  // ── 4. 读取 XML ──
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`找不到工作表：${sheetPath}`);
  let sheetXml = await sheetFile.async('text');

  const ssFile = zip.file('xl/sharedStrings.xml');
  if (!ssFile) throw new Error('找不到 sharedStrings.xml');
  let ssXml = await ssFile.async('text');

  // ── 5. 写入字符串到单元格的统一辅助函数 ──
  const writeStr = (cellRef: string, value: string, styleId?: number) => {
    if (!value && value !== '0') return;
    const r = upsertSharedString(ssXml, value);
    ssXml = r.newSsXml;
    sheetXml = patchCellStrWithStyle(sheetXml, cellRef, r.index, styleId);
  };

  // ── 6. 填写头部区域 ──
  // 企业固定信息默认兜底：公司、部门、考核人、考核人部门、考核人岗位均固定为图示标准内容
  const finalCompany = (company || '').trim() || DEFAULT_FORMAL_COMPANY;
  const finalDepartment = (department || '').trim() || DEFAULT_FORMAL_DEPARTMENT;
  const finalEvaluator = (evaluator || '').trim() || DEFAULT_FORMAL_EVALUATOR;
  const finalEvaluatorDept = (evaluatorDepartment || '').trim() || DEFAULT_FORMAL_EVALUATOR_DEPARTMENT;
  const finalEvaluatorPos = (evaluatorPosition || '').trim() || DEFAULT_FORMAL_EVALUATOR_POSITION;

  // 动态字段：被考核人姓名与岗位由生成或用户输入确定
  const finalName = (name || '').trim() || '杨祝翔';
  const finalPosition = (position || '').trim() || 'APP开发工程师';

  // A1：在 sharedStrings 第 0 条中直接替换月份文字（不改 XML 结构与居中样式）
  {
    const siList = [...ssXml.matchAll(/<si>[\s\S]*?<\/si>/g)];
    if (siList.length > 0) {
      const oldSi = siList[0][0];
      const newSi = oldSi.replace(
        /(<t[^>]*>)([\s\S]*?)(<\/t>)/,
        (_m, open, text, close) => {
          const newText = text.replace(/^.*?月/, `${month}月`);
          return `${open}${newText}${close}`;
        }
      );
      ssXml = ssXml.replace(oldSi, newSi);
    }
  }

  // 写入并显式绑定对应居中对齐样式（与模板图示严格一致）：
  // D2 (所属公司): 样式 276 (12pt 宋体 Bold, 居中对齐)
  writeStr('D2', finalCompany, 276);
  // D3 (被考核人): 样式 279 (12pt 宋体 Regular, 居中对齐)
  writeStr('D3', finalName, 279);
  // H3 (所属部门): 样式 279 (12pt 宋体 Regular, 居中对齐)
  writeStr('H3', finalDepartment, 279);
  // L3 (岗位名称): 样式 280 (12pt 宋体 Regular, 居中对齐)
  writeStr('L3', finalPosition, 280);
  // D4 (考核人姓名): 样式 279 (12pt 宋体 Regular, 居中对齐)
  writeStr('D4', finalEvaluator, 279);
  // H4 (考核人部门): 样式 279 (12pt 宋体 Regular, 居中对齐)
  writeStr('H4', finalEvaluatorDept, 279);
  // L4 (考核人岗位): 样式 280 (12pt 宋体 Regular, 居中对齐)
  writeStr('L4', finalEvaluatorPos, 280);

  // D5：考核周期 (样式 278: 12pt 宋体 Bold, 居中对齐)
  {
    const periodStr = `${year} 年   ${month}   月   1 日 至 ${year} 年   ${month}   月 ${lastDay} 日`;
    writeStr('D5', periodStr, 278);
  }

  // ── 7. 处理任务行 ──
  // 规范化任务行样式（全部白底、10pt 宋体 Regular，不加粗，字号与字重完全一致）：
  // - 样式 40：10pt 宋体 Regular，无背景，细边框，居中对齐，支持换行
  // - 样式 52：10pt 宋体 Regular，无背景，细边框，百分比数字格式 (0%)，居中对齐
  // - 样式 74：10pt 宋体 Regular，无背景，细边框，居左对齐，支持换行与垂直居中（用于质量目标与质量标准）
  const STYLE_CENTER = 40;
  const STYLE_PERCENT = 52;
  const STYLE_LEFT = 74;

  const N = tasks.length;

  // 7a. 若任务数 > 4，在第 13 行之后插入 N-4 个新行
  if (N > 4) {
    const insertCount = N - 4;

    // 提取第 13 行的完整 XML 作为模板
    const row13Match = sheetXml.match(/<row r="13"[\s\S]*?<\/row>/);
    if (!row13Match) throw new Error('模板结构异常：找不到第 13 行');
    const row13Xml = row13Match[0];

    // 生成新行 XML（行 14 ~ 13+insertCount），先清空各单元格的值
    const insertedRows: string[] = [];
    for (let k = 0; k < insertCount; k++) {
      const newRn = 14 + k;
      // 替换行号和单元格列引用中的行号
      let newRowXml = row13Xml
        .replace(/r="13"/g, `r="${newRn}"`)
        .replace(/<c r="([A-Z]+)13"/g, `<c r="$1${newRn}"`)
        .replace(/ t="[^"]*"/g, '')
        .replace(/<v>[^<]*<\/v>/g, '')
        .replace(/><\/c>/g, '/>');
      insertedRows.push(newRowXml);
    }

    // 将行 14 及以上的行号全部向后偏移 insertCount（先处理 mergeCell，再处理行）
    sheetXml = sheetXml.replace(
      /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g,
      (match, c1, r1s, c2, r2s) => {
        const r1 = parseInt(r1s, 10);
        const r2 = parseInt(r2s, 10);
        const nr1 = r1 >= 14 ? r1 + insertCount : r1;
        const nr2 = r2 >= 14 ? r2 + insertCount : r2;
        return `<mergeCell ref="${c1}${nr1}:${c2}${nr2}"`;
      }
    );
    // 先处理单元格引用（避免被行号替换误伤）
    sheetXml = sheetXml.replace(/<c r="([A-Z]+)(\d+)"/g, (match, col, rns) => {
      const rn = parseInt(rns, 10);
      return rn >= 14 ? `<c r="${col}${rn + insertCount}"` : match;
    });
    // 再处理行引用
    sheetXml = sheetXml.replace(/<row r="(\d+)"/g, (match, rns) => {
      const rn = parseInt(rns, 10);
      return rn >= 14 ? `<row r="${rn + insertCount}"` : match;
    });

    // 在第 13 行 </row> 之后插入新行
    sheetXml = sheetXml.replace(
      /(<row r="13"[\s\S]*?<\/row>)/,
      `$1${insertedRows.join('')}`
    );
  }

  // 7b. 写入各行任务数据（强制统一字号为 10pt、字重为 Regular，质量目标/质量标准统一居左）
  for (let i = 0; i < N; i++) {
    const rn = 10 + i;
    const task = tasks[i];

    // 序号（A列，数字，居中）
    sheetXml = patchCellNumWithStyle(sheetXml, `A${rn}`, i + 1, STYLE_CENTER);

    // 指标类型（B列，居中）
    { const r = upsertSharedString(ssXml, task.type || 'KPI');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `B${rn}`, r.index, STYLE_CENTER); }

    // 指标等级（C列，居中）
    { const r = upsertSharedString(ssXml, task.level || '重要关键任务');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `C${rn}`, r.index, STYLE_CENTER); }

    // 权重（D列）：扣分项使用文本居中样式，数值使用百分比格式样式
    const isDeduction =
      task.weight === '扣分项' ||
      (typeof task.weight === 'string' && task.weight.includes('扣分')) ||
      task.category === '市场侧临时新增开发任务' ||
      (task.description && task.description.includes('市场侧临时新增'));

    if (isDeduction) {
      const r = upsertSharedString(ssXml, '扣分项');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `D${rn}`, r.index, STYLE_CENTER);
    } else {
      let weightNum = 0.25;
      if (typeof task.weight === 'number') {
        weightNum = task.weight > 1 ? task.weight / 100 : task.weight;
      } else if (typeof task.weight === 'string') {
        const clean = (task.weight as string).replace('%', '').trim();
        const p = parseFloat(clean);
        if (!isNaN(p)) weightNum = p > 1 ? p / 100 : p;
      }
      sheetXml = patchCellNumWithStyle(sheetXml, `D${rn}`, weightNum, STYLE_PERCENT);
    }

    // 所属板块（E列，居中）
    { const r = upsertSharedString(ssXml, task.category || '/');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `E${rn}`, r.index, STYLE_CENTER); }

    // 解释说明（F列，居左对齐，便于阅读）
    { const r = upsertSharedString(ssXml, task.description || '');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `F${rn}`, r.index, STYLE_LEFT); }

    // 时间目标（G列，居中）
    { const r = upsertSharedString(ssXml, task.time_target || '');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `G${rn}`, r.index, STYLE_CENTER); }

    // 数量目标（H列，居中）
    { const r = upsertSharedString(ssXml, task.count_target || '/');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `H${rn}`, r.index, STYLE_CENTER); }

    // ★ 质量目标（I列）：严格居左对齐，10pt 宋体 Regular，不加粗
    { const r = upsertSharedString(ssXml, task.quality_target || '');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `I${rn}`, r.index, STYLE_LEFT); }

    // 时间标准（J列，居中）
    { const r = upsertSharedString(ssXml, task.time_standard || '');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `J${rn}`, r.index, STYLE_CENTER); }

    // 数量标准（K列，居中）
    { const r = upsertSharedString(ssXml, task.count_standard || '/');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `K${rn}`, r.index, STYLE_CENTER); }

    // ★ 质量标准（L列）：严格居左对齐，10pt 宋体 Regular，不加粗
    { const r = upsertSharedString(ssXml, task.quality_standard || '');
      ssXml = r.newSsXml;
      sheetXml = patchCellStrWithStyle(sheetXml, `L${rn}`, r.index, STYLE_LEFT); }
  }

  // 7c. 若任务数 < 4，清空剩余行（保留结构与样式）
  if (N < 4) {
    for (let r = 10 + N; r <= 13; r++) {
      for (const col of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']) {
        sheetXml = patchCellEmpty(sheetXml, `${col}${r}`, STYLE_CENTER);
      }
    }
  }

  // ── 8. 写回 ZIP，生成 ArrayBuffer ──
  zip.file(sheetPath, sheetXml);
  zip.file('xl/sharedStrings.xml', ssXml);

  const outputBuffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
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

  const sheetName = '绩效计划表（基层员工）';
  const worksheet = workbook.getWorksheet(sheetName) || workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('未能在导入的 Excel 中找到有效的工作表 (Worksheet)。');
  }

  // 读取头部字段：所属公司(D2)、被考核人(D3)、所属部门(H3)、岗位(L3)、考核人(D4)、考核人部门(H4)、考核人岗位(L4)
  const company = String(worksheet.getCell('D2').value || '').trim();
  const name = String(worksheet.getCell('D3').value || '').trim();
  const department = String(worksheet.getCell('H3').value || '').trim();
  const position = String(worksheet.getCell('L3').value || '').trim();
  const evaluator = String(worksheet.getCell('D4').value || '').trim();
  const evaluatorDepartment = String(worksheet.getCell('H4').value || '').trim();
  const evaluatorPosition = String(worksheet.getCell('L4').value || '').trim();

  const tasks: PerformanceTask[] = [];
  let rowNum = 10;

  while (true) {
    const row = worksheet.getRow(rowNum);
    const seqVal = row.getCell(1).value;
    const typeVal = String(row.getCell(2).value || '').trim();

    // 如果序号为空，或者指标类型为"固定项"，说明工作考核项已读完，跳出循环
    if (seqVal === null || seqVal === undefined || seqVal === '' || typeVal === '固定项') {
      break;
    }

    // 尝试读取权重值并统一化
    let weightVal: number | string = 0.25;
    const rawWeight = row.getCell(4).value;
    if (typeof rawWeight === 'string') {
      const clean = rawWeight.trim();
      if (clean === '扣分项' || clean.includes('扣分')) {
        weightVal = '扣分项';
      } else {
        const parsed = parseFloat(clean.replace('%', ''));
        if (!isNaN(parsed)) {
          weightVal = parsed > 1 ? parsed / 100 : parsed;
        }
      }
    } else if (typeof rawWeight === 'number') {
      weightVal = rawWeight > 1 ? rawWeight / 100 : rawWeight;
    } else if (rawWeight && typeof rawWeight === 'object' && 'result' in rawWeight) {
      // 针对有公式计算的情况，如果是 ExcelJS 包含的计算公式结果
      const res = (rawWeight as any).result;
      if (typeof res === 'number') weightVal = res;
    }

    tasks.push({
      type: typeVal,
      level: String(row.getCell(3).value || '').trim(),
      weight: weightVal,
      category: String(row.getCell(5).value || '').trim(),
      description: String(row.getCell(6).value || '').trim(),
      time_target: String(row.getCell(7).value || '').trim(),
      count_target: String(row.getCell(8).value || '').trim(),
      quality_target: String(row.getCell(9).value || '').trim(),
      time_standard: String(row.getCell(10).value || '').trim(),
      count_standard: String(row.getCell(11).value || '').trim(),
      quality_standard: String(row.getCell(12).value || '').trim(),
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
    tasks
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
    const categoryType = String(row.getCell(1).value || '').trim();
    const qualityTarget = String(row.getCell(2).value || '').trim();
    const qualityStandard = String(row.getCell(3).value || '').trim();

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
