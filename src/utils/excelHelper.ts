/**
 * @fileoverview Excel 模板读写辅助工具函数。
 * 包含解析 Excel 模板文件、往指定单元格回写数据、保留样式并导出等功能。
 * 遵循 Google TypeScript 编码标准。
 */

import ExcelJS from 'exceljs';

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
  weight: number;      // 权重，如 0.3
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
 * 将正式员工月度绩效任务写入特定模板
 * 仅修改和填写红色背景的部分（被考核人D3、岗位L3以及第10行起的任务明细）
 * 当任务数大于4时，动态在14行起插行并拷贝样式；若不大于4则覆盖10-13行
 *
 * @param templateBuffer 原始模板文件的 ArrayBuffer
 * @param name 员工姓名 (回写到 D3)
 * @param position 岗位名称 (回写到 L3)
 * @param tasks 任务数组 (回写到 10 行起)
 * @returns Promise<ArrayBuffer> 修改后的 Excel 文件二进制 buffer
 */
export async function writePerformanceToTemplate(
  templateBuffer: ArrayBuffer,
  name: string,
  position: string,
  tasks: PerformanceTask[]
): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer.slice(0));

  const sheetName = '绩效计划表（基层员工）';
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`未能在模板中找到 "${sheetName}" 工作表。`);
  }

  // 1. 写入姓名与岗位 (D3, L3) 并清除红色背景，同时强制设为水平与垂直双向居中对齐
  const cellD3 = worksheet.getCell('D3');
  cellD3.value = name;
  cellD3.fill = { type: 'pattern', pattern: 'none' };
  cellD3.alignment = { horizontal: 'center', vertical: 'middle' };

  const cellL3 = worksheet.getCell('L3');
  cellL3.value = position;
  cellL3.fill = { type: 'pattern', pattern: 'none' };
  cellL3.alignment = { horizontal: 'center', vertical: 'middle' };

  const N = tasks.length;

  // 2. 如果生成任务数大于 4 行，我们需要从第 14 行开始插入 N - 4 行，防止覆盖下方固定项
  if (N > 4) {
    const insertCount = N - 4;
    // 插入空行
    worksheet.spliceRows(14, 0, ...Array(insertCount).fill([]));

    // 复制第 10 行的样式与行高
    for (let i = 0; i < insertCount; i++) {
      const targetRowNum = 14 + i;
      const srcRow = worksheet.getRow(10);
      const destRow = worksheet.getRow(targetRowNum);
      destRow.height = srcRow.height;

      for (let col = 1; col <= 12; col++) {
        const srcCell = srcRow.getCell(col);
        const destCell = destRow.getCell(col);
        
        // 复制样式
        if (srcCell.font) destCell.font = { ...srcCell.font };
        if (srcCell.fill) destCell.fill = { ...srcCell.fill };
        if (srcCell.border) destCell.border = { ...srcCell.border };
        if (srcCell.alignment) destCell.alignment = { ...srcCell.alignment };
        if (srcCell.numFmt) destCell.numFmt = srcCell.numFmt;
      }
      destRow.commit();
    }
  }

  // 3. 填入任务明细数据，并清除任务行的红色背景填充
  for (let i = 0; i < N; i++) {
    const rowNum = 10 + i;
    const task = tasks[i];
    const row = worksheet.getRow(rowNum);

    // 清除整行 12 列单元格的背景颜色填充
    for (let col = 1; col <= 12; col++) {
      row.getCell(col).fill = { type: 'pattern', pattern: 'none' };
    }

    row.getCell(1).value = i + 1; // 序号
    row.getCell(2).value = task.type || 'KPI';
    row.getCell(3).value = task.level || '重要关键任务';

    // 转换权重：若是带%字符的，转为小数；若已经大于1，除以100；否则直接转浮点
    let weightVal = 0.25;
    if (typeof task.weight === 'string') {
      const cleanWeight = (task.weight as string).replace('%', '').trim();
      weightVal = parseFloat(cleanWeight);
      if (weightVal > 1) {
        weightVal = weightVal / 100;
      }
    } else if (typeof task.weight === 'number') {
      weightVal = task.weight;
      if (weightVal > 1) {
        weightVal = weightVal / 100;
      }
    }
    row.getCell(4).value = weightVal;

    row.getCell(5).value = task.category || '';
    row.getCell(6).value = task.description || '';
    row.getCell(7).value = task.time_target || '';
    row.getCell(8).value = task.count_target || '/';
    row.getCell(9).value = task.quality_target || '';
    row.getCell(10).value = task.time_standard || '';
    row.getCell(11).value = task.count_standard || '/';
    row.getCell(12).value = task.quality_standard || '';

    // 设置单元格对齐与排版样式：
    // - 质量目标（第 9 列）与质量标准（第 12 列）设为垂直居中、水平靠左（靠前）对齐，并开启自动换行（wrapText）以确保多条描述单独成行排版
    // - 其它所有单元格（序号、类型、等级、权重、所属板块、任务说明、时间、数量等）一律强制设为水平、垂直双向居中对齐（对于可能存在换行倾向的列 6 与列 10 也开启 wrapText 支持）
    for (let col = 1; col <= 12; col++) {
      const cell = row.getCell(col);
      if (col === 9 || col === 12) {
        cell.alignment = {
          ...(cell.alignment || {}),
          wrapText: true,
          vertical: 'middle',
          horizontal: 'left',
        };
      } else {
        cell.alignment = {
          ...(cell.alignment || {}),
          wrapText: col === 6 || col === 10,
          vertical: 'middle',
          horizontal: 'center',
        };
      }
    }

    row.commit();
  }

  // 4. 保存为二进制 buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}

/**
 * 从已填写的正式员工绩效 Excel 中读取姓名、岗位及工作考核任务项
 *
 * @param file 历史绩效 Excel 文件
 * @returns Promise<{ name: string; position: string; tasks: PerformanceTask[] }>
 */
export async function readPerformanceFromExcel(file: File): Promise<{
  name: string;
  position: string;
  tasks: PerformanceTask[];
}> {
  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();
  await workbook.xlsx.load(arrayBuffer);

  const sheetName = '绩效计划表（基层员工）';
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`未能在导入的 Excel 中找到 "${sheetName}" 工作表。`);
  }

  // 读取姓名 (D3) 与岗位 (L3)
  const name = String(worksheet.getCell('D3').value || '').trim();
  const position = String(worksheet.getCell('L3').value || '').trim();

  const tasks: PerformanceTask[] = [];
  let rowNum = 10;

  while (true) {
    const row = worksheet.getRow(rowNum);
    const seqVal = row.getCell(1).value;
    const typeVal = String(row.getCell(2).value || '').trim();

    // 如果序号为空，或者指标类型为“固定项”，说明工作考核项已读完，跳出循环
    if (seqVal === null || seqVal === undefined || seqVal === '' || typeVal === '固定项') {
      break;
    }

    // 尝试读取权重值并统一化
    let weightVal = 0.25;
    const rawWeight = row.getCell(4).value;
    if (typeof rawWeight === 'number') {
      weightVal = rawWeight;
    } else if (typeof rawWeight === 'string') {
      weightVal = parseFloat((rawWeight as string).replace('%', '').trim());
      if (weightVal > 1) {
        weightVal = weightVal / 100;
      }
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

  return { name, position, tasks };
}
