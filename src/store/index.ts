/**
 * @fileoverview 全局状态管理模块。
 * 基于 Vue 3 的 reactive 与 ref，实现轻量级、无依赖的状态共享与持久化。
 * 存储 API 配置、Excel 模板、字段映射关系以及批量员工绩效生成的状态。
 * 遵循 Google TypeScript 编码标准。
 */

import { reactive, ref, watch } from 'vue';
import type { CellMapping, PerformanceTask } from '../utils/excelHelper';

/**
 * 员工绩效行数据结构
 */
export interface EmployeeRow {
  /** 唯一标识 */
  id: string;
  /** 用户输入的单元格数据，键为单元格坐标，值为输入内容，如 { "B3": "张三" } */
  inputs: Record<string, string | number>;
  /** AI 生成的单元格数据，键为单元格坐标，如 { "C8": 95, "A10": "表现优秀" } */
  aiOutputs: Record<string, string | number>;
  /** 当前行生成状态：idle-闲置，generating-生成中，success-成功，error-失败 */
  status: 'idle' | 'generating' | 'success' | 'error';
  /** 发生错误时的异常描述 */
  errorMessage?: string;
  /** 生成的个人 Excel 文件名 */
  fileName?: string;
  /** 生成的个人 Excel 文件二进制缓存 */
  outputBuffer?: ArrayBuffer | null;
}

// 缓存键名常量
const STORAGE_KEYS = {
  GEMINI_CONFIG: 'ai_performance_gemini_config',
  CELL_MAPPINGS: 'ai_performance_cell_mappings',
  PERFORMANCE_MONTH: 'ai_performance_month',
  AVAILABLE_MODELS: 'ai_performance_available_models',
};

/**
 * 获取本地系统的当前月份，格式为 YYYY-MM。
 * 采用本地时间 API 以防范 UTC 时区导致的前后天月份偏差问题。
 * @returns {string} 本地当前年月字符串
 */
const initMonth = (): string => {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
};

// 全局响应式考核月份变量，自 localStorage 初始化，若无则默认为当前月
const savedMonth = localStorage.getItem(STORAGE_KEYS.PERFORMANCE_MONTH);
export const performanceMonth = ref(savedMonth || initMonth());

// 监听月份变化并将其同步保存至本地缓存中
watch(performanceMonth, (newVal) => {
  localStorage.setItem(STORAGE_KEYS.PERFORMANCE_MONTH, newVal);
});

// 可用的 API 模型列表，从本地 localStorage 初始化
const savedModels = localStorage.getItem(STORAGE_KEYS.AVAILABLE_MODELS);
export const availableModels = ref<string[]>(savedModels ? JSON.parse(savedModels) : []);

// 监听模型列表变化并保存至本地缓存
watch(
  availableModels,
  (newModels) => {
    localStorage.setItem(STORAGE_KEYS.AVAILABLE_MODELS, JSON.stringify(newModels));
  },
  { deep: true }
);

// 1. API 配置状态（从 localStorage 初始化或使用默认值）
const savedConfig = localStorage.getItem(STORAGE_KEYS.GEMINI_CONFIG);
export const geminiConfig = reactive({
  apiKey: '',
  proxyUrl: '',
  model: 'gemini-2.5-flash',
  systemInstruction: '你是一个专业、严谨且富有建设性眼光的企业HR绩效评定官。请基于员工的工作汇报，产出符合规范的评分与总结。',
  ...(savedConfig ? JSON.parse(savedConfig) : {}),
});

// 监听 API 配置变化并持久化
watch(
  () => ({ ...geminiConfig }),
  (newConfig) => {
    localStorage.setItem(STORAGE_KEYS.GEMINI_CONFIG, JSON.stringify(newConfig));
  },
  { deep: true }
);

// 2. 单元格映射关系状态（从 localStorage 初始化或使用默认值）
const savedMappings = localStorage.getItem(STORAGE_KEYS.CELL_MAPPINGS);
export const cellMappings = ref<CellMapping[]>(
  savedMappings ? JSON.parse(savedMappings) : [
    { cellRef: 'B3', fieldName: 'name', label: '员工姓名', type: 'input' },
    { cellRef: 'B4', fieldName: 'department', label: '所属部门', type: 'input' },
    { cellRef: 'A6', fieldName: 'work_summary', label: '本月工作总结', type: 'input' },
    { cellRef: 'C8', fieldName: 'quality_score', label: '工作质量评分 (1-100)', type: 'ai', aiInstruction: '根据工作总结，给出1-100的工作质量数字评分，必须只输出纯数字' },
    { cellRef: 'D8', fieldName: 'efficiency_score', label: '工作效率评分 (1-100)', type: 'ai', aiInstruction: '根据工作总结，给出1-100的工作效率数字评分，必须只输出纯数字' },
    { cellRef: 'A10', fieldName: 'hr_review', label: 'HR详细考核评语', type: 'ai', aiInstruction: '针对该员工的工作总结，生成一段150字以内的、专业的、客观中立的HR绩效考核评语' }
  ]
);

// 监听映射变化并持久化
watch(
  cellMappings,
  (newMappings) => {
    localStorage.setItem(STORAGE_KEYS.CELL_MAPPINGS, JSON.stringify(newMappings));
  },
  { deep: true }
);

// 3. 上传的 Excel 模板状态
export const excelTemplate = reactive<{
  fileName: string;
  buffer: ArrayBuffer | null;
}>({
  fileName: '',
  buffer: null,
});

// 4. 员工列表数据
export const employees = ref<EmployeeRow[]>([]);

/**
 * 添加一个空白的员工数据行
 */
export function addEmployee(): void {
  const newRow: EmployeeRow = {
    id: crypto.randomUUID(),
    inputs: {},
    aiOutputs: {},
    status: 'idle',
  };
  
  // 初始化输入项对应的键值为空字符串
  cellMappings.value.forEach(m => {
    if (m.type === 'input') {
      newRow.inputs[m.cellRef] = '';
    } else {
      newRow.aiOutputs[m.cellRef] = '';
    }
  });

  employees.value.push(newRow);
}

/**
 * 移除指定员工数据行
 * @param id 员工行唯一标识
 */
export function removeEmployee(id: string): void {
  employees.value = employees.value.filter(emp => emp.id !== id);
}

/**
 * 清空所有员工数据
 */
export function clearEmployees(): void {
  employees.value = [];
}

/**
 * 重置所有员工的生成状态为待生成
 */
export function resetGenerationStatus(): void {
  employees.value.forEach(emp => {
    emp.status = 'idle';
    emp.errorMessage = undefined;
    emp.outputBuffer = null;
    emp.fileName = undefined;
    // 清空 AI 的输出值
    cellMappings.value.forEach(m => {
      if (m.type === 'ai') {
        emp.aiOutputs[m.cellRef] = '';
      }
    });
  });
}

// ==========================================
// 5. 正式员工特定模板数据状态与操作
// ==========================================

/**
 * 正式员工行数据结构
 */
export interface FormalEmployeeRow {
  /** 唯一标识 */
  id: string;
  /** 员工姓名 (回写到 D3) */
  name: string;
  /** 岗位名称 (回写到 L3) */
  position: string;
  /** 上个月绩效回顾 */
  lastMonthPerformance: string;
  /** 本月工作内容计划 */
  thisMonthWorkContent: string;
  /** AI 生成的计划任务列表 */
  tasks: PerformanceTask[];
  /** 当前行生成状态：idle-闲置，generating-生成中，success-成功，error-失败 */
  status: 'idle' | 'generating' | 'success' | 'error';
  /** 发生错误时的异常描述 */
  errorMessage?: string;
  /** 生成的个人 Excel 文件名 */
  fileName?: string;
  /** 生成的个人 Excel 文件二进制缓存 */
  outputBuffer?: ArrayBuffer | null;
}

// 缓存键名
const STORAGE_KEYS_FORMAL = {
  FORMAL_EMPLOYEES: 'ai_performance_formal_employees',
};

const savedFormal = localStorage.getItem(STORAGE_KEYS_FORMAL.FORMAL_EMPLOYEES);
export const formalEmployees = ref<FormalEmployeeRow[]>(
  savedFormal ? JSON.parse(savedFormal) : []
);

// 自动持久化正式员工数据
watch(
  formalEmployees,
  (newFormal) => {
    // 过滤掉二进制 buffer 缓存，避免超出 localStorage 空间限制
    const stripped = newFormal.map(emp => ({
      ...emp,
      outputBuffer: null // 不持久化 buffer
    }));
    localStorage.setItem(STORAGE_KEYS_FORMAL.FORMAL_EMPLOYEES, JSON.stringify(stripped));
  },
  { deep: true }
);

/**
 * 添加一位空白的正式员工
 */
export function addFormalEmployee(): void {
  formalEmployees.value.push({
    id: crypto.randomUUID(),
    name: '新员工',
    position: 'APP开发工程师',
    lastMonthPerformance: '',
    thisMonthWorkContent: '',
    tasks: [],
    status: 'idle',
  });
}

/**
 * 移除指定正式员工
 * @param id 唯一标识
 */
export function removeFormalEmployee(id: string): void {
  formalEmployees.value = formalEmployees.value.filter(emp => emp.id !== id);
}

/**
 * 清空所有正式员工
 */
export function clearFormalEmployees(): void {
  formalEmployees.value = [];
}

/**
 * 重置所有正式员工的生成状态
 */
export function resetFormalGenerationStatus(): void {
  formalEmployees.value.forEach(emp => {
    emp.status = 'idle';
    emp.errorMessage = undefined;
    emp.outputBuffer = null;
    emp.fileName = undefined;
    emp.tasks = [];
  });
}
