/**
 * @fileoverview 全局状态管理模块。
 * 基于 Vue 3 的 reactive 与 ref，实现轻量级、无依赖的状态共享与持久化。
 * 存储 API 配置、Excel 模板、字段映射关系以及批量员工绩效生成的状态。
 * 遵循 Google TypeScript 编码标准。
 */

import { reactive, ref, watch } from 'vue';
import type { CellMapping, PerformanceTask, QualityStandardItem } from '../utils/excelHelper';
import {
  DEFAULT_FORMAL_COMPANY,
  DEFAULT_FORMAL_DEPARTMENT,
  DEFAULT_FORMAL_EVALUATOR,
  DEFAULT_FORMAL_EVALUATOR_DEPARTMENT,
  DEFAULT_FORMAL_EVALUATOR_POSITION
} from '../utils/excelHelper';

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

/** 支持的主流大模型服务商标识 */
export type LLMProvider = 'gemini' | 'deepseek' | 'qwen' | 'openai' | 'zhipu' | 'custom';

/**
 * 服务商预置配置项定义
 */
export interface ProviderPreset {
  id: LLMProvider;
  name: string;
  badge: string;
  defaultBaseUrl: string;
  defaultModel: string;
  presetModels: string[];
  keyPlaceholder: string;
  keyHelpText: string;
  keyUrl: string;
  isDirectGemini?: boolean;
}

/**
 * 常见主流大模型预置字典
 */
export const LLM_PROVIDERS: Record<LLMProvider, ProviderPreset> = {
  gemini: {
    id: 'gemini',
    name: 'Google Gemini',
    badge: '官方直连',
    defaultBaseUrl: 'https://generativelanguage.googleapis.com',
    defaultModel: 'gemini-2.5-flash',
    presetModels: ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro'],
    keyPlaceholder: 'AIzaSy...',
    keyHelpText: '支持浏览器直连免跨域，获取 API Key：',
    keyUrl: 'https://aistudio.google.com/app/apikey',
    isDirectGemini: true,
  },
  deepseek: {
    id: 'deepseek',
    name: 'DeepSeek',
    badge: '深度求索',
    defaultBaseUrl: 'https://api.deepseek.com/v1',
    defaultModel: 'deepseek-chat',
    presetModels: ['deepseek-chat', 'deepseek-reasoner'],
    keyPlaceholder: 'sk-...',
    keyHelpText: '支持 DeepSeek-V3 与 R1 深度思考推理，获取 API Key：',
    keyUrl: 'https://platform.deepseek.com/api_keys',
  },
  qwen: {
    id: 'qwen',
    name: '通义千问',
    badge: '阿里百炼',
    defaultBaseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    defaultModel: 'qwen-plus',
    presetModels: ['qwen-plus', 'qwen-turbo', 'qwen-max', 'qwen-long', 'qwen2.5-72b-instruct'],
    keyPlaceholder: 'sk-...',
    keyHelpText: '阿里云百炼兼容模式，获取 API Key：',
    keyUrl: 'https://bailian.console.aliyun.com/',
  },
  openai: {
    id: 'openai',
    name: 'ChatGPT',
    badge: 'OpenAI',
    defaultBaseUrl: 'https://api.openai.com/v1',
    defaultModel: 'gpt-4o',
    presetModels: ['gpt-4o', 'gpt-4o-mini', 'gpt-4.5-preview', 'o3-mini', 'gpt-4-turbo'],
    keyPlaceholder: 'sk-...',
    keyHelpText: '官方 OpenAI 或第三方反代网关，获取 API Key：',
    keyUrl: 'https://platform.openai.com/api-keys',
  },
  zhipu: {
    id: 'zhipu',
    name: '智谱 AI',
    badge: 'GLM大模型',
    defaultBaseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    defaultModel: 'glm-4-flash',
    presetModels: ['glm-4-flash', 'glm-4-plus', 'glm-4-air', 'glm-4'],
    keyPlaceholder: '请输入智谱 API Key',
    keyHelpText: 'glm-4-flash 高性价比推荐，获取 API Key：',
    keyUrl: 'https://bigmodel.cn/usercenter/apikeys',
  },
  custom: {
    id: 'custom',
    name: '自定义大模型',
    badge: '自定义 URL',
    defaultBaseUrl: '',
    defaultModel: '',
    presetModels: [],
    keyPlaceholder: '请输入 API Key（可为空）',
    keyHelpText: '支持 OneAPI、NewAPI、Ollama、vLLM 或私有部署端点',
    keyUrl: '',
  },
};

/**
 * 全局 AI 配置数据接口
 */
export interface AIConfig {
  /** 当前选中的服务商类型 */
  provider: LLMProvider;
  /** 当前 API 密钥 */
  apiKey: string;
  /** 接口基础 URL (Base URL) */
  baseUrl: string;
  /** 备用/代理网关地址（保持与 baseUrl 互通，向后兼容） */
  proxyUrl: string;
  /** 调用的模型名称 */
  model: string;
  /** 全局系统提示词 */
  systemInstruction: string;
  /** 针对各服务商独立记忆保存的 API Key 映射字典 */
  providerKeys: Record<string, string>;
  /** 针对各服务商独立记忆保存的 Base URL 映射字典 */
  providerUrls: Record<string, string>;
}

// 1. API 配置状态（从 localStorage 初始化或使用默认值）
const savedConfigRaw = localStorage.getItem(STORAGE_KEYS.GEMINI_CONFIG);
let parsedConfig: Partial<AIConfig> = {};
if (savedConfigRaw) {
  try {
    parsedConfig = JSON.parse(savedConfigRaw);
  } catch (e) {
    console.error('解析已保存的 API 配置失败:', e);
  }
}

const initialProvider: LLMProvider = (parsedConfig.provider as LLMProvider) || 'gemini';
const defaultPreset = LLM_PROVIDERS[initialProvider] || LLM_PROVIDERS.gemini;

export const geminiConfig = reactive<AIConfig>({
  provider: initialProvider,
  apiKey: parsedConfig.apiKey || '',
  baseUrl: parsedConfig.baseUrl || parsedConfig.proxyUrl || defaultPreset.defaultBaseUrl,
  proxyUrl: parsedConfig.proxyUrl || parsedConfig.baseUrl || defaultPreset.defaultBaseUrl,
  model: parsedConfig.model || defaultPreset.defaultModel,
  systemInstruction:
    parsedConfig.systemInstruction ||
    '你是一个专业、严谨且富有建设性眼光的企业HR绩效评定官。请基于员工的工作汇报，产出符合规范的评分与总结。',
  providerKeys: parsedConfig.providerKeys || {},
  providerUrls: parsedConfig.providerUrls || {},
});

// 向后兼容别名
export const aiConfig = geminiConfig;

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

/** 预置默认模板路径与名称常量 */
export const DEFAULT_TEMPLATE_NAME = '正式员工绩效表单套表-JX1.7.xlsx';
export const DEFAULT_TEMPLATE_PATH = '/正式员工绩效表单套表-JX1.7.xlsx';

/** 绩效质量标准库路径与名称常量 */
export const DEFAULT_STANDARDS_NAME = '绩效质量标准库.xlsx';
export const DEFAULT_STANDARDS_PATH = '/绩效质量标准库.xlsx';

// 3. 上传或预置的 Excel 模板状态
export const excelTemplate = reactive<{
  fileName: string;
  buffer: ArrayBuffer | null;
  isCustom: boolean;
}>({
  fileName: '',
  buffer: null,
  isCustom: false,
});

// 3.1 绩效质量标准库状态
export const qualityStandards = ref<QualityStandardItem[]>([]);
export const qualityStandardsLoaded = ref<boolean>(false);

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
  /** 所属公司 (回写到 D2) */
  company?: string;
  /** 员工姓名 (回写到 D3) */
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
 * 添加一位空白的正式员工（预置图示标准固定的公司、部门、考核人及岗位信息）
 */
export function addFormalEmployee(): void {
  formalEmployees.value.push({
    id: crypto.randomUUID(),
    company: DEFAULT_FORMAL_COMPANY,
    name: '杨祝翔',
    department: DEFAULT_FORMAL_DEPARTMENT,
    position: 'APP开发工程师',
    evaluator: DEFAULT_FORMAL_EVALUATOR,
    evaluatorDepartment: DEFAULT_FORMAL_EVALUATOR_DEPARTMENT,
    evaluatorPosition: DEFAULT_FORMAL_EVALUATOR_POSITION,
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
