<script setup lang="ts">
/**
 * @fileoverview 正式员工绩效表单自动生成组件。
 * 提供宽屏大卡片工作台、隐藏上月绩效详情以精简视觉、全宽大型本月工作安排编辑框、以及 100% 满宽绩效编辑网格。
 * 默认使用本地预置模板，支持随时上传自定义模板，导出时保留 Excel 字体、合并单元格及边框样式。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref, onMounted, onUnmounted, computed, watch } from "vue";
import {
  geminiConfig,
  excelTemplate,
  formalEmployees,
  addFormalEmployee,
  removeFormalEmployee,
  resetFormalGenerationStatus,
  performanceMonth,
  DEFAULT_TEMPLATE_PATH,
  DEFAULT_TEMPLATE_NAME,
  DEFAULT_STANDARDS_PATH,
  qualityStandards,
  defaultQualityStandards,
  customQualityStandards,
  customStandardsFileName,
  setDefaultQualityStandards,
  mergeQualityStandards,
  resetQualityStandardsToDefault,
} from "../store";
import type { FormalEmployeeRow } from "../store";
import { generateFormalPerformance } from "../utils/geminiHelper";
import {
  writePerformanceToTemplate,
  downloadExcelFile,
  readPerformanceFromExcel,
  parseQualityStandards,
} from "../utils/excelHelper";
import type { PerformanceTask } from "../utils/excelHelper";
import {
  Users,
  Trash2,
  Sparkles,
  Download,
  AlertTriangle,
  FileSpreadsheet,
  RefreshCw,
  Upload,
  Plus,
  XCircle,
  RotateCcw,
  X,
  Target,
  CheckCircle2,
  FileUp,
  Search,
  Info,
} from "lucide-vue-next";

// 记录当前正在编辑任务的员工 ID
const activeEmpId = ref<string | null>(null);
// 记录批量生成状态
const isGeneratingBatch = ref(false);
// 加载提示
const templateLoadingStatus = ref("");

// 当前活跃的推导中止控制器（用于支持随时取消与中断重试）
let activeAbortController: AbortController | null = null;
// 已生成耗时秒数（用于倒计时与超时拦截感知）
const generatingElapsedSeconds = ref(0);
let elapsedTimerId: any = null;

// 计算当前选中的正在编辑绩效计划的员工
const activeEmployee = computed(() => {
  return formalEmployees.value.find((emp) => emp.id === activeEmpId.value) || null;
});

// 全局基础校验
const validationError = computed(() => {
  if (!excelTemplate.buffer) {
    return "未加载 Excel 模板，请确保 public 下存在模板或上传新模板。";
  }
  if (!geminiConfig.apiKey) {
    return "请先在上方设置中填写 API Key。";
  }
  if (formalEmployees.value.length === 0) {
    return "请至少添加一名正式员工。";
  }
  return "";
});

// 计算当前正在编辑员工的任务指标规则校验
// 包含五项核心校验规则：
// 1. 各项常规指标权重最大不超过 30%
// 2. 常规任务/辅助任务 (CPI) 单项权重不能超过 20%
// 3. 常规指标总权重刚好等于 100%（扣分项任务免计）
// 4. 重要战略/关键任务对应的类型必须为 KPI，常规/辅助任务对应的类型必须为 CPI
// 5. 重要任务比例最低不能低于 20%
const activeEmpValidationError = computed(() => {
  const emp = activeEmployee.value;
  if (!emp || emp.status !== "success" || !emp.tasks || emp.tasks.length === 0) {
    return "";
  }

  // 1. 校验常规任务权重总和是否为 1.0 (100%)
  const totalW = getTasksTotalWeight(emp.tasks);
  if (Math.abs(totalW - 1.0) > 0.001) {
    return `常规考核任务总权重为 ${(totalW * 100).toFixed(0)}%，必须刚好等于 100%（已自动排除扣分项）。`;
  }

  // 2. 校验指标最大限制与最低比例约束（排除扣分项）
  for (let i = 0; i < emp.tasks.length; i++) {
    const t = emp.tasks[i];
    if (t.weight === "扣分项" || (typeof t.weight === "string" && t.weight.includes("扣分"))) {
      continue; // 扣分项不计入 30% 上限与 20% 下限校验
    }

    const w = parseFloat((t.weight as any) || 0);

    if (w > 0.30001) {
      return `第 ${i + 1} 项“${t.category || "指标"}”权重为 ${(w * 100).toFixed(0)}%，已超过最大限制 30%。`;
    }

    if ((t.level === "常规执行任务" || t.level === "辅助零散任务" || t.type === "CPI") && w > 0.20001) {
      return `常规/辅助任务（CPI，如第 ${i + 1} 项“${t.category || "指标"}”）权重为 ${(w * 100).toFixed(0)}%，不能超过 20%。`;
    }

    if (t.level === "重要关键任务" && w < 0.1999) {
      return `重要关键任务（如第 ${i + 1} 项）的权重为 ${(w * 100).toFixed(0)}%，低于最低占比要求 20%。`;
    }

    if ((t.level === "常规执行任务" || t.level === "辅助零散任务") && t.type !== "CPI") {
      return `常规/辅助任务（如第 ${i + 1} 项）类型当前为 ${t.type}，必须设为 CPI。`;
    }
  }

  return "";
});

/**
 * 组件加载时自动拉取 Public 下的默认 Excel 模板文件与质量标准库
 */
onMounted(async () => {
  // 1. 加载默认 Excel 模板
  if (!excelTemplate.buffer) {
    templateLoadingStatus.value = `正在自动加载预置绩效模板 (${DEFAULT_TEMPLATE_NAME})...`;
    try {
      const response = await fetch(DEFAULT_TEMPLATE_PATH);
      if (response.ok) {
        const buffer = await response.arrayBuffer();
        excelTemplate.buffer = buffer;
        excelTemplate.fileName = DEFAULT_TEMPLATE_NAME;
        excelTemplate.isCustom = false;
        templateLoadingStatus.value = `已载入预置模板: ${DEFAULT_TEMPLATE_NAME}`;
      } else {
        templateLoadingStatus.value = "未找到预置绩效模板，请手动上传模板文件。";
      }
    } catch (e) {
      console.error("加载预置模板失败:", e);
      templateLoadingStatus.value = "加载预置模板失败，请手动上传。";
    }
  } else {
    templateLoadingStatus.value = `已载入模板: ${excelTemplate.fileName}`;
  }

  // 2. 自动加载并解析企业绩效质量标准库
  if (defaultQualityStandards.value.length === 0) {
    try {
      // 携带时间戳参数破坏浏览器 HTTP 缓存，确保模板更新后即刻生效
      const res = await fetch(`${DEFAULT_STANDARDS_PATH}?t=${Date.now()}`);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const list = await parseQualityStandards(buf, false);
        setDefaultQualityStandards(list);
      }
    } catch (err) {
      console.warn("加载绩效质量标准库失败:", err);
    }
  }

  // 默认填充一个空白行
  if (formalEmployees.value.length === 0) {
    addFormalEmployee();
  }

  // 默认将第一行设为编辑态
  if (formalEmployees.value.length > 0) {
    activeEmpId.value = formalEmployees.value[0].id;
  }
});

/**
 * 监听员工手动上传并覆盖模板
 */
async function handleTemplateUpload(event: Event) {
  const target = event.target as HTMLInputElement;
  const files = target.files;
  if (!files || files.length === 0) return;

  const file = files[0];
  try {
    const arrayBuffer = await file.arrayBuffer();
    excelTemplate.fileName = file.name;
    excelTemplate.buffer = arrayBuffer;
    excelTemplate.isCustom = true;
    templateLoadingStatus.value = `已载入自定义模板: ${file.name}`;

    // 如果已有成功生成的员工，重新刷新其导出的 buffer
    formalEmployees.value.forEach((emp) => {
      if (emp.status === "success" && emp.tasks.length > 0) {
        updateExcelBuffer(emp);
      }
    });
  } catch (error: any) {
    console.error("解析模板失败:", error);
    templateLoadingStatus.value = "自定义模板载入失败，请检查文件。";
  } finally {
    target.value = "";
  }
}

/**
 * 重算并更新某位员工导出的 Excel 二进制缓存 (回写并清除红色背景)
 */
async function updateExcelBuffer(emp: FormalEmployeeRow): Promise<boolean> {
  if (!excelTemplate.buffer || emp.tasks.length === 0) return false;

  try {
    const buffer = await writePerformanceToTemplate(excelTemplate.buffer, emp.name, emp.position, emp.tasks, performanceMonth.value, {
      company: emp.company,
      department: emp.department,
      evaluator: emp.evaluator,
      evaluatorDepartment: emp.evaluatorDepartment,
      evaluatorPosition: emp.evaluatorPosition,
    });
    emp.outputBuffer = buffer;
    emp.fileName = `绩效计划表_${emp.name}_${emp.position}_${performanceMonth.value}.xlsx`;
    return true;
  } catch (e: any) {
    console.error("回写 Excel 缓存失败:", e);
    emp.errorMessage = `回写 Excel 失败: ${e.message}`;
    emp.status = "error";
    return false;
  }
}

/**
 * 单个员工绩效智能生成逻辑（支持 60 秒超时与自动/手动重试）
 */
async function generateSingle(emp: FormalEmployeeRow) {
  if (!geminiConfig.apiKey) return;

  // 1. 若存在未结束的前序请求，先行中断，防止并发冲突
  if (activeAbortController) {
    activeAbortController.abort(new Error('用户重新发起了推导，已终止前序请求。'));
    activeAbortController = null;
  }

  // 2. 初始化中止控制器与计时器
  activeAbortController = new AbortController();
  const currentSignal = activeAbortController.signal;

  clearInterval(elapsedTimerId);
  generatingElapsedSeconds.value = 0;
  elapsedTimerId = setInterval(() => {
    generatingElapsedSeconds.value += 1;
  }, 1000);

  emp.status = "generating";
  emp.errorMessage = undefined;
  emp.tasks = [];

  try {
    const result = await generateFormalPerformance(
      {
        provider: geminiConfig.provider,
        apiKey: geminiConfig.apiKey,
        baseUrl: geminiConfig.baseUrl,
        proxyUrl: geminiConfig.proxyUrl,
        model: geminiConfig.model,
        systemInstruction: geminiConfig.systemInstruction,
      },
      emp.name,
      emp.position,
      emp.lastMonthPerformance,
      emp.thisMonthWorkContent,
      performanceMonth.value,
      qualityStandards.value, // 注入绩效质量标准库
      {
        timeoutMs: 60000,
        abortSignal: currentSignal,
      }
    );

    if (result.success) {
      emp.tasks = result.tasks;
      emp.name = result.name;
      emp.position = result.position;

      // 渲染回写并清除红色背景
      const successWrite = await updateExcelBuffer(emp);
      if (successWrite) {
        emp.status = "success";
      }
    } else {
      emp.errorMessage = result.message || "推导失败，请点击重新推导。";
      emp.status = "error";
    }
  } catch (err: any) {
    console.error("生成异常:", err);
    emp.errorMessage = err.message || "大模型响应超时（超过 1 分钟）或网络异常，请点击重新推导。";
    emp.status = "error";
  } finally {
    clearInterval(elapsedTimerId);
    activeAbortController = null;
  }
}

/**
 * 手动取消当前正在生成的推导任务
 */
function cancelGeneration(emp: FormalEmployeeRow) {
  if (activeAbortController) {
    activeAbortController.abort(new Error('用户已手动取消该次推导。'));
    activeAbortController = null;
  }
  clearInterval(elapsedTimerId);
  generatingElapsedSeconds.value = 0;
  emp.status = "idle";
  emp.errorMessage = "已取消本次推导，您可以调整工作内容或配置后重新开始。";
}

/**
 * 取消当前并立即重新发起推导
 */
function cancelAndRetry(emp: FormalEmployeeRow) {
  cancelGeneration(emp);
  generateSingle(emp);
}

/**
 * 触发单个下载
 */
function downloadSingle(emp: FormalEmployeeRow) {
  if (!emp.outputBuffer || !emp.fileName) return;
  downloadExcelFile(emp.outputBuffer, emp.fileName);
}

/**
 * 一键下载所有生成的 Excel
 */
function downloadAll() {
  const successRows = formalEmployees.value.filter((emp) => emp.status === "success" && emp.outputBuffer && emp.fileName);
  successRows.forEach((emp, index) => {
    setTimeout(() => {
      if (emp.outputBuffer && emp.fileName) {
        downloadExcelFile(emp.outputBuffer, emp.fileName);
      }
    }, index * 200);
  });
}

/**
 * 手动新增任务行
 */
function addTask(emp: FormalEmployeeRow) {
  emp.tasks.push({
    type: "KPI",
    level: "重要关键任务",
    weight: 0.25,
    category: "系统开发",
    description: "新增的考核任务",
    time_target: "当月月底前",
    count_target: "/",
    quality_target: "1.按时按质交付开发任务\n2.单元测试通过率达100%",
    time_standard: "每延迟一天扣2分",
    count_standard: "/",
    quality_standard: "1.未按原型要求交付该项扣10分\n2.收到投诉扣2分/次",
  });
  updateExcelBuffer(emp);
}

/**
 * 手动删除任务行
 */
function removeTask(emp: FormalEmployeeRow, index: number) {
  emp.tasks.splice(index, 1);
  updateExcelBuffer(emp);
}

/**
 * 校验当前常规任务的总权重是否为 1.0 (100%)（自动排除“扣分项”等非数值任务）
 */
function getTasksTotalWeight(tasks: PerformanceTask[]): number {
  return tasks
    .filter((t) => t.weight !== "扣分项" && !(typeof t.weight === "string" && t.weight.includes("扣分")))
    .reduce((sum, t) => {
      const val = typeof t.weight === "number" ? t.weight : parseFloat(t.weight);
      return sum + (isNaN(val) ? 0 : val);
    }, 0);
}

/**
 * 表格单元格修改后的回调
 */
function handleCellChange(emp: FormalEmployeeRow) {
  updateExcelBuffer(emp);
}

function handleAddEmployee() {
  addFormalEmployee();
  if (formalEmployees.value.length > 0) {
    activeEmpId.value = formalEmployees.value[formalEmployees.value.length - 1].id;
  }
}

/**
 * 移除当前选中员工
 */
function handleRemoveActive() {
  if (!activeEmpId.value) return;
  const currentId = activeEmpId.value;
  const index = formalEmployees.value.findIndex((emp) => emp.id === currentId);
  removeFormalEmployee(currentId);
  if (formalEmployees.value.length > 0) {
    const nextIndex = Math.max(0, index - 1);
    activeEmpId.value = formalEmployees.value[nextIndex].id;
  } else {
    activeEmpId.value = null;
  }
}

/**
 * 导入并解析员工上个月绩效 Excel 文件
 */
async function handleLastMonthExcelUpload(event: Event, emp: FormalEmployeeRow) {
  const target = event.target as HTMLInputElement;
  const files = target.files;
  if (!files || files.length === 0) return;

  const file = files[0];
  try {
    const data = await readPerformanceFromExcel(file);

    // 自动回填公司、部门、姓名、岗位、考核人等所有头部信息
    if (data.name) emp.name = data.name;
    if (data.position) emp.position = data.position;
    if (data.company) emp.company = data.company;
    if (data.department) emp.department = data.department;
    if (data.evaluator) emp.evaluator = data.evaluator;
    if (data.evaluatorDepartment) emp.evaluatorDepartment = data.evaluatorDepartment;
    if (data.evaluatorPosition) emp.evaluatorPosition = data.evaluatorPosition;

    // 格式化上月指标，用于 Gemini 提示上下文
    const taskSummary = data.tasks
      .map((t: PerformanceTask, idx: number) => {
        const weightStr = typeof t.weight === "number" ? `${(t.weight * 100).toFixed(0)}%` : t.weight || "扣分项";
        return `【任务 ${idx + 1}】
- 类型: ${t.type} (${t.level}) | 权重: ${weightStr} | 板块: ${t.category}
- 任务描述: ${t.description}
- 时间目标: ${t.time_target} | 数量目标: ${t.count_target}
- 质量目标: ${t.quality_target}
- 时间标准: ${t.time_standard} | 数量标准: ${t.count_standard}
- 质量标准: ${t.quality_standard}`;
      })
      .join("\n\n");

    emp.lastMonthPerformance = `[已解析自上月绩效 Excel] ${file.name}
所属公司：${data.company || "无"} | 所属部门：${data.department || "无"}
被考核人：${data.name} | 岗位：${data.position} | 考核人：${data.evaluator || "无"}

上月考核项回顾：
${taskSummary}`;

    // 如果已有生成的任务，重新刷新导出的二进制 buffer
    if (emp.status === "success" && emp.tasks.length > 0) {
      updateExcelBuffer(emp);
    }
  } catch (error: any) {
    console.error("解析上月绩效 Excel 失败:", error);
    alert(`解析失败: ${error.message || "文件格式错误"}`);
  } finally {
    target.value = "";
  }
}

// ==========================================
// 绩效质量标准库上传、模板下载与弹窗状态管理
// ==========================================

// 绩效质量标准库配置弹窗显隐状态
const showStandardsModal = ref(false);
// 标准库实时预览搜索过滤关键词
const standardsSearchQuery = ref("");
// 标准库上传解析错误提示
const standardsUploadError = ref("");
// 是否正在解析标准库文件
const isStandardsUploading = ref(false);

/**
 * 计算符合当前搜索关键词的标准库条目列表
 */
const filteredStandards = computed(() => {
  const query = standardsSearchQuery.value.trim().toLowerCase();
  if (!query) {
    return qualityStandards.value;
  }
  return qualityStandards.value.filter(
    (item) =>
      item.categoryType.toLowerCase().includes(query) ||
      item.qualityTarget.toLowerCase().includes(query) ||
      item.qualityStandard.toLowerCase().includes(query)
  );
});

/**
 * 安全解析并导入用户上传的自定义标准库文件
 *
 * @param file 用户选中的 Excel 文件对象
 */
const processCustomStandardsFile = async (file: File) => {
  // 安全校验：文件大小限制 10MB
  if (file.size > 10 * 1024 * 1024) {
    standardsUploadError.value = "上传的标准库文件不能超过 10MB！";
    return;
  }

  // 安全校验：文件格式限制
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    standardsUploadError.value = "请上传 .xlsx 格式的标准库 Excel 文件！";
    return;
  }

  isStandardsUploading.value = true;
  standardsUploadError.value = "";

  try {
    const arrayBuffer = await file.arrayBuffer();
    const customList = await parseQualityStandards(arrayBuffer, true);
    if (customList.length === 0) {
      standardsUploadError.value =
        "未能从文件中解析出有效的质量标准条目，请确认表头包含[类型/质量目标/质量标准]并参考模板。";
    } else {
      mergeQualityStandards(customList, file.name);
      standardsUploadError.value = "";
    }
  } catch (err: any) {
    console.error("解析自定义质量标准库失败:", err);
    standardsUploadError.value = `解析失败: ${err?.message || "Excel 格式异常，请参照规范模板"}`;
  } finally {
    isStandardsUploading.value = false;
  }
};

/**
 * 处理用户上传自定义绩效质量标准库 Excel 文件
 */
const handleCustomStandardsUpload = async (event: Event) => {
  const target = event.target as HTMLInputElement;
  const file = target.files?.[0];
  if (!file) return;

  await processCustomStandardsFile(file);
  target.value = "";
};

/**
 * 处理拖拽上传质量标准库文件
 */
const handleStandardsDrop = async (event: DragEvent) => {
  event.preventDefault();
  const file = event.dataTransfer?.files?.[0];
  if (!file) return;

  await processCustomStandardsFile(file);
};

/**
 * 下载标准库官方规范模板
 */
const handleDownloadStandardsTemplate = async () => {
  try {
    const res = await fetch(`${DEFAULT_STANDARDS_PATH}?t=${Date.now()}`);
    if (!res.ok) throw new Error("获取标准库模板文件失败");
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "绩效质量标准库_模板.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err: any) {
    console.error("下载模板出错:", err);
    alert(`下载模板失败: ${err?.message || "请稍后重试"}`);
  }
};

/**
 * 恢复为系统预置默认标准库
 */
const handleResetStandardsToDefault = () => {
  resetQualityStandardsToDefault();
  standardsUploadError.value = "";
};

/**
 * 监听标准库弹窗开关，动态控制全局页面滚动锁定 (Scroll Lock)
 * 确保弹窗内部滚动时，外部宿主页面被牢固锁定，禁止任何背景联动与滚动穿透。
 */
watch(showStandardsModal, (isOpen) => {
  if (typeof document !== "undefined") {
    if (isOpen) {
      document.body.classList.add("modal-scroll-locked");
    } else {
      document.body.classList.remove("modal-scroll-locked");
    }
  }
});

/**
 * 全局键盘事件监听：按下 Escape 键时关闭标准库管理弹窗
 */
const handleGlobalKeyDown = (e: KeyboardEvent) => {
  if (e.key === "Escape" && showStandardsModal.value) {
    showStandardsModal.value = false;
  }
};

onMounted(() => {
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", handleGlobalKeyDown);
  }
});

// 组件卸载时释放全局监听及滚动锁，防止页面永久处于锁定状态
onUnmounted(() => {
  if (typeof document !== "undefined") {
    document.body.classList.remove("modal-scroll-locked");
  }
  if (typeof window !== "undefined") {
    window.removeEventListener("keydown", handleGlobalKeyDown);
  }
});
</script>

<template>
  <div class="glass-card formal-generator">
    <!-- 头部栏 -->
    <div class="card-header">
      <FileSpreadsheet class="header-icon" />
      <div class="header-title-group">
        <h2>正式员工绩效智能填报</h2>
        <span class="header-subtitle" v-if="excelTemplate.fileName">({{ excelTemplate.fileName }})</span>
        <button
          type="button"
          class="standards-badge"
          @click="showStandardsModal = true"
          title="点击查看并配置绩效质量标准库（支持上传与模板下载）"
        >
          🎯 质量标准库 ({{ qualityStandards.length }}条规范<span v-if="customQualityStandards.length > 0" class="custom-sub-tag"> · 自定义</span>)
        </button>
      </div>
      <div class="header-actions">
        <button
          type="button"
          class="btn btn-outline btn-sm-action"
          @click="showStandardsModal = true"
          title="配置并管理绩效质量标准库"
        >
          <Target :size="14" />
          <span>管理标准库</span>
        </button>
        <span class="template-status-tag" :class="{ ok: excelTemplate.buffer }">
          {{ excelTemplate.fileName ? `已载入: ${excelTemplate.fileName}` : templateLoadingStatus || "未检测到模板" }}
        </span>
        <label class="btn btn-outline btn-sm-upload">
          <input type="file" accept=".xlsx" class="hidden-input" @change="handleTemplateUpload" />
          <span>{{ excelTemplate.fileName ? "更换 Excel 模板" : "上传 Excel 模板" }}</span>
        </label>
      </div>
    </div>

    <div class="card-body">
      <!-- 错误验证提示 -->
      <div v-if="validationError" class="validation-warning">
        <AlertTriangle :size="20" class="warning-icon" />
        <span>{{ validationError }}</span>
      </div>

      <!-- 员工选择 Tabs 导航 (全屏横排) -->
      <div class="employee-tabs">
        <button v-for="emp in formalEmployees" :key="emp.id" class="emp-tab" :class="{ active: emp.id === activeEmpId, success: emp.status === 'success' }" @click="activeEmpId = emp.id">
          <Users :size="14" class="tab-icon" />
          <span class="emp-tab-name">{{ emp.name || "新员工" }}</span>
          <span class="emp-tab-pos" v-if="emp.position">({{ emp.position }})</span>
          <span class="emp-tab-badge" :class="emp.status"></span>
        </button>
        <button type="button" class="emp-tab add-btn" @click="handleAddEmployee" :disabled="isGeneratingBatch">
          <Plus :size="14" />
          <span>添加新员工</span>
        </button>
      </div>

      <!-- 核心工作台：大型输入配置区 (满宽卡片) -->
      <div class="input-panel-card" v-if="activeEmployee">
        <!-- 员工头部信息与导入配置 -->
        <div class="panel-row">
          <div class="form-group w-20">
            <label>被考核人姓名 (D3)</label>
            <input type="text" v-model="activeEmployee.name" class="form-control" placeholder="例如：杨砚翔" :disabled="isGeneratingBatch" @change="updateExcelBuffer(activeEmployee)" />
          </div>
          <div class="form-group w-20">
            <label>岗位名称 (L3)</label>
            <input type="text" v-model="activeEmployee.position" class="form-control" placeholder="例如：APP开发工程师" :disabled="isGeneratingBatch" @change="updateExcelBuffer(activeEmployee)" />
          </div>
          <div class="form-group w-20">
            <label>所属部门 (H3/H4)</label>
            <input type="text" v-model="activeEmployee.department" class="form-control" placeholder="默认：软件研发部" :disabled="isGeneratingBatch" @change="updateExcelBuffer(activeEmployee)" />
          </div>
          <div class="form-group w-20">
            <label>所属公司 (D2)</label>
            <input type="text" v-model="activeEmployee.company" class="form-control" placeholder="默认：四川久宏川科技有限公司" :disabled="isGeneratingBatch" @change="updateExcelBuffer(activeEmployee)" />
          </div>
          <div class="form-group w-20">
            <label>考核人姓名 (D4)</label>
            <input type="text" v-model="activeEmployee.evaluator" class="form-control" placeholder="默认：李杰、张剑锋" :disabled="isGeneratingBatch" @change="updateExcelBuffer(activeEmployee)" />
          </div>
        </div>

        <!-- 快速导入上一周期数据 -->
        <div class="panel-row">
          <div class="form-group w-100 file-upload-wrapper">
            <label>上月绩效数据源 (一键导入上月已填写的 Excel，自动回填公司、部门、考核人及历史考核任务)</label>
            <div class="file-upload-row">
              <label class="btn btn-outline upload-btn-full" :class="{ 'has-file': activeEmployee.lastMonthPerformance }">
                <Upload :size="15" />
                <span>{{ activeEmployee.lastMonthPerformance ? "重新导入上月 Excel 表单" : "一键导入上月 Excel 表单 (自动回填公司/部门/考核人/历史任务)" }}</span>
                <input type="file" accept=".xlsx" class="hidden-input" @change="(e) => handleLastMonthExcelUpload(e, activeEmployee!)" :disabled="isGeneratingBatch" />
              </label>
              <span class="imported-success-badge" v-if="activeEmployee.lastMonthPerformance">✅ 已解析并同步上月头部信息与工作任务</span>
            </div>
          </div>
        </div>

        <!-- 本月工作安排核心输入区：超大文本框，方便大屏下录入本月详细工作规划与要求 -->
        <div class="form-group margin-top">
          <label class="large-label">本月主要工作内容安排与计划要求 (AI 将依据此安排生成可量化的质量目标与标准，请尽量详细描述)</label>
          <textarea v-model="activeEmployee.thisMonthWorkContent" rows="12" placeholder="请在此输入本月任务计划，可分行描述，例如：&#10;1. 甲烷探测器APP开发：完成手持端视频播放、图片截图和历史报警查询功能。&#10;2. 无人机模块升级：配合后端实现V2.1.8版本安卓端和鸿蒙端的完整对接与联调。&#10;3. 测试BUG整改：负责解决测试人员反馈的APP侧缺陷，BUG解决率需达到90%以上。" class="form-control textarea large-textarea" :disabled="isGeneratingBatch"></textarea>
        </div>

        <!-- 控制按钮行 -->
        <div class="panel-actions">
          <button
            type="button"
            class="btn btn-primary btn-lg"
            :disabled="!geminiConfig.apiKey || isGeneratingBatch"
            @click="activeEmployee.status === 'generating' ? cancelAndRetry(activeEmployee) : generateSingle(activeEmployee)"
          >
            <RefreshCw v-if="activeEmployee.status === 'generating'" :size="16" class="animate-spin" />
            <RotateCcw v-else-if="activeEmployee.status === 'error'" :size="16" />
            <Sparkles v-else :size="16" />
            <span>
              {{ activeEmployee.status === "generating"
                ? `推导进行中 (${generatingElapsedSeconds}s)... 点击可重新推导`
                : activeEmployee.status === "error"
                ? "推导失败，点击重新推导"
                : activeEmployee.status === "success"
                ? "重新智能推导本月绩效"
                : "开始智能推导本月绩效" }}
            </span>
          </button>

          <!-- 正在生成时提供取消按钮 -->
          <button
            v-if="activeEmployee.status === 'generating'"
            type="button"
            class="btn btn-outline btn-lg text-danger"
            @click="cancelGeneration(activeEmployee)"
            title="中止当前正在等待的推导"
          >
            <XCircle :size="16" />
            <span>取消推导</span>
          </button>

          <button type="button" class="btn btn-outline text-danger btn-lg" @click="handleRemoveActive" :disabled="isGeneratingBatch || formalEmployees.length <= 1">
            <Trash2 :size="16" />
            <span>删除此员工</span>
          </button>
        </div>
      </div>

      <!-- 下部分：大屏预览与编辑区域 (100% 宽度，仿 Excel 网格样式) -->
      <div class="spreadsheet-panel-card" v-if="activeEmployee">
        <div class="section-title-row">
          <div class="title-group" style="display: flex; align-items: center; gap: 14px; flex-wrap: wrap">
            <span class="section-title">
              本月绩效计划书网格预览：
              <strong class="text-accent">{{ activeEmployee.name || "未命名" }}</strong>
            </span>
            <!-- 规则验证指示器 -->
            <div class="active-validation-warning" v-if="activeEmpValidationError">
              <AlertTriangle :size="14" style="color: #ef4444" />
              <span>规则警示: {{ activeEmpValidationError }}</span>
            </div>
          </div>
          <div class="right-actions" v-if="activeEmployee.status === 'success' && activeEmployee.tasks.length > 0">
            <!-- 权重检测 -->
            <span
              class="weight-badge"
              :class="{
                ok: Math.abs(getTasksTotalWeight(activeEmployee.tasks) - 1.0) < 0.001,
              }"
            >
              权重之和: {{ (getTasksTotalWeight(activeEmployee.tasks) * 100).toFixed(0) }}% ({{ Math.abs(getTasksTotalWeight(activeEmployee.tasks) - 1.0) < 0.001 ? "符合 100%" : "未达 100%" }})
            </span>

            <button type="button" class="btn btn-primary" @click="downloadSingle(activeEmployee)">
              <Download :size="15" />
              <span>导出已填报 Excel</span>
            </button>
          </div>
        </div>

        <!-- 电子表格网格 -->
        <div class="spreadsheet-container">
          <!-- 智能生成结果展示 -->
          <div v-if="activeEmployee.status === 'success' && activeEmployee.tasks.length > 0" class="excel-grid-wrapper">
            <table class="excel-grid">
              <thead>
                <tr>
                  <th class="col-seq">序号</th>
                  <th class="col-type">类型</th>
                  <th class="col-level">等级</th>
                  <th class="col-weight">权重</th>
                  <th class="col-cat">板块</th>
                  <th class="col-desc">解释说明 (工作目标/内容)</th>
                  <th class="col-time">时间目标</th>
                  <th class="col-count">数量</th>
                  <th class="col-quality">质量目标 (分条量化，禁用主观词)</th>
                  <th class="col-time-std">时间标准</th>
                  <th class="col-count-std">数量标准</th>
                  <th class="col-quality-std">质量标准 (分条量化，禁用主观词)</th>
                  <th class="col-action">操作</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(task, index) in activeEmployee.tasks" :key="index">
                  <!-- 序号 -->
                  <td class="cell-seq-num">{{ index + 1 }}</td>

                  <!-- 指标类型 (KPI / CPI) -->
                  <td>
                    <select v-model="task.type" class="grid-select" @change="handleCellChange(activeEmployee)">
                      <option value="KPI">KPI</option>
                      <option value="CPI">CPI</option>
                    </select>
                  </td>

                  <!-- 指标等级 -->
                  <td>
                    <select v-model="task.level" class="grid-select font-s" @change="handleCellChange(activeEmployee)">
                      <option value="核心战略任务">核心战略任务</option>
                      <option value="重要关键任务">重要关键任务</option>
                      <option value="常规执行任务">常规执行任务</option>
                      <option value="辅助零散任务">辅助零散任务</option>
                    </select>
                  </td>

                  <!-- 权重 -->
                  <td>
                    <input type="text" v-model="task.weight" class="grid-input text-center bold-text" placeholder="如 0.25 或 扣分项" @change="handleCellChange(activeEmployee)" />
                  </td>

                  <!-- 所属板块 -->
                  <td>
                    <input type="text" v-model="task.category" class="grid-input" @change="handleCellChange(activeEmployee)" />
                  </td>

                  <!-- 解释说明 (居中对齐) -->
                  <td>
                    <textarea v-model="task.description" class="grid-textarea desc-textarea" @change="handleCellChange(activeEmployee)"></textarea>
                  </td>

                  <!-- 时间目标 -->
                  <td>
                    <input type="text" v-model="task.time_target" class="grid-input" @change="handleCellChange(activeEmployee)" />
                  </td>

                  <!-- 数量目标 -->
                  <td>
                    <input type="text" v-model="task.count_target" class="grid-input text-center" @change="handleCellChange(activeEmployee)" />
                  </td>

                  <!-- 质量目标 -->
                  <td class="highlight-objective">
                    <textarea v-model="task.quality_target" class="grid-textarea font-s" @change="handleCellChange(activeEmployee)"></textarea>
                  </td>

                  <!-- 时间标准 -->
                  <td>
                    <textarea v-model="task.time_standard" class="grid-textarea font-s" @change="handleCellChange(activeEmployee)"></textarea>
                  </td>

                  <!-- 数量标准 -->
                  <td>
                    <input type="text" v-model="task.count_standard" class="grid-input text-center" @change="handleCellChange(activeEmployee)" />
                  </td>

                  <!-- 质量标准 -->
                  <td class="highlight-objective">
                    <textarea v-model="task.quality_standard" class="grid-textarea font-s" @change="handleCellChange(activeEmployee)"></textarea>
                  </td>

                  <!-- 删除单行 -->
                  <td class="text-center">
                    <button type="button" class="grid-action-delete" @click="removeTask(activeEmployee, index)" :disabled="activeEmployee.tasks.length <= 4" title="删除该任务行">
                      <Trash2 :size="13" />
                    </button>
                  </td>
                </tr>
              </tbody>
            </table>

            <div class="grid-actions-row">
              <button type="button" class="btn btn-outline btn-sm" @click="addTask(activeEmployee)">
                <Plus :size="14" />
                <span>新增考核项目 (任务行)</span>
              </button>
            </div>
          </div>

          <!-- 等待生成 -->
          <div v-else-if="activeEmployee.status === 'idle'" class="grid-placeholder">
            <Sparkles :size="50" class="glow-icon" />
            <h3>计划书尚未生成</h3>
            <p>请在上方面板中录入工作安排，并点击“开始智能推导”获取可量化的正式绩效计划明细表。</p>
          </div>

          <!-- 生成中 -->
          <div v-else-if="activeEmployee.status === 'generating'" class="grid-placeholder">
            <div class="spinner-large"></div>
            <h3>绩效计划自动生成中，正在严密对齐量化标准...</h3>
            <p>大模型正在严格匹配质量标准库，量化 1. 2. 3. 客观指标并自动平衡权重。</p>
            <p class="generating-timer-tip">
              ⏱️ 已等待 <strong>{{ generatingElapsedSeconds }}</strong> 秒（系统设置了 60 秒自动超时拦截，支持随时手动取消或重新推导）
            </p>
            <div style="margin-top: 18px; display: flex; gap: 12px; justify-content: center;">
              <button type="button" class="btn btn-outline" @click="cancelGeneration(activeEmployee)">
                <XCircle :size="16" />
                <span>取消等待</span>
              </button>
              <button type="button" class="btn btn-primary" @click="cancelAndRetry(activeEmployee)">
                <RefreshCw :size="16" />
                <span>立即重新推导</span>
              </button>
            </div>
          </div>

          <!-- 出错状态 -->
          <div v-else-if="activeEmployee.status === 'error'" class="grid-placeholder error-box">
            <AlertTriangle :size="48" class="text-danger" />
            <h3>推导遇到问题</h3>
            <p class="text-danger-detail">{{ activeEmployee.errorMessage }}</p>
            <div style="margin-top: 18px; display: flex; gap: 12px; justify-content: center;">
              <button type="button" class="btn btn-primary btn-lg" @click="generateSingle(activeEmployee)">
                <RefreshCw :size="16" />
                <span>重新推导绩效</span>
              </button>
            </div>
          </div>
        </div>
      </div>

      <!-- 全局批量打包控制栏 -->
      <div class="footer-actions-bar" v-if="formalEmployees.length > 0 && formalEmployees.some((e) => e.status === 'success')">
        <div class="left-stats">
          已成功推导
          <strong>{{ formalEmployees.filter((e) => e.status === "success").length }}</strong>
          / {{ formalEmployees.length }} 份正式绩效计划书
        </div>
        <div class="right-buttons">
          <button type="button" class="btn btn-outline" @click="resetFormalGenerationStatus" :disabled="isGeneratingBatch">
            <RefreshCw :size="16" />
            <span>重置所有状态</span>
          </button>
          <button type="button" class="btn btn-primary" @click="downloadAll" :disabled="isGeneratingBatch">
            <Download :size="16" />
            <span>一键导出全部 Excel</span>
          </button>
        </div>
      </div>
    </div>

    <!-- 绩效质量标准库管理弹窗 (全局挂载至 body，彻底脱离父级卡片上下文) -->
    <Teleport to="body">
      <div
        class="modal-mask standards-modal-mask"
        v-if="showStandardsModal"
        @click.self="showStandardsModal = false"
        @wheel.self.prevent
        @touchmove.self.prevent
      >
        <div class="modal-container glass-card standards-modal-container" @wheel.stop>
          <!-- 弹窗头部 -->
          <div class="modal-header">
            <div class="modal-header-title">
              <Target class="modal-title-icon" :size="22" />
              <div>
                <h3>绩效质量标准库配置与管理</h3>
                <p class="modal-header-desc">
                  支持下载官方规范模板、上传自定义标准库，并与系统默认库智能融合生效
                </p>
              </div>
            </div>
            <button type="button" class="btn-icon-close" @click="showStandardsModal = false" title="关闭弹窗 (Esc)">
              <X :size="20" />
            </button>
          </div>

          <!-- 弹窗内容主体 -->
          <div class="modal-body standards-modal-body">
            <!-- 状态概览看板 -->
            <div class="standards-stats-bar">
              <div class="stat-pill">
                <span class="pill-label">系统预置条目</span>
                <span class="pill-value">{{ defaultQualityStandards.length }} 条</span>
              </div>
              <div class="stat-pill" :class="{ 'pill-active': customQualityStandards.length > 0 }">
                <span class="pill-label">自定义已融合</span>
                <span class="pill-value">{{ customQualityStandards.length }} 条</span>
              </div>
              <div class="stat-pill pill-total">
                <span class="pill-label">当前生效总数</span>
                <span class="pill-value highlight">{{ qualityStandards.length }} 条</span>
              </div>
              <div class="standards-reset-wrap" v-if="customQualityStandards.length > 0">
                <button
                  type="button"
                  class="btn btn-outline btn-xs-reset"
                  @click="handleResetStandardsToDefault"
                  title="清除自定义上传，还原为纯系统预置标准库"
                >
                  <RotateCcw :size="13" />
                  <span>恢复系统默认</span>
                </button>
              </div>
            </div>

            <!-- 自定义库已激活横幅提示 -->
            <div class="custom-active-banner" v-if="customStandardsFileName">
              <CheckCircle2 :size="16" class="banner-icon" />
              <span>当前已融合自定义文件：<strong>{{ customStandardsFileName }}</strong>（同名分类优先覆盖，新增分类自动合流）。</span>
            </div>

            <!-- 操作卡片网格：下载模板 vs 上传文件 -->
            <div class="standards-actions-grid">
              <!-- 卡片 1: 下载模板 -->
              <div class="action-card">
                <div class="card-icon-wrap template-icon">
                  <FileSpreadsheet :size="24" />
                </div>
                <div class="card-text">
                  <h4>下载质量标准库模板</h4>
                  <p>包含 [序号、类型、质量目标、质量标准扣分细则] 的标准 4 列规范 Excel 模板。</p>
                </div>
                <button
                  type="button"
                  class="btn btn-outline btn-download-template"
                  @click="handleDownloadStandardsTemplate"
                >
                  <Download :size="15" />
                  <span>下载标准库模板 (.xlsx)</span>
                </button>
              </div>

              <!-- 卡片 2: 上传自定义标准 -->
              <div
                class="action-card upload-card"
                @dragover.prevent
                @drop="handleStandardsDrop"
              >
                <div class="card-icon-wrap upload-icon">
                  <FileUp :size="24" />
                </div>
                <div class="card-text">
                  <h4>上传自定义质量标准库</h4>
                  <p>支持拖拽或选择 .xlsx 文件，相同类型将智能覆盖扣分细则，新增类型自动合流。</p>
                </div>
                <label class="btn btn-primary btn-upload-standards" :class="{ disabled: isStandardsUploading }">
                  <input
                    type="file"
                    accept=".xlsx"
                    class="hidden-input"
                    @change="handleCustomStandardsUpload"
                    :disabled="isStandardsUploading"
                  />
                  <Upload :size="15" />
                  <span>{{ isStandardsUploading ? "正在解析导入..." : "选择文件上传 (.xlsx)" }}</span>
                </label>
              </div>
            </div>

            <!-- 上传错误警告 -->
            <div v-if="standardsUploadError" class="validation-warning standards-warning">
              <AlertTriangle :size="18" class="warning-icon" />
              <span>{{ standardsUploadError }}</span>
            </div>

            <!-- 当前生效规则实时预览区 -->
            <div class="standards-preview-section">
              <div class="preview-header">
                <div class="preview-title">
                  <h4>当前生效标准细则清单 ({{ filteredStandards.length }}/{{ qualityStandards.length }})</h4>
                  <span class="preview-hint">AI 绩效推导时将 100% 优先按以下扣分细则与目标执行对齐</span>
                </div>
                <div class="preview-search">
                  <Search :size="14" class="search-icon" />
                  <input
                    type="text"
                    v-model="standardsSearchQuery"
                    placeholder="搜索分类类型或标准关键词..."
                    class="search-input"
                  />
                </div>
              </div>

              <!-- 规则列表表格 -->
              <div class="standards-table-wrapper">
                <table class="standards-table">
                  <thead>
                    <tr>
                      <th style="width: 50px;">序号</th>
                      <th style="width: 140px;">业务类型 / 分类</th>
                      <th style="width: 90px;">来源</th>
                      <th style="width: 38%;">质量目标</th>
                      <th>质量标准（违规扣分细则）</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-if="filteredStandards.length === 0">
                      <td colspan="5" class="table-empty">
                        未找到符合 "{{ standardsSearchQuery }}" 的标准规范条目
                      </td>
                    </tr>
                    <tr
                      v-for="(item, idx) in filteredStandards"
                      :key="idx"
                      :class="{ 'row-custom': item.isCustom }"
                    >
                      <td class="text-center font-mono">{{ idx + 1 }}</td>
                      <td class="font-medium category-cell">
                        <span>{{ item.categoryType }}</span>
                      </td>
                      <td>
                        <span
                          class="source-badge"
                          :class="item.isCustom ? 'badge-custom' : 'badge-default'"
                        >
                          {{ item.isCustom ? "用户自定义" : "系统默认" }}
                        </span>
                      </td>
                      <td class="pre-cell">{{ item.qualityTarget }}</td>
                      <td class="pre-cell highlight-deduct">{{ item.qualityStandard }}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <!-- 弹窗底部 -->
          <div class="modal-footer standards-modal-footer">
            <div class="footer-tip">
              <Info :size="14" />
              <span>提示：修改或上传新标准库后，无需刷新页面，点击对应员工卡片的「重新推导」或「智能生成」即按新扣分细则生成。</span>
            </div>
            <button type="button" class="btn btn-outline" @click="showStandardsModal = false">
              完成并关闭
            </button>
          </div>
        </div>
      </div>
    </Teleport>
  </div>
</template>

<style scoped>
.formal-generator {
  animation: fadeIn 0.5s ease-out 0.2s both;
  width: 100%;
}

.card-header {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 18px 24px;
  border-bottom: 1px solid var(--border);
}

.header-icon {
  color: var(--accent);
}

.header-title-group {
  display: flex;
  align-items: baseline;
  gap: 8px;
  flex-grow: 1;
  flex-wrap: wrap;
}

.card-header h2 {
  margin: 0;
  font-size: 1.25rem;
  font-weight: 600;
  color: var(--text-h);
  text-align: left;
}

.header-subtitle {
  font-size: 0.85rem;
  color: var(--accent);
  font-weight: 500;
}

.standards-badge {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  font-size: 0.775rem;
  padding: 4px 10px;
  border-radius: 20px;
  background: rgba(168, 85, 247, 0.15);
  border: 1px solid rgba(168, 85, 247, 0.35);
  color: #c084fc;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s ease;
}

.standards-badge:hover {
  background: rgba(168, 85, 247, 0.25);
  border-color: rgba(168, 85, 247, 0.5);
  transform: translateY(-1px);
}

.custom-sub-tag {
  color: #34d399;
  font-weight: 600;
}

.btn-sm-action {
  font-size: 0.8rem !important;
  padding: 6px 12px !important;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  cursor: pointer;
}

.header-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.template-status-tag {
  font-size: 0.775rem;
  padding: 4px 10px;
  border-radius: 4px;
  background: rgba(239, 68, 68, 0.1);
  border: 1px solid rgba(239, 68, 68, 0.3);
  color: #ef4444;
}

.template-status-tag.ok {
  background: rgba(16, 185, 129, 0.1);
  border-color: rgba(16, 185, 129, 0.3);
  color: #10b981;
}

.btn-sm-upload {
  font-size: 0.8rem !important;
  padding: 6px 12px !important;
  cursor: pointer;
}

.hidden-input {
  display: none;
}

.card-body {
  padding: 24px;
  display: flex;
  flex-direction: column;
  gap: 24px;
}

/* 警示提示 */
.validation-warning {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 18px;
  border-radius: 6px;
  background: rgba(239, 68, 68, 0.08);
  border: 1px solid rgba(239, 68, 68, 0.25);
  color: #f87171;
  font-size: 0.875rem;
  text-align: left;
}

.warning-icon {
  flex-shrink: 0;
}

/* 员工选择横向页签 (全屏满宽显示) */
.employee-tabs {
  display: flex;
  gap: 8px;
  overflow-x: auto;
  border-bottom: 2px solid var(--border);
  padding-bottom: 8px;
  box-sizing: border-box;
}

.emp-tab {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  border: 1px solid var(--border);
  border-bottom: none;
  background: rgba(255, 255, 255, 0.02);
  color: var(--text);
  padding: 10px 18px;
  font-size: 0.85rem;
  font-weight: 600;
  cursor: pointer;
  border-top-left-radius: 6px;
  border-top-right-radius: 6px;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
  white-space: nowrap;
}

.emp-tab:hover:not(.active) {
  background: rgba(255, 255, 255, 0.06);
  color: var(--text-h);
}

.emp-tab.active {
  background: var(--bg-gradient);
  border-color: var(--accent-border);
  border-bottom: 2px solid #0f172a;
  color: var(--accent);
  box-shadow: 0 -2px 10px rgba(168, 85, 247, 0.05);
}

.emp-tab-pos {
  font-weight: 400;
  font-size: 0.725rem;
  color: var(--text);
  opacity: 0.7;
}

.emp-tab.success {
  border-top: 2px solid rgba(16, 185, 129, 0.5);
}

.emp-tab-badge {
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: var(--text);
}

.emp-tab-badge.idle {
  background: var(--text);
}
.emp-tab-badge.generating {
  background: var(--accent);
  animation: pulse 1.5s infinite;
}
.emp-tab-badge.success {
  background: #10b981;
}
.emp-tab-badge.error {
  background: #ef4444;
}

.tab-icon {
  opacity: 0.6;
}

.add-btn {
  border: 1px dashed var(--border);
  border-bottom: none;
  background: transparent;
  color: var(--text-h);
}

.add-btn:hover {
  border-color: var(--accent);
  color: var(--accent) !important;
  background: var(--accent-bg) !important;
}

/* 核心工作台：大型输入配置卡片 */
.input-panel-card {
  background: rgba(255, 255, 255, 0.015);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 24px;
  display: flex;
  flex-direction: column;
  gap: 16px;
  text-align: left;
}

.panel-row {
  display: flex;
  gap: 20px;
  flex-wrap: wrap;
}

.w-20 {
  width: calc(20% - 16px);
}
.w-30 {
  width: calc(30% - 10px);
}
.w-40 {
  width: calc(40% - 10px);
}
.w-50 {
  width: calc(50% - 10px);
}
.w-100 {
  width: 100%;
}

@media (max-width: 1100px) {
  .w-20 {
    width: calc(33.333% - 14px);
  }
}

@media (max-width: 768px) {
  .w-20,
  .w-30,
  .w-40,
  .w-50 {
    width: 100%;
  }
}

.form-group {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.form-group label {
  font-size: 0.8rem;
  font-weight: 700;
  color: var(--text-h);
}

.large-label {
  font-size: 0.875rem !important;
  color: var(--accent) !important;
  margin-bottom: 2px;
}

.form-control {
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid var(--border);
  color: var(--text-h);
  padding: 10px 14px;
  border-radius: 6px;
  font-size: 0.9rem;
  outline: none;
  transition: all 0.3s;
  box-sizing: border-box;
  width: 100%;
}

.form-control:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px var(--accent-bg);
  background: rgba(255, 255, 255, 0.06);
}

.textarea {
  font-family: inherit;
  resize: vertical;
  line-height: 1.5;
}

.large-textarea {
  font-size: 0.925rem;
  padding: 14px;
}

/* 导入数据源排版 */
.file-upload-wrapper {
  justify-content: space-between;
}

.file-upload-row {
  display: flex;
  align-items: center;
  gap: 14px;
}

.upload-btn-full {
  flex-grow: 1;
  padding: 8px 16px;
  font-size: 0.85rem !important;
  cursor: pointer;
  justify-content: center;
  background: rgba(168, 85, 247, 0.05);
  border-color: rgba(168, 85, 247, 0.3);
  color: var(--accent);
}

.upload-btn-full:hover {
  background: var(--accent-bg);
  border-color: var(--accent);
}

.upload-btn-full.has-file {
  background: rgba(16, 185, 129, 0.05);
  border-color: rgba(16, 185, 129, 0.3);
  color: #10b981;
}

.imported-success-badge {
  font-size: 0.775rem;
  color: #10b981;
  font-weight: 600;
  white-space: nowrap;
}

/* 上月折叠区 */
.history-collapse-section {
  border: 1px dashed var(--border);
  border-radius: 6px;
  background: rgba(255, 255, 255, 0.005);
  overflow: hidden;
}

.collapse-toggle-btn {
  width: 100%;
  padding: 8px 14px;
  background: transparent;
  border: none;
  color: var(--text);
  font-size: 0.775rem;
  display: flex;
  align-items: center;
  gap: 6px;
  cursor: pointer;
  transition: color 0.2s;
}

.collapse-toggle-btn:hover {
  color: var(--text-h);
}

.collapse-content {
  padding: 10px 14px;
  border-top: 1px dashed var(--border);
}

.margin-top {
  margin-top: 6px;
}

.panel-actions {
  display: flex;
  justify-content: flex-end;
  gap: 12px;
  margin-top: 10px;
  border-top: 1px solid var(--border);
  padding-top: 18px;
}

.btn-lg {
  padding: 10px 24px !important;
  font-size: 0.9rem !important;
}

/* 下半部分：100%满宽表格卡片 */
.spreadsheet-panel-card {
  background: rgba(255, 255, 255, 0.015);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 24px;
  display: flex;
  flex-direction: column;
  gap: 16px;
  text-align: left;
  width: 100%;
  box-sizing: border-box;
}

.spreadsheet-container {
  border: 1px solid var(--border);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.002);
  overflow: hidden;
  min-height: 250px;
  display: flex;
  flex-direction: column;
}

/* 仿 Excel 电子表格样式 (核心打磨) */
.excel-grid-wrapper {
  overflow-x: auto;
  width: 100%;
}

.excel-grid {
  width: 100%;
  border-collapse: collapse;
  text-align: left;
  font-size: 0.825rem;
}

.excel-grid th,
.excel-grid td {
  padding: 8px 10px;
  border-bottom: 1px solid var(--border);
  border-right: 1px solid var(--border);
  vertical-align: top;
}

.excel-grid th {
  background: rgba(255, 255, 255, 0.03);
  color: var(--text-h);
  font-weight: 600;
  font-size: 0.8rem;
  padding: 10px 8px;
  white-space: nowrap;
}

.excel-grid tr:hover td {
  background: rgba(255, 255, 255, 0.01);
}

/* 列宽微调 - 加宽以满足大屏编辑与预览需求 */
.col-seq {
  width: 45px;
  text-align: center;
}
.col-type {
  width: 85px;
}
.col-level {
  width: 140px;
}
.col-weight {
  width: 100px;
}
.col-cat {
  width: 110px;
}
.col-desc {
  min-width: 220px;
  text-align: center;
} /* 原 180px */
.col-time {
  width: 150px;
} /* 原 105px */
.col-count {
  width: 65px;
}
.col-quality {
  min-width: 300px;
} /* 原 200px */
.col-time-std {
  min-width: 130px;
} /* 原 150px */
.col-count-std {
  width: 85px;
}
.col-quality-std {
  min-width: 300px;
} /* 原 200px */
.col-action {
  width: 45px;
  text-align: center;
}

.cell-seq-num {
  text-align: center;
  background: rgba(255, 255, 255, 0.015);
  font-weight: 600;
  color: var(--text);
  user-select: none;
}

/* 仿 Excel 单元格内无缝编辑 */
.grid-input {
  background: transparent;
  border: 1px solid transparent;
  color: var(--text-h);
  padding: 4px 6px;
  width: 100%;
  box-sizing: border-box;
  font-size: 0.825rem;
  border-radius: 4px;
  outline: none;
  transition: all 0.2s;
}

.grid-input:hover {
  border-color: rgba(255, 255, 255, 0.15);
  background: rgba(255, 255, 255, 0.02);
}

.grid-input:focus {
  border-color: var(--accent);
  background: rgba(255, 255, 255, 0.06);
  box-shadow: 0 0 0 2px var(--accent-bg);
}

.grid-select {
  /* 隐藏浏览器默认的下拉箭头，以便应用自定义 SVG 指针 */
  appearance: none;
  background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23a855f7' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: right 8px center;
  background-size: 12px;
  background-color: rgba(255, 255, 255, 0.03);
  border: 1px solid var(--border);
  color: var(--text-h);
  padding: 4px 24px 4px 8px;
  border-radius: 4px;
  outline: none;
  font-size: 0.8rem;
  width: 100%;
  cursor: pointer;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
}

.grid-select:hover {
  border-color: var(--accent);
  background-color: rgba(255, 255, 255, 0.06);
}

.grid-select:focus {
  border-color: var(--accent);
  background-color: rgba(255, 255, 255, 0.08);
  box-shadow: 0 0 0 2px var(--accent-bg);
}

/* 下拉菜单项的暗色背景与高对比度文字 */
.grid-select option {
  background-color: #1e293b;
  color: #f8fafc;
  padding: 8px;
}

.grid-textarea {
  text-align: left;
  background: transparent;
  border: 1px solid transparent;
  color: var(--text-h);
  padding: 8px;
  width: 100%;
  height: 100px; /* 从 80px 调整为更舒适的 100px */
  resize: vertical;
  box-sizing: border-box;
  font-size: 0.825rem;
  line-height: 1.45;
  font-family: inherit;
  outline: none;
  border-radius: 4px;
  transition: all 0.2s;
}

.desc-textarea {
  text-align: center !important;
}

.grid-textarea:hover {
  border-color: rgba(255, 255, 255, 0.15);
  background: rgba(255, 255, 255, 0.02);
}

.grid-textarea:focus {
  border-color: var(--accent);
  background: rgba(255, 255, 255, 0.06);
  box-shadow: 0 0 0 2px var(--accent-bg);
}

/* 客观量化列的柔和标记 */
.highlight-objective {
  background: rgba(168, 85, 247, 0.005);
}

.highlight-objective:focus-within {
  background: rgba(168, 85, 247, 0.015);
}

.font-s {
  font-size: 0.75rem !important;
}

.bold-text {
  font-weight: 700;
}

/* 单元格删除按钮的精致化设计 */
.grid-action-delete {
  background: rgba(239, 68, 68, 0.05);
  border: 1px solid rgba(239, 68, 68, 0.15);
  color: var(--text);
  cursor: pointer;
  padding: 6px;
  border-radius: 4px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-top: 4px;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.grid-action-delete:hover:not(:disabled) {
  background: rgba(239, 68, 68, 0.18);
  border-color: #ef4444;
  color: #ef4444;
  transform: scale(1.08);
}

.grid-action-delete:active:not(:disabled) {
  transform: scale(0.92);
}

.grid-action-delete:disabled {
  opacity: 0.25;
  cursor: not-allowed;
  background: transparent;
  border-color: transparent;
}

/* 统一基础按钮样式 (符合现有高档玻璃拟态与渐变风格) */
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  font-size: 0.9rem;
  font-weight: 600;
  padding: 8px 18px;
  border-radius: 8px;
  cursor: pointer;
  border: 1px solid rgba(255, 255, 255, 0.08);
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  background: rgba(255, 255, 255, 0.03);
  color: var(--text-h);
  backdrop-filter: blur(4px);
}

.btn-primary {
  background: linear-gradient(135deg, #a855f7 0%, #6366f1 100%);
  color: white;
  border: 1px solid rgba(255, 255, 255, 0.15);
  box-shadow: 0 4px 14px rgba(168, 85, 247, 0.25);
  text-shadow: 0 1px 2px rgba(0, 0, 0, 0.15);
}

.btn-primary:hover:not(:disabled) {
  transform: translateY(-2px);
  box-shadow: 0 6px 20px rgba(168, 85, 247, 0.4);
  background: linear-gradient(135deg, #b86bfb 0%, #7578f3 100%);
}

.btn-primary:active:not(:disabled) {
  transform: translateY(0);
  box-shadow: 0 2px 10px rgba(168, 85, 247, 0.2);
}

.btn-primary:disabled {
  opacity: 0.4;
  cursor: not-allowed;
  transform: none !important;
  box-shadow: none !important;
}

.btn-outline {
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid var(--border);
  color: var(--text);
}

.btn-outline:hover:not(:disabled) {
  background: rgba(255, 255, 255, 0.08);
  border-color: rgba(168, 85, 247, 0.5);
  color: var(--text-h);
  box-shadow: 0 4px 12px rgba(168, 85, 247, 0.1);
  transform: translateY(-1px);
}

.btn-outline:active:not(:disabled) {
  transform: translateY(0);
  background: rgba(255, 255, 255, 0.05);
}

.btn-outline:disabled {
  opacity: 0.4;
  cursor: not-allowed;
  transform: none !important;
  box-shadow: none !important;
}

.text-danger {
  color: #ef4444 !important;
}

.btn-outline.text-danger:hover:not(:disabled) {
  border-color: rgba(239, 68, 68, 0.5) !important;
  background: rgba(239, 68, 68, 0.08) !important;
  color: #ef4444 !important;
  box-shadow: 0 4px 12px rgba(239, 68, 68, 0.1) !important;
}

.grid-actions-row {
  padding: 12px 16px;
  background: rgba(255, 255, 255, 0.01);
  border-top: 1px solid var(--border);
}

/* 占位区 */
.grid-placeholder {
  padding: 80px 40px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  text-align: center;
  flex-grow: 1;
}

.grid-placeholder.error-box {
  background: rgba(239, 68, 68, 0.005);
}

.glow-icon {
  color: var(--accent);
  opacity: 0.5;
  filter: drop-shadow(0 0 10px var(--accent-border));
}

.grid-placeholder h3 {
  margin: 0;
  font-size: 1.1rem;
  color: var(--text-h);
}

.grid-placeholder p {
  margin: 0;
  font-size: 0.85rem;
  color: var(--text);
  max-width: 480px;
  line-height: 1.6;
}

.text-danger-detail {
  font-size: 0.8rem;
  color: #f87171;
  max-width: 500px;
  word-break: break-all;
}

.spinner-large {
  width: 40px;
  height: 40px;
  border: 3px solid rgba(255, 255, 255, 0.1);
  border-radius: 50%;
  border-top-color: var(--accent);
  animation: spin 1s ease-in-out infinite;
}

@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}

@keyframes pulse {
  0%,
  100% {
    transform: scale(1);
    opacity: 1;
  }
  50% {
    transform: scale(1.1);
    opacity: 0.5;
  }
}

/* 核心网格预览区辅助样式 */
.section-title-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 16px;
  border-bottom: 1px solid var(--border);
  padding-bottom: 12px;
  margin-bottom: 8px;
}

.section-title {
  font-size: 1.1rem;
  color: var(--text-h);
  font-weight: 600;
}

.weight-badge {
  font-size: 0.775rem;
  padding: 4px 10px;
  border-radius: 4px;
  background: rgba(239, 68, 68, 0.1);
  border: 1px solid rgba(239, 68, 68, 0.25);
  color: #ef4444;
  font-weight: 700;
  display: inline-flex;
  align-items: center;
}

.weight-badge.ok {
  background: rgba(16, 185, 129, 0.1);
  border-color: rgba(16, 185, 129, 0.25);
  color: #10b981;
}

.right-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.active-validation-warning {
  background: rgba(239, 68, 68, 0.08);
  border: 1px solid rgba(239, 68, 68, 0.25);
  color: #f87171;
  padding: 4px 12px;
  border-radius: 6px;
  font-size: 0.8rem;
  font-weight: 500;
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.generating-timer-tip {
  color: var(--accent);
  font-size: 0.875rem;
  margin-top: 6px;
}

.generating-timer-tip strong {
  font-size: 1.15rem;
  color: #f8fafc;
  padding: 0 4px;
}

.btn-error-retry {
  background: linear-gradient(135deg, #f59e0b 0%, #ef4444 100%) !important;
  box-shadow: 0 4px 16px rgba(239, 68, 68, 0.3) !important;
}

.btn-error-retry:hover {
  background: linear-gradient(135deg, #fbbf24 0%, #f87171 100%) !important;
  transform: translateY(-2px);
}

/* ============================================================
   绩效质量标准库管理弹窗样式
   ============================================================ */

/* 全局宿主页面滚动锁定（弹窗开启期间禁止外部文档滚动穿透） */
:global(body.modal-scroll-locked) {
  overflow: hidden !important;
  overscroll-behavior: none !important;
  touch-action: none !important;
}

.modal-mask {
  position: fixed;
  inset: 0;
  width: 100vw;
  height: 100vh;
  background: rgba(0, 0, 0, 0.72);
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
  z-index: 99999;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px;
  box-sizing: border-box;
  overscroll-behavior: contain;
}

.standards-modal-container {
  width: 1020px;
  max-width: 95vw;
  max-height: 90vh;
  background: #111827;
  border: 1px solid rgba(255, 255, 255, 0.12);
  border-radius: 16px;
  box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.75);
  display: flex;
  flex-direction: column;
  overflow: hidden;
  overscroll-behavior: contain;
  animation: modalScaleIn 0.25s cubic-bezier(0.16, 1, 0.3, 1);
}

@keyframes modalScaleIn {
  from {
    opacity: 0;
    transform: scale(0.95) translateY(10px);
  }
  to {
    opacity: 1;
    transform: scale(1) translateY(0);
  }
}

.standards-modal-container .modal-header {
  padding: 18px 24px;
  border-bottom: 1px solid rgba(255, 255, 255, 0.08);
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: rgba(255, 255, 255, 0.02);
}

.modal-header-title {
  display: flex;
  align-items: center;
  gap: 12px;
}

.modal-title-icon {
  color: var(--accent);
  flex-shrink: 0;
}

.modal-header-title h3 {
  margin: 0;
  font-size: 1.15rem;
  font-weight: 600;
  color: #f8fafc;
}

.modal-header-desc {
  margin: 4px 0 0 0;
  font-size: 0.8rem;
  color: #94a3b8;
}

.btn-icon-close {
  background: transparent;
  border: none;
  color: #94a3b8;
  cursor: pointer;
  padding: 6px;
  border-radius: 8px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s;
}

.btn-icon-close:hover {
  background: rgba(255, 255, 255, 0.1);
  color: #f8fafc;
}

.standards-modal-body {
  padding: 20px 24px;
  overflow-y: auto;
  overscroll-behavior: contain;
  display: flex;
  flex-direction: column;
  gap: 16px;
  max-height: calc(90vh - 150px);
}

/* 状态统计看板 */
.standards-stats-bar {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
  padding: 12px 16px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid rgba(255, 255, 255, 0.08);
  border-radius: 10px;
}

.stat-pill {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 6px 12px;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid rgba(255, 255, 255, 0.06);
  border-radius: 6px;
  font-size: 0.825rem;
}

.stat-pill.pill-active {
  background: rgba(16, 185, 129, 0.1);
  border-color: rgba(16, 185, 129, 0.3);
  color: #34d399;
}

.stat-pill.pill-total {
  background: rgba(99, 102, 241, 0.12);
  border-color: rgba(99, 102, 241, 0.3);
}

.pill-label {
  color: #94a3b8;
}

.pill-value {
  font-weight: 600;
  color: #f1f5f9;
}

.pill-value.highlight {
  color: #818cf8;
}

.standards-reset-wrap {
  margin-left: auto;
}

.btn-xs-reset {
  font-size: 0.775rem !important;
  padding: 4px 10px !important;
  display: inline-flex;
  align-items: center;
  gap: 5px;
  color: #f87171 !important;
  border-color: rgba(239, 68, 68, 0.3) !important;
}

.btn-xs-reset:hover {
  background: rgba(239, 68, 68, 0.1) !important;
  border-color: rgba(239, 68, 68, 0.5) !important;
}

.custom-active-banner {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 16px;
  background: rgba(16, 185, 129, 0.08);
  border: 1px solid rgba(16, 185, 129, 0.25);
  border-radius: 8px;
  font-size: 0.825rem;
  color: #34d399;
}

.banner-icon {
  flex-shrink: 0;
}

/* 操作网格 */
.standards-actions-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px;
}

.action-card {
  display: flex;
  flex-direction: column;
  gap: 12px;
  padding: 18px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid rgba(255, 255, 255, 0.08);
  border-radius: 12px;
  transition: all 0.2s ease;
}

.action-card:hover {
  background: rgba(255, 255, 255, 0.05);
  border-color: rgba(255, 255, 255, 0.15);
}

.upload-card {
  border-style: dashed;
}

.upload-card:hover {
  border-color: var(--accent);
  background: rgba(99, 102, 241, 0.05);
}

.card-icon-wrap {
  width: 44px;
  height: 44px;
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.card-icon-wrap.template-icon {
  background: rgba(16, 185, 129, 0.12);
  color: #10b981;
}

.card-icon-wrap.upload-icon {
  background: rgba(99, 102, 241, 0.12);
  color: #818cf8;
}

.card-text h4 {
  margin: 0 0 4px 0;
  font-size: 0.95rem;
  font-weight: 600;
  color: #f1f5f9;
}

.card-text p {
  margin: 0;
  font-size: 0.775rem;
  color: #94a3b8;
  line-height: 1.4;
}

.btn-download-template,
.btn-upload-standards {
  margin-top: auto;
  font-size: 0.825rem;
  padding: 8px 14px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  cursor: pointer;
}

.btn-upload-standards.disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.standards-warning {
  margin: 0;
}

/* 预览区 */
.standards-preview-section {
  display: flex;
  flex-direction: column;
  gap: 12px;
  margin-top: 4px;
}

.preview-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  flex-wrap: wrap;
}

.preview-title h4 {
  margin: 0;
  font-size: 0.95rem;
  font-weight: 600;
  color: #f1f5f9;
}

.preview-hint {
  font-size: 0.75rem;
  color: #94a3b8;
  display: block;
  margin-top: 2px;
}

.preview-search {
  position: relative;
  width: 260px;
}

.preview-search .search-icon {
  position: absolute;
  left: 10px;
  top: 50%;
  transform: translateY(-50%);
  color: #64748b;
  pointer-events: none;
}

.preview-search .search-input {
  width: 100%;
  box-sizing: border-box;
  padding: 6px 12px 6px 30px;
  border-radius: 6px;
  background: rgba(0, 0, 0, 0.3);
  border: 1px solid rgba(255, 255, 255, 0.1);
  color: #f1f5f9;
  font-size: 0.8rem;
  outline: none;
  transition: all 0.2s;
}

.preview-search .search-input:focus {
  border-color: var(--accent);
  background: rgba(0, 0, 0, 0.4);
}

.standards-table-wrapper {
  max-height: 320px;
  overflow-y: auto;
  overscroll-behavior: contain;
  border: 1px solid rgba(255, 255, 255, 0.08);
  border-radius: 10px;
  background: rgba(0, 0, 0, 0.2);
}

.standards-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.8rem;
  text-align: left;
}

.standards-table th {
  position: sticky;
  top: 0;
  background: #1e293b;
  padding: 10px 14px;
  font-weight: 600;
  color: #94a3b8;
  border-bottom: 1px solid rgba(255, 255, 255, 0.1);
  z-index: 2;
  white-space: nowrap;
}

.standards-table td {
  padding: 10px 14px;
  border-bottom: 1px solid rgba(255, 255, 255, 0.05);
  vertical-align: top;
  color: #cbd5e1;
  line-height: 1.5;
}

.standards-table tr:hover {
  background: rgba(255, 255, 255, 0.02);
}

.row-custom {
  background: rgba(16, 185, 129, 0.03);
}

.row-custom td:first-child {
  border-left: 3px solid #10b981;
}

.category-cell {
  color: #f1f5f9;
  font-weight: 600;
}

.source-badge {
  display: inline-block;
  font-size: 0.7rem;
  padding: 2px 7px;
  border-radius: 4px;
  white-space: nowrap;
}

.badge-default {
  background: rgba(148, 163, 184, 0.12);
  color: #94a3b8;
  border: 1px solid rgba(148, 163, 184, 0.2);
}

.badge-custom {
  background: rgba(16, 185, 129, 0.15);
  color: #34d399;
  border: 1px solid rgba(16, 185, 129, 0.35);
  font-weight: 600;
}

.pre-cell {
  white-space: pre-wrap;
  word-break: break-word;
}

.highlight-deduct {
  color: #fcd34d;
}

.table-empty {
  text-align: center;
  padding: 30px !important;
  color: #94a3b8;
}

/* 弹窗底部 */
.standards-modal-footer {
  padding: 14px 24px;
  border-top: 1px solid rgba(255, 255, 255, 0.08);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  background: rgba(255, 255, 255, 0.02);
}

.footer-tip {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 0.775rem;
  color: #94a3b8;
}

.footer-tip svg {
  flex-shrink: 0;
  color: var(--accent);
}

@media (max-width: 768px) {
  .standards-actions-grid {
    grid-template-columns: 1fr;
  }
  .preview-header {
    flex-direction: column;
    align-items: stretch;
  }
  .preview-search {
    width: 100%;
  }
  .standards-modal-footer {
    flex-direction: column;
    align-items: stretch;
  }
}
</style>
