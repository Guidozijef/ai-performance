<script setup lang="ts">
/**
 * @fileoverview 正式员工绩效表单（JX1.6套表）自动生成组件。
 * 提供宽屏大卡片工作台、隐藏上月绩效详情以精简视觉、全宽大型本月工作安排编辑框、以及 100% 满宽绩效编辑网格。
 * 导出时自动清除红色背景填充，原样保留 Excel 字体、合并单元格及边框样式。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref, onMounted, computed } from 'vue';
import { 
  geminiConfig, 
  excelTemplate, 
  formalEmployees, 
  addFormalEmployee, 
  removeFormalEmployee, 
  resetFormalGenerationStatus
} from '../store';
import type { FormalEmployeeRow } from '../store';
import { generateFormalPerformance } from '../utils/geminiHelper';
import { writePerformanceToTemplate, downloadExcelFile, readPerformanceFromExcel } from '../utils/excelHelper';
import type { PerformanceTask } from '../utils/excelHelper';
import { 
  Users, 
  Trash2, 
  Sparkles, 
  Download, 
  AlertTriangle, 
  FileSpreadsheet,
  RefreshCw,
  Upload,
  Plus
} from 'lucide-vue-next';

// 记录当前正在编辑任务的员工 ID
const activeEmpId = ref<string | null>(null);
// 记录批量生成状态
const isGeneratingBatch = ref(false);
// 加载提示
const templateLoadingStatus = ref('');

// 计算当前选中的正在编辑绩效计划的员工
const activeEmployee = computed(() => {
  return formalEmployees.value.find(emp => emp.id === activeEmpId.value) || null;
});

// 全局基础校验
const validationError = computed(() => {
  if (!excelTemplate.buffer) {
    return '未加载 Excel 模板，请确保 public 下存在模板或上传新模板。';
  }
  if (!geminiConfig.apiKey) {
    return '请先在设置中填写 Gemini API Key。';
  }
  if (formalEmployees.value.length === 0) {
    return '请至少添加一名正式员工。';
  }
  return '';
});

// 计算当前正在编辑员工的任务指标规则校验
// 包含四项核心校验规则：
// 1. 各项指标权重不超过 30%
// 2. 指标总权重刚好等于 100%
// 3. 重要战略/关键任务对应的类型必须为 KPI，常规/辅助任务对应的类型必须为 CPI
// 4. 重要任务比例最低不能低于 20%
const activeEmpValidationError = computed(() => {
  const emp = activeEmployee.value;
  if (!emp || emp.status !== 'success' || !emp.tasks || emp.tasks.length === 0) {
    return '';
  }

  // 1. 校验权重总和是否为 1.0 (100%)
  const totalW = getTasksTotalWeight(emp.tasks);
  if (Math.abs(totalW - 1.0) > 0.001) {
    return `总权重为 ${(totalW * 100).toFixed(0)}%，必须刚好等于 100%。`;
  }

  // 2. 校验指标最大限制与最低比例约束
  for (let i = 0; i < emp.tasks.length; i++) {
    const t = emp.tasks[i];
    const w = parseFloat(t.weight as any || 0);

    if (w > 0.30001) {
      return `第 ${i + 1} 项“${t.category || '指标'}”权重为 ${(w * 100).toFixed(0)}%，已超过最大限制 30%。`;
    }

    if (t.level === '重要关键任务' && w < 0.1999) {
      return `重要关键任务（如第 ${i + 1} 项）的权重为 ${(w * 100).toFixed(0)}%，低于最低占比要求 20%。`;
    }

    if ((t.level === '常规执行任务' || t.level === '辅助零散任务') && t.type !== 'CPI') {
      return `常规/辅助任务（如第 ${i + 1} 项）类型当前为 ${t.type}，必须设为 CPI。`;
    }
  }

  return '';
});

/**
 * 组件加载时自动拉取 Public 下的默认 Excel 模板文件
 */
onMounted(async () => {
  if (!excelTemplate.buffer) {
    templateLoadingStatus.value = '正在自动加载预置绩效模板 JX1.6...';
    try {
      const response = await fetch('/正式员工绩效表单套表-JX1.6.xlsx');
      if (response.ok) {
        const buffer = await response.arrayBuffer();
        excelTemplate.buffer = buffer;
        excelTemplate.fileName = '正式员工绩效表单套表-JX1.6.xlsx (预置)';
        templateLoadingStatus.value = '成功载入预置绩效模板 JX1.6';
      } else {
        templateLoadingStatus.value = '未找到预置绩效模板，请上传模板文件。';
      }
    } catch (e) {
      console.error('加载预置模板失败:', e);
      templateLoadingStatus.value = '加载预置模板失败，请手动上传。';
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
    excelTemplate.fileName = `${file.name} (手动上传)`;
    excelTemplate.buffer = arrayBuffer;
    templateLoadingStatus.value = '成功载入自定义绩效模板';
  } catch (error: any) {
    console.error('解析模板失败:', error);
    templateLoadingStatus.value = '自定义模板载入失败，请检查文件。';
  } finally {
    target.value = '';
  }
}

/**
 * 重算并更新某位员工导出的 Excel 二进制缓存 (回写并清除红色背景)
 */
async function updateExcelBuffer(emp: FormalEmployeeRow): Promise<boolean> {
  if (!excelTemplate.buffer || emp.tasks.length === 0) return false;

  try {
    const buffer = await writePerformanceToTemplate(
      excelTemplate.buffer,
      emp.name,
      emp.position,
      emp.tasks
    );
    emp.outputBuffer = buffer;
    emp.fileName = `绩效计划表_${emp.name}_${emp.position}_${new Date().toISOString().slice(0, 7)}.xlsx`;
    return true;
  } catch (e: any) {
    console.error('回写 Excel 缓存失败:', e);
    emp.errorMessage = `回写 Excel 失败: ${e.message}`;
    emp.status = 'error';
    return false;
  }
}

/**
 * 单个生成逻辑
 */
async function generateSingle(emp: FormalEmployeeRow) {
  if (!geminiConfig.apiKey) return;
  emp.status = 'generating';
  emp.errorMessage = undefined;
  emp.tasks = [];

  const result = await generateFormalPerformance(
    {
      apiKey: geminiConfig.apiKey,
      proxyUrl: geminiConfig.proxyUrl,
      model: geminiConfig.model,
      systemInstruction: geminiConfig.systemInstruction
    },
    emp.name,
    emp.position,
    emp.lastMonthPerformance,
    emp.thisMonthWorkContent
  );

  if (result.success) {
    emp.tasks = result.tasks;
    emp.name = result.name;
    emp.position = result.position;
    
    // 渲染回写并清除红色背景
    const successWrite = await updateExcelBuffer(emp);
    if (successWrite) {
      emp.status = 'success';
    }
  } else {
    emp.errorMessage = result.message;
    emp.status = 'error';
  }
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
  const successRows = formalEmployees.value.filter(emp => emp.status === 'success' && emp.outputBuffer && emp.fileName);
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
    type: 'KPI',
    level: '重要关键任务',
    weight: 0.25,
    category: '系统开发',
    description: '新增的考核任务',
    time_target: '当月月底前',
    count_target: '/',
    quality_target: '1.按时按质交付开发任务\n2.单元测试通过率达100%',
    time_standard: '每延迟一天扣2分',
    count_standard: '/',
    quality_standard: '1.未按原型要求交付该项扣10分\n2.收到投诉扣2分/次'
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
 * 校验当前任务的总权重是否为 1.0 (100%)
 */
function getTasksTotalWeight(tasks: PerformanceTask[]): number {
  return tasks.reduce((sum, t) => sum + parseFloat(t.weight as any || 0), 0);
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
  const index = formalEmployees.value.findIndex(emp => emp.id === currentId);
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
    
    // 自动回填姓名与岗位
    if (data.name) emp.name = data.name;
    if (data.position) emp.position = data.position;
    
    // 格式化上月指标，用于 Gemini 提示上下文
    const taskSummary = data.tasks.map((t: PerformanceTask, idx: number) => {
      return `【任务 ${idx + 1}】
- 类型: ${t.type} (${t.level}) | 权重: ${(t.weight * 100).toFixed(0)}% | 板块: ${t.category}
- 任务描述: ${t.description}
- 时间目标: ${t.time_target} | 数量目标: ${t.count_target}
- 质量目标: ${t.quality_target}
- 时间标准: ${t.time_standard} | 数量标准: ${t.count_standard}
- 质量标准: ${t.quality_standard}`;
    }).join('\n\n');

    emp.lastMonthPerformance = `[已解析自上月绩效 Excel] ${file.name}\n姓名：${data.name} | 岗位：${data.position}\n\n上月考核项回顾：\n${taskSummary}`;
  } catch (error: any) {
    console.error('解析上月绩效 Excel 失败:', error);
    alert(`解析失败: ${error.message || '文件格式错误'}`);
  } finally {
    target.value = '';
  }
}
</script>

<template>
  <div class="glass-card formal-generator">
    <!-- 头部栏 -->
    <div class="card-header">
      <FileSpreadsheet class="header-icon" />
      <h2>正式员工绩效智能填报 (JX1.6 套表专区)</h2>
      <div class="header-actions">
        <span class="template-status-tag" :class="{ ok: excelTemplate.buffer }">
          {{ templateLoadingStatus || '未检测到模板' }}
        </span>
        <label class="btn btn-outline btn-sm-upload">
          <input type="file" accept=".xlsx" class="hidden-input" @change="handleTemplateUpload" />
          <span>更换 Excel 模板</span>
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
        <button 
          v-for="emp in formalEmployees" 
          :key="emp.id" 
          class="emp-tab"
          :class="{ active: emp.id === activeEmpId, success: emp.status === 'success' }"
          @click="activeEmpId = emp.id"
        >
          <Users :size="14" class="tab-icon" />
          <span class="emp-tab-name">{{ emp.name || '新员工' }}</span>
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
        <!-- 员工信息与导入行 -->
        <div class="panel-row">
          <div class="form-group w-30">
            <label>被考核人姓名 (D3)</label>
            <input 
              type="text" 
              v-model="activeEmployee.name" 
              class="form-control" 
              placeholder="例如：杨砚翔" 
              :disabled="isGeneratingBatch"
            />
          </div>
          <div class="form-group w-30">
            <label>岗位名称 (L3)</label>
            <input 
              type="text" 
              v-model="activeEmployee.position" 
              class="form-control" 
              placeholder="例如：APP开发工程师" 
              :disabled="isGeneratingBatch"
            />
          </div>
          <div class="form-group w-40 file-upload-wrapper">
            <label>上月绩效数据源 (智能回填并参考)</label>
            <div class="file-upload-row">
              <label class="btn btn-outline upload-btn-full" :class="{ 'has-file': activeEmployee.lastMonthPerformance }">
                <Upload :size="15" />
                <span>{{ activeEmployee.lastMonthPerformance ? '重新导入上月 Excel' : '直接导入上月已填写 Excel' }}</span>
                <input 
                  type="file" 
                  accept=".xlsx" 
                  class="hidden-input" 
                  @change="(e) => handleLastMonthExcelUpload(e, activeEmployee!)" 
                  :disabled="isGeneratingBatch"
                />
              </label>
              <span class="imported-success-badge" v-if="activeEmployee.lastMonthPerformance">
                ✅ 已解析上月任务
              </span>
            </div>
          </div>
        </div>

        <!-- 本月工作安排核心输入区：超大文本框，方便大屏下录入本月详细工作规划与要求 -->
        <div class="form-group margin-top">
          <label class="large-label">本月主要工作内容安排与计划要求 (AI 将依据此安排生成可量化的质量目标与标准，请尽量详细描述)</label>
          <textarea 
            v-model="activeEmployee.thisMonthWorkContent" 
            rows="12" 
            placeholder="请在此输入本月任务计划，可分行描述，例如：&#10;1. 甲烷探测器APP开发：完成手持端视频播放、图片截图和历史报警查询功能。&#10;2. 无人机模块升级：配合后端实现V2.1.8版本安卓端和鸿蒙端的完整对接与联调。&#10;3. 测试BUG整改：负责解决测试人员反馈的APP侧缺陷，BUG解决率需达到90%以上。" 
            class="form-control textarea large-textarea"
            :disabled="isGeneratingBatch"
          ></textarea>
        </div>

        <!-- 控制按钮行 -->
        <div class="panel-actions">
          <button 
            type="button" 
            class="btn btn-primary btn-lg" 
            :disabled="!geminiConfig.apiKey || activeEmployee.status === 'generating' || isGeneratingBatch"
            @click="generateSingle(activeEmployee)"
          >
            <Sparkles :size="16" />
            <span>{{ activeEmployee.status === 'success' ? '重新智能推导本月绩效' : '开始智能推导本月绩效' }}</span>
          </button>
          
          <button 
            type="button" 
            class="btn btn-outline text-danger btn-lg" 
            @click="handleRemoveActive"
            :disabled="isGeneratingBatch || formalEmployees.length <= 1"
          >
            <Trash2 :size="16" />
            <span>删除此员工</span>
          </button>
        </div>
      </div>

      <!-- 下部分：大屏预览与编辑区域 (100% 宽度，仿 Excel 网格样式) -->
      <div class="spreadsheet-panel-card" v-if="activeEmployee">
        <div class="section-title-row">
          <div class="title-group" style="display: flex; align-items: center; gap: 14px; flex-wrap: wrap;">
            <span class="section-title">
              本月绩效计划书网格预览：<strong class="text-accent">{{ activeEmployee.name || '未命名' }}</strong>
            </span>
            <!-- 规则验证指示器 -->
            <div class="active-validation-warning" v-if="activeEmpValidationError">
              <AlertTriangle :size="14" style="color: #ef4444;" />
              <span>规则警示: {{ activeEmpValidationError }}</span>
            </div>
          </div>
          <div class="right-actions" v-if="activeEmployee.status === 'success' && activeEmployee.tasks.length > 0">
            <!-- 权重检测 -->
            <span 
              class="weight-badge" 
              :class="{ 
                ok: Math.abs(getTasksTotalWeight(activeEmployee.tasks) - 1.0) < 0.001 
              }"
            >
              权重之和: {{ (getTasksTotalWeight(activeEmployee.tasks) * 100).toFixed(0) }}%
              ({{ Math.abs(getTasksTotalWeight(activeEmployee.tasks) - 1.0) < 0.001 ? '符合 100%' : '未达 100%' }})
            </span>
            
            <button 
              type="button" 
              class="btn btn-primary" 
              @click="downloadSingle(activeEmployee)"
            >
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
                    <input 
                      type="number" 
                      step="0.05" 
                      min="0" 
                      max="1"
                      v-model.number="task.weight" 
                      class="grid-input text-center bold-text" 
                      @change="handleCellChange(activeEmployee)"
                    />
                  </td>

                  <!-- 所属板块 -->
                  <td>
                    <input 
                      type="text" 
                      v-model="task.category" 
                      class="grid-input" 
                      @change="handleCellChange(activeEmployee)"
                    />
                  </td>

                  <!-- 解释说明 -->
                  <td>
                    <textarea 
                      v-model="task.description" 
                      class="grid-textarea" 
                      @change="handleCellChange(activeEmployee)"
                    ></textarea>
                  </td>

                  <!-- 时间目标 -->
                  <td>
                    <input 
                      type="text" 
                      v-model="task.time_target" 
                      class="grid-input" 
                      @change="handleCellChange(activeEmployee)"
                    />
                  </td>

                  <!-- 数量目标 -->
                  <td>
                    <input 
                      type="text" 
                      v-model="task.count_target" 
                      class="grid-input text-center" 
                      @change="handleCellChange(activeEmployee)"
                    />
                  </td>

                  <!-- 质量目标 -->
                  <td class="highlight-objective">
                    <textarea 
                      v-model="task.quality_target" 
                      class="grid-textarea font-s" 
                      @change="handleCellChange(activeEmployee)"
                    ></textarea>
                  </td>

                  <!-- 时间标准 -->
                  <td>
                    <textarea 
                      v-model="task.time_standard" 
                      class="grid-textarea font-s" 
                      @change="handleCellChange(activeEmployee)"
                    ></textarea>
                  </td>

                  <!-- 数量标准 -->
                  <td>
                    <input 
                      type="text" 
                      v-model="task.count_standard" 
                      class="grid-input text-center" 
                      @change="handleCellChange(activeEmployee)"
                    />
                  </td>

                  <!-- 质量标准 -->
                  <td class="highlight-objective">
                    <textarea 
                      v-model="task.quality_standard" 
                      class="grid-textarea font-s" 
                      @change="handleCellChange(activeEmployee)"
                    ></textarea>
                  </td>

                  <!-- 删除单行 -->
                  <td class="text-center">
                    <button 
                      type="button" 
                      class="grid-action-delete" 
                      @click="removeTask(activeEmployee, index)"
                      :disabled="activeEmployee.tasks.length <= 4"
                      title="删除该任务行"
                    >
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
            <h3>绩效计划自动生成中，并排除主观文字...</h3>
            <p>Gemini 正在严格审查质量目标及质量标准，以 1. 2. 3. 的客观量化指标呈现，请稍等数秒。</p>
          </div>

          <!-- 出错状态 -->
          <div v-else-if="activeEmployee.status === 'error'" class="grid-placeholder error-box">
            <AlertTriangle :size="48" class="text-danger" />
            <h3>推导失败</h3>
            <p class="text-danger-detail">{{ activeEmployee.errorMessage }}</p>
            <button type="button" class="btn btn-primary" @click="generateSingle(activeEmployee)">
              重新生成
            </button>
          </div>
        </div>
      </div>

      <!-- 全局批量打包控制栏 -->
      <div 
        class="footer-actions-bar" 
        v-if="formalEmployees.length > 0 && formalEmployees.some(e => e.status === 'success')"
      >
        <div class="left-stats">
          已成功推导 <strong>{{ formalEmployees.filter(e => e.status === 'success').length }}</strong> / {{ formalEmployees.length }} 份正式绩效计划书
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

.card-header h2 {
  margin: 0;
  font-size: 1.25rem;
  font-weight: 600;
  color: var(--text-h);
  flex-grow: 1;
  text-align: left;
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
  background: rgba(255,255,255,0.02);
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
  background: rgba(255,255,255,0.06);
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

.emp-tab-badge.idle { background: var(--text); }
.emp-tab-badge.generating { background: var(--accent); animation: pulse 1.5s infinite; }
.emp-tab-badge.success { background: #10b981; }
.emp-tab-badge.error { background: #ef4444; }

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

.w-30 { width: calc(30% - 10px); }
.w-40 { width: calc(40% - 10px); }

@media (max-width: 900px) {
  .w-30, .w-40 {
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
  background: rgba(255,255,255,0.06);
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
  background: rgba(255,255,255,0.005);
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
  background: rgba(255,255,255,0.002);
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
  background: rgba(255,255,255,0.01);
}

/* 列宽微调 - 加宽以满足大屏编辑与预览需求 */
.col-seq { width: 45px; text-align: center; }
.col-type { width: 85px; }
.col-level { width: 140px; }
.col-weight { width: 100px; }
.col-cat { width: 110px; }
.col-desc { min-width: 220px; } /* 原 180px */
.col-time { width: 150px; }      /* 原 105px */
.col-count { width: 65px; }
.col-quality { min-width: 300px; } /* 原 200px */
.col-time-std { min-width: 130px; } /* 原 150px */
.col-count-std { width: 85px; }
.col-quality-std { min-width: 300px; } /* 原 200px */
.col-action { width: 45px; text-align: center; }

.cell-seq-num {
  text-align: center;
  background: rgba(255,255,255,0.015);
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
  background: rgba(255,255,255,0.02);
}

.grid-input:focus {
  border-color: var(--accent);
  background: rgba(255,255,255,0.06);
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

.grid-textarea:hover {
  border-color: rgba(255, 255, 255, 0.15);
  background: rgba(255,255,255,0.02);
}

.grid-textarea:focus {
  border-color: var(--accent);
  background: rgba(255,255,255,0.06);
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
  to { transform: rotate(360deg); }
}

@keyframes pulse {
  0%, 100% { transform: scale(1); opacity: 1; }
  50% { transform: scale(1.1); opacity: 0.5; }
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
</style>
