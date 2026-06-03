<script setup lang="ts">
/**
 * @fileoverview 批量绩效生成与 Excel 导出组件。
 * 提供可视化的员工数据网格、Excel 复制粘贴快速导入、Gemini 批量 API 调度生成、以及带格式 Excel 回写下载。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref, computed } from 'vue';
import { 
  geminiConfig, 
  cellMappings, 
  excelTemplate, 
  employees, 
  addEmployee, 
  removeEmployee, 
  clearEmployees,
  resetGenerationStatus
} from '../store';
import type { EmployeeRow } from '../store';
import { generateEmployeePerformance } from '../utils/geminiHelper';
import { writeDataToTemplate, downloadExcelFile } from '../utils/excelHelper';
import { 
  Users, 
  UserPlus, 
  Trash2, 
  Sparkles, 
  Download, 
  AlertTriangle, 
  CheckCircle,
  Clock,
  ClipboardPaste,
  FileDown,
  RefreshCw
} from 'lucide-vue-next';

// 复制粘贴快捷导入输入框的值
const importText = ref('');
// 是否显示导入抽屉/对话框
const showImportModal = ref(false);
// 是否正在进行批量生成
const isGeneratingBatch = ref(false);

// 提取当前所有的已知输入映射和 AI 生成映射
const inputMappings = computed(() => cellMappings.value.filter(m => m.type === 'input'));
const aiMappings = computed(() => cellMappings.value.filter(m => m.type === 'ai'));

// 表头定义
const tableHeaders = computed(() => {
  return [
    ...inputMappings.value.map(m => ({ ref: m.cellRef, label: m.label, type: 'input' })),
    ...aiMappings.value.map(m => ({ ref: m.cellRef, label: m.label, type: 'ai' }))
  ];
});

// 表单验证状态
const validationError = computed(() => {
  if (!excelTemplate.buffer) {
    return '请先在上方上传您的 Excel 绩效模板。';
  }
  if (!geminiConfig.apiKey) {
    return '请先在设置中填写 Gemini API Key。';
  }
  if (employees.value.length === 0) {
    return '请添加至少一名员工的数据。';
  }
  return '';
});

/**
 * 从剪贴板/文本域 TSV 格式数据解析并导入到表格中
 * 方便用户直接从公司的另一个 Excel 表格中复制一整列/一整行数据直接粘贴
 */
function handlePasteImport() {
  if (!importText.value.trim()) return;

  const lines = importText.value.trim().split('\n');
  let importedCount = 0;

  lines.forEach(line => {
    // 按 Tab 键或逗号分割字段
    const cells = line.split(/\t|,/);
    if (cells.length === 0 || (cells.length === 1 && cells[0].trim() === '')) return;

    // 创建新员工行
    const newRow: EmployeeRow = {
      id: crypto.randomUUID(),
      inputs: {},
      aiOutputs: {},
      status: 'idle'
    };

    // 填充输入字段
    inputMappings.value.forEach((m, idx) => {
      // 若复制的数据列数不够，使用空字符串填充
      newRow.inputs[m.cellRef] = cells[idx] ? cells[idx].trim() : '';
    });

    // 初始化 AI 字段为空
    aiMappings.value.forEach(m => {
      newRow.aiOutputs[m.cellRef] = '';
    });

    employees.value.push(newRow);
    importedCount++;
  });

  importText.value = '';
  showImportModal.value = false;
}

/**
 * 触发单行员工绩效的 AI 生成
 */
async function generateSingle(row: EmployeeRow) {
  if (!geminiConfig.apiKey) return;
  row.status = 'generating';
  row.errorMessage = undefined;

  // 1. 构建 AI 输入上下文
  const inputContext: Record<string, string | number> = {};
  inputMappings.value.forEach(m => {
    inputContext[m.label] = row.inputs[m.cellRef] || '';
  });

  // 2. 调用 Gemini API 生成内容
  const result = await generateEmployeePerformance(
    {
      apiKey: geminiConfig.apiKey,
      proxyUrl: geminiConfig.proxyUrl,
      model: geminiConfig.model,
      systemInstruction: geminiConfig.systemInstruction
    },
    inputContext,
    aiMappings.value
  );

  // 3. 处理生成结果
  if (result.success) {
    row.aiOutputs = result.cellData;
    row.status = 'success';

    // 4. 在内存中预渲染该用户的 Excel 文件
    try {
      const nameField = inputMappings.value.find(m => m.label.includes('名') || m.fieldName === 'name');
      const nameVal = nameField ? String(row.inputs[nameField.cellRef] || '未命名') : '员工';
      row.fileName = `绩效表_${nameVal}_${new Date().toISOString().slice(0, 7)}.xlsx`;

      // 组装回写数据 (合并 input 和 aiOutputs)
      const mergedData: Record<string, string | number> = {};
      Object.entries(row.inputs).forEach(([k, v]) => mergedData[k] = v);
      Object.entries(row.aiOutputs).forEach(([k, v]) => mergedData[k] = v);

      if (excelTemplate.buffer) {
        row.outputBuffer = await writeDataToTemplate(excelTemplate.buffer, mergedData);
      }
    } catch (err: any) {
      row.errorMessage = `回写 Excel 格式失败: ${err.message}`;
      row.status = 'error';
    }
  } else {
    row.errorMessage = result.message;
    row.status = 'error';
  }
}

/**
 * 批量执行所有未成功生成的员工行
 */
async function generateAll() {
  if (validationError.value) return;

  isGeneratingBatch.value = true;
  
  // 仅筛选需要生成的行
  const pendingRows = employees.value.filter(emp => emp.status !== 'success');
  
  // 串行生成，避免触发 API 并发限流 (Rate Limit)
  for (const row of pendingRows) {
    if (!isGeneratingBatch.value) break; // 如果点击了停止/中断
    await generateSingle(row);
  }

  isGeneratingBatch.value = false;
}

/**
 * 停止当前的批量生成
 */
function stopGeneration() {
  isGeneratingBatch.value = false;
}

/**
 * 触发单个员工 Excel 文件下载
 */
function downloadSingle(row: EmployeeRow) {
  if (!row.outputBuffer || !row.fileName) return;
  downloadExcelFile(row.outputBuffer, row.fileName);
}

/**
 * 批量下载所有已成功生成的 Excel 文件
 */
function downloadAll() {
  const successRows = employees.value.filter(emp => emp.status === 'success' && emp.outputBuffer && emp.fileName);
  if (successRows.length === 0) return;

  // 逐个触发浏览器下载（浏览器会并发弹出下载，确保用户允许了多文件下载）
  successRows.forEach((row, index) => {
    // 延迟几毫秒触发，防止浏览器拦截并发下载
    setTimeout(() => {
      if (row.outputBuffer && row.fileName) {
        downloadExcelFile(row.outputBuffer, row.fileName);
      }
    }, index * 200);
  });
}
</script>

<template>
  <div class="glass-card batch-generator">
    <div class="card-header">
      <Users class="header-icon" />
      <h2>批量员工绩效列表</h2>
      <div class="header-actions">
        <!-- 快速导入按钮 -->
        <button type="button" class="btn btn-outline" @click="showImportModal = true">
          <ClipboardPaste :size="16" />
          <span>从 Excel 复制导入</span>
        </button>
        <!-- 添加空白行按钮 -->
        <button type="button" class="btn btn-outline" @click="addEmployee">
          <UserPlus :size="16" />
          <span>添加员工</span>
        </button>
      </div>
    </div>

    <div class="card-body">
      <!-- 错误提示带 -->
      <div v-if="validationError" class="validation-warning">
        <AlertTriangle :size="20" class="warning-icon" />
        <span>{{ validationError }}</span>
      </div>

      <!-- 网格数据表格 -->
      <div class="grid-table-wrapper" v-if="employees.length > 0">
        <table class="grid-table">
          <thead>
            <tr>
              <th class="col-index">#</th>
              <th class="col-status">生成状态</th>
              <th v-for="h in tableHeaders" :key="h.ref" :class="{ 'ai-th': h.type === 'ai' }">
                <span class="th-title">{{ h.label }}</span>
                <span class="th-ref">[{{ h.ref }}]</span>
              </th>
              <th class="col-row-actions">操作</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(emp, idx) in employees" :key="emp.id" :class="`row-${emp.status}`">
              <!-- 行号 -->
              <td class="col-index">{{ idx + 1 }}</td>
              
              <!-- 状态标志 -->
              <td class="col-status">
                <div class="status-badge" :class="emp.status">
                  <Clock v-if="emp.status === 'idle'" :size="14" />
                  <div v-else-if="emp.status === 'generating'" class="spinner-small"></div>
                  <CheckCircle v-else-if="emp.status === 'success'" :size="14" />
                  <AlertTriangle v-else-if="emp.status === 'error'" :size="14" />
                  <span>
                    {{ emp.status === 'idle' ? '等待生成' : emp.status === 'generating' ? 'AI评估中...' : emp.status === 'success' ? '生成成功' : '生成失败' }}
                  </span>
                </div>
                <div v-if="emp.errorMessage" class="row-error-message" :title="emp.errorMessage">
                  {{ emp.errorMessage }}
                </div>
              </td>

              <!-- 输入字段列 -->
              <td v-for="m in inputMappings" :key="m.cellRef">
                <input
                  type="text"
                  v-model="emp.inputs[m.cellRef]"
                  placeholder="请输入内容"
                  class="cell-input"
                  :disabled="emp.status === 'generating' || isGeneratingBatch"
                />
              </td>

              <!-- AI 生成字段列（只读，且带有星光动态效果） -->
              <td v-for="m in aiMappings" :key="m.cellRef" class="ai-td">
                <textarea
                  v-model="emp.aiOutputs[m.cellRef]"
                  class="cell-textarea"
                  placeholder="AI 自动生成处"
                  readonly
                ></textarea>
              </td>

              <!-- 操作行 -->
              <td class="col-row-actions">
                <div class="row-actions-group">
                  <button
                    type="button"
                    class="action-btn download-btn"
                    :disabled="emp.status !== 'success'"
                    @click="downloadSingle(emp)"
                    title="下载此员工的绩效 Excel 文件"
                  >
                    <Download :size="16" />
                  </button>
                  <button
                    type="button"
                    class="action-btn delete-btn"
                    @click="removeEmployee(emp.id)"
                    title="删除此行"
                    :disabled="isGeneratingBatch"
                  >
                    <Trash2 :size="16" />
                  </button>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- 空白列表状态 -->
      <div v-else class="empty-state">
        <Users :size="48" class="empty-icon" />
        <h3>暂无员工绩效数据</h3>
        <p>您可以点击右上角的“添加员工”手动填写，或者点击“从 Excel 复制导入”直接批量导入数据。</p>
        <button type="button" class="btn btn-primary" style="margin-top: 10px;" @click="addEmployee">
          <UserPlus :size="16" />
          <span>添加第一名员工</span>
        </button>
      </div>

      <!-- 底部全局操作条 -->
      <div class="footer-actions-bar" v-if="employees.length > 0">
        <div class="left-stats">
          共 <strong>{{ employees.length }}</strong> 名员工 | 
          已生成 <strong class="success-text">{{ employees.filter(e => e.status === 'success').length }}</strong> 个
        </div>
        <div class="right-buttons">
          <button 
            type="button" 
            class="btn btn-outline" 
            @click="clearEmployees" 
            :disabled="isGeneratingBatch"
          >
            清空列表
          </button>
          <button 
            type="button" 
            class="btn btn-outline" 
            @click="resetGenerationStatus" 
            :disabled="isGeneratingBatch"
          >
            <RefreshCw :size="16" />
            <span>重置状态</span>
          </button>
          <button
            v-if="employees.filter(e => e.status === 'success').length > 0"
            type="button"
            class="btn btn-outline success-outline-btn"
            @click="downloadAll"
          >
            <FileDown :size="16" />
            <span>一键下载所有成功 Excel</span>
          </button>
          
          <button
            v-if="!isGeneratingBatch"
            type="button"
            class="btn btn-primary"
            :disabled="!!validationError"
            @click="generateAll"
          >
            <Sparkles :size="16" />
            <span>开始批量生成绩效</span>
          </button>
          <button
            v-else
            type="button"
            class="btn btn-danger"
            @click="stopGeneration"
          >
            <span>停止生成</span>
          </button>
        </div>
      </div>
    </div>

    <!-- 复制粘贴导入弹窗 -->
    <div class="modal-mask" v-if="showImportModal">
      <div class="modal-container glass-card">
        <div class="modal-header">
          <h3>从外部 Excel 批量粘贴导入数据</h3>
        </div>
        <div class="modal-body">
          <p class="modal-help">
            在您的 Excel 表格中，选择要填写的输入项列（例如：<strong>员工姓名、部门、工作总结</strong> 等列），复制（Ctrl+C），并在下方文本框中直接粘贴（Ctrl+V），系统会自动按行和列解析导入。
          </p>
          <textarea
            v-model="importText"
            rows="10"
            placeholder="请在此处直接粘贴从 Excel 复制的内容..."
            class="form-control import-textarea"
          ></textarea>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-outline" @click="showImportModal = false">取消</button>
          <button type="button" class="btn btn-primary" @click="handlePasteImport" :disabled="!importText.trim()">
            确认导入
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.batch-generator {
  animation: fadeIn 0.5s ease-out 0.2s both;
}

.card-header {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 20px 24px;
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
  gap: 10px;
}

.card-body {
  padding: 24px;
  display: flex;
  flex-direction: column;
  gap: 20px;
}

/* 验证警告框 */
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

/* 表格容器样式 */
.grid-table-wrapper {
  overflow-x: auto;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: rgba(255,255,255,0.01);
}

.grid-table {
  width: 100%;
  border-collapse: collapse;
  text-align: left;
  font-size: 0.875rem;
}

.grid-table th, 
.grid-table td {
  padding: 10px 12px;
  border-bottom: 1px solid var(--border);
  border-right: 1px solid var(--border);
  vertical-align: middle;
  min-width: 120px;
}

.grid-table th:last-child, 
.grid-table td:last-child {
  border-right: none;
}

.grid-table th {
  background: rgba(255, 255, 255, 0.03);
  color: var(--text-h);
  font-weight: 600;
  position: sticky;
  top: 0;
  z-index: 10;
}

.th-title {
  display: block;
}

.th-ref {
  display: block;
  font-size: 0.725rem;
  color: var(--text);
  font-weight: 400;
}

.ai-th {
  background: rgba(170, 59, 255, 0.04) !important;
}

.col-index {
  width: 40px !important;
  min-width: 40px !important;
  text-align: center;
}

.col-status {
  width: 120px !important;
  min-width: 120px !important;
}

.col-row-actions {
  width: 90px !important;
  min-width: 90px !important;
  text-align: center;
}

/* 输入控件样式 */
.cell-input {
  background: transparent;
  border: 1px solid transparent;
  color: var(--text-h);
  padding: 6px 8px;
  width: 100%;
  box-sizing: border-box;
  font-size: 0.875rem;
  border-radius: 4px;
  outline: none;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.cell-input:hover {
  border-color: rgba(255, 255, 255, 0.1);
  background: rgba(255,255,255,0.02);
}

.cell-input:focus {
  border-color: var(--accent);
  background: rgba(255,255,255,0.06);
  box-shadow: 0 0 0 2px var(--accent-bg);
}

.ai-td {
  background: rgba(170, 59, 255, 0.01);
}

.cell-textarea {
  background: transparent;
  border: 1px solid transparent; /* 统一边框风格 */
  color: var(--text-h);
  padding: 6px 8px;
  width: 100%;
  min-height: 50px;
  resize: vertical;
  box-sizing: border-box;
  font-size: 0.825rem;
  font-family: inherit;
  outline: none;
  border-radius: 4px;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.cell-textarea:hover {
  border-color: rgba(255, 255, 255, 0.1);
  background: rgba(255,255,255,0.02);
}

.cell-textarea:focus {
  border-color: var(--accent);
  background: rgba(255,255,255,0.06);
  box-shadow: 0 0 0 2px var(--accent-bg);
}

/* 状态徽章 */
.status-badge {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 4px 8px;
  border-radius: 20px;
  font-size: 0.75rem;
  font-weight: 600;
}

.status-badge.idle {
  background: rgba(255,255,255,0.06);
  color: var(--text);
}

.status-badge.generating {
  background: rgba(170, 59, 255, 0.15);
  color: var(--accent);
}

.status-badge.success {
  background: rgba(16, 185, 129, 0.15);
  color: #10b981;
}

.status-badge.error {
  background: rgba(239, 68, 68, 0.15);
  color: #ef4444;
}

.row-error-message {
  font-size: 0.725rem;
  color: #ef4444;
  margin-top: 4px;
  max-width: 120px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  cursor: help;
}

/* 小转轮 */
.spinner-small {
  width: 12px;
  height: 12px;
  border: 2px solid rgba(255, 255, 255, 0.1);
  border-radius: 50%;
  border-top-color: var(--accent);
  animation: spin 1s ease-in-out infinite;
}

/* 行状态动画背景 */
.row-generating {
  background: rgba(170, 59, 255, 0.02);
}

.row-success {
  background: rgba(16, 185, 129, 0.01);
}

/* 操作列组 */
.row-actions-group {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
}

.action-btn {
  background: transparent;
  border: none;
  cursor: pointer;
  padding: 6px;
  border-radius: 4px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s;
}

.download-btn {
  color: #10b981;
}

.download-btn:hover:not(:disabled) {
  background: rgba(16, 185, 129, 0.1);
}

.download-btn:disabled {
  color: var(--text);
  opacity: 0.3;
  cursor: not-allowed;
}

.delete-btn {
  color: var(--text);
}

.delete-btn:hover:not(:disabled) {
  background: rgba(255,255,255,0.06);
  color: #ef4444;
}

.delete-btn:disabled {
  opacity: 0.3;
  cursor: not-allowed;
}

/* 空白状态 */
.empty-state {
  padding: 48px;
  border: 1px dashed var(--border);
  border-radius: 8px;
  background: rgba(255,255,255,0.01);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
}

.empty-icon {
  color: var(--text);
  opacity: 0.5;
}

.empty-state h3 {
  margin: 0;
  font-size: 1.1rem;
  font-weight: 600;
  color: var(--text-h);
}

.empty-state p {
  margin: 0;
  font-size: 0.875rem;
  color: var(--text);
  max-width: 480px;
  line-height: 1.5;
}

/* 底部操作条 */
.footer-actions-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-top: 1px solid var(--border);
  padding-top: 16px;
  margin-top: 8px;
}

.left-stats {
  font-size: 0.85rem;
  color: var(--text);
}

.left-stats strong {
  color: var(--text-h);
}

.right-buttons {
  display: flex;
  gap: 10px;
}

/* 按钮基础及自定义 (符合现有高档玻璃拟态与渐变风格) */
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  font-size: 0.875rem;
  font-weight: 600;
  padding: 8px 18px;
  border-radius: 8px;
  cursor: pointer;
  border: 1px solid rgba(255, 255, 255, 0.08);
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  box-sizing: border-box;
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

.success-outline-btn {
  color: #10b981;
  border-color: rgba(16, 185, 129, 0.3);
  background: rgba(16, 185, 129, 0.02);
}

.success-outline-btn:hover:not(:disabled) {
  background: rgba(16, 185, 129, 0.08) !important;
  border-color: #10b981 !important;
  color: #10b981 !important;
  box-shadow: 0 4px 12px rgba(16, 185, 129, 0.15) !important;
  transform: translateY(-1px);
}

.btn-danger {
  background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
  color: white;
  border: 1px solid rgba(255, 255, 255, 0.1);
  box-shadow: 0 4px 12px rgba(239, 68, 68, 0.2);
}

.btn-danger:hover:not(:disabled) {
  transform: translateY(-2px);
  background: linear-gradient(135deg, #f87171 0%, #e11d48 100%);
  box-shadow: 0 6px 18px rgba(239, 68, 68, 0.35);
}

.btn-danger:active:not(:disabled) {
  transform: translateY(0);
}

/* 导入 Modal 弹窗 */
.modal-mask {
  position: fixed;
  top: 0;
  left: 0;
  width: 100vw;
  height: 100vh;
  background: rgba(0, 0, 0, 0.6);
  backdrop-filter: blur(4px);
  z-index: 100;
  display: flex;
  align-items: center;
  justify-content: center;
}

.modal-container {
  width: 600px;
  max-width: 90%;
  background: var(--bg);
  border: 1px solid var(--border);
  border-radius: 12px;
  box-shadow: var(--shadow);
  display: flex;
  flex-direction: column;
  overflow: hidden;
  animation: modalScale 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
}

@keyframes modalScale {
  from {
    opacity: 0;
    transform: scale(0.9);
  }
  to {
    opacity: 1;
    transform: scale(1);
  }
}

.modal-header {
  padding: 16px 20px;
  border-bottom: 1px solid var(--border);
  text-align: left;
}

.modal-header h3 {
  margin: 0;
  font-size: 1.1rem;
  font-weight: 600;
  color: var(--text-h);
}

.modal-body {
  padding: 20px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  text-align: left;
}

.modal-help {
  margin: 0;
  font-size: 0.825rem;
  color: var(--text);
  line-height: 1.5;
}

.import-textarea {
  width: 100%;
  box-sizing: border-box;
  font-family: var(--mono);
  font-size: 0.8rem;
  line-height: 1.4;
  padding: 12px;
}

.modal-footer {
  padding: 14px 20px;
  border-top: 1px solid var(--border);
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}

@keyframes fadeIn {
  from {
    opacity: 0;
    transform: translateY(10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}
</style>
