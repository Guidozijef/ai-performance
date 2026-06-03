<script setup lang="ts">
/**
 * @fileoverview Excel 模板与字段映射配置组件。
 * 支持绩效 Excel 模板文件的上传，以及定义每个单元格数据（用户输入或 AI 生成）的绑定坐标及规则。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref } from 'vue';
import { excelTemplate, cellMappings } from '../store';
import { loadWorkbook } from '../utils/excelHelper';
import { 
  UploadCloud, 
  Plus, 
  Trash2, 
  Map, 
  Sparkles, 
  Keyboard,
  Info
} from 'lucide-vue-next';

// 记录读取中的 Loading 状态
const loadingTemplate = ref(false);
// 错误提示
const errorMessage = ref('');

/**
 * 处理模板 Excel 文件的上传与预读
 */
async function handleTemplateUpload(event: Event) {
  const target = event.target as HTMLInputElement;
  const files = target.files;
  if (!files || files.length === 0) return;

  const file = files[0];
  loadingTemplate.value = true;
  errorMessage.value = '';

  try {
    // 调用 excelHelper 加载工作簿，校验文件有效性
    const workbook = await loadWorkbook(file);
    const firstSheet = workbook.worksheets[0];
    
    if (!firstSheet) {
      throw new Error('该 Excel 文件中没有任何工作表 (Sheet)。');
    }

    // 将二进制文件内容及文件名存入全局 Store
    const arrayBuffer = await file.arrayBuffer();
    excelTemplate.fileName = file.name;
    excelTemplate.buffer = arrayBuffer;
  } catch (error: any) {
    console.error('解析模板失败:', error);
    errorMessage.value = `解析模板失败: ${error.message || '格式错误'}`;
    excelTemplate.fileName = '';
    excelTemplate.buffer = null;
  } finally {
    loadingTemplate.value = false;
    // 重置 input 以便能重复触发 change 事件
    target.value = '';
  }
}

/**
 * 增加一个单元格映射行
 */
function addMapping() {
  cellMappings.value.push({
    cellRef: '',
    fieldName: `field_${Date.now()}`,
    label: '新考核项',
    type: 'input',
    aiInstruction: ''
  });
}

/**
 * 移除一个单元格映射行
 */
function removeMapping(index: number) {
  cellMappings.value.splice(index, 1);
}
</script>

<template>
  <div class="glass-card template-config">
    <div class="card-header">
      <Map class="header-icon" />
      <h2>Excel 绩效模板与坐标映射设定</h2>
    </div>

    <div class="card-body">
      <!-- 1. 模板上传区域 -->
      <div class="template-upload-zone">
        <label class="upload-label" :class="{ 'has-file': excelTemplate.buffer }">
          <input
            type="file"
            accept=".xlsx"
            class="hidden-input"
            @change="handleTemplateUpload"
          />
          <div class="upload-content" v-if="!loadingTemplate">
            <UploadCloud :size="36" class="upload-icon" />
            <div class="upload-text" v-if="!excelTemplate.fileName">
              <strong>上传公司 Excel 绩效模板 (.xlsx)</strong>
              <span>点击或拖拽文件到此处上传</span>
            </div>
            <div class="upload-text" v-else>
              <strong class="success-text">已加载绩效模板: {{ excelTemplate.fileName }}</strong>
              <span>点击此处重新上传替换</span>
            </div>
          </div>
          <div class="upload-content loading" v-else>
            <div class="spinner"></div>
            <span>正在解析 Excel 模板结构及样式...</span>
          </div>
        </label>
        <div v-if="errorMessage" class="error-text-box">
          {{ errorMessage }}
        </div>
      </div>

      <!-- 2. 说明框 -->
      <div class="alert-box info">
        <Info class="alert-icon" />
        <div class="alert-content">
          <strong>设置说明：</strong>
          请输入模板中需要填写的**单元格准确坐标**（如 B3、C10 等），并标记是“人工输入”（如员工姓名、本月总结）还是“AI 生成”（如评分、评语）。回写时程序将完美保留其他原样公式和文字。
        </div>
      </div>

      <!-- 3. 映射绑定列表 -->
      <div class="mappings-container">
        <div class="mappings-header-row">
          <span class="col-ref">单元格坐标</span>
          <span class="col-label">项目名称 (如姓名/评语)</span>
          <span class="col-type">来源类型</span>
          <span class="col-instruction">说明与 AI 生成指令 (仅 AI 字段生效)</span>
          <span class="col-actions">操作</span>
        </div>

        <div class="mappings-list">
          <div 
            v-for="(mapping, index) in cellMappings" 
            :key="mapping.fieldName" 
            class="mapping-item"
          >
            <!-- 单元格坐标，例如 B3 -->
            <div class="col-ref">
              <input
                type="text"
                v-model="mapping.cellRef"
                placeholder="例如 B3"
                class="form-control text-center uppercase"
                maxlength="5"
              />
            </div>

            <!-- 项目标签，例如 员工姓名 -->
            <div class="col-label">
              <input
                type="text"
                v-model="mapping.label"
                placeholder="例如 员工姓名"
                class="form-control"
              />
            </div>

            <!-- 映射类型：input/ai -->
            <div class="col-type">
              <div class="type-toggle">
                <button
                  type="button"
                  class="toggle-btn"
                  :class="{ active: mapping.type === 'input' }"
                  @click="mapping.type = 'input'"
                  title="用户输入的已知字段"
                >
                  <Keyboard :size="14" />
                  <span>已知输入</span>
                </button>
                <button
                  type="button"
                  class="toggle-btn"
                  :class="{ active: mapping.type === 'ai' }"
                  @click="mapping.type = 'ai'"
                  title="AI 自动生成的字段"
                >
                  <Sparkles :size="14" />
                  <span>AI 生成</span>
                </button>
              </div>
            </div>

            <!-- AI 提示指令 -->
            <div class="col-instruction">
              <input
                v-if="mapping.type === 'ai'"
                type="text"
                v-model="mapping.aiInstruction"
                placeholder="指导 AI 如何生成此项，例如：根据总结打出1-100的数字，只填数字"
                class="form-control instruction-input"
              />
              <span v-else class="text-muted">（作为 AI 绩效生成的参考上下文）</span>
            </div>

            <!-- 操作按钮 -->
            <div class="col-actions">
              <button 
                type="button" 
                class="btn-delete" 
                @click="removeMapping(index)" 
                title="删除映射"
                :disabled="cellMappings.length <= 1"
              >
                <Trash2 :size="16" />
              </button>
            </div>
          </div>
        </div>

        <div class="add-mapping-btn-row">
          <button type="button" class="btn btn-outline" @click="addMapping">
            <Plus :size="16" />
            <span>添加绑定单元格</span>
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.template-config {
  margin-bottom: 24px;
  animation: fadeIn 0.5s ease-out 0.1s both;
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
}

.card-body {
  padding: 24px;
  display: flex;
  flex-direction: column;
  gap: 20px;
}

/* 模板上传区域 */
.template-upload-zone {
  width: 100%;
}

.upload-label {
  display: flex;
  align-items: center;
  justify-content: center;
  border: 2px dashed var(--border);
  border-radius: 8px;
  padding: 24px;
  cursor: pointer;
  background: rgba(255, 255, 255, 0.02);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

.upload-label:hover {
  border-color: var(--accent);
  background: var(--accent-bg);
}

.upload-label.has-file {
  border-color: rgba(16, 185, 129, 0.5);
  background: rgba(16, 185, 129, 0.03);
}

.upload-label.has-file:hover {
  border-color: rgba(16, 185, 129, 0.8);
  background: rgba(16, 185, 129, 0.06);
}

.hidden-input {
  display: none;
}

.upload-content {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
}

.upload-icon {
  color: var(--text);
  transition: color 0.3s;
}

.upload-label:hover .upload-icon {
  color: var(--accent);
}

.upload-label.has-file .upload-icon {
  color: #10b981;
}

.upload-text {
  display: flex;
  flex-direction: column;
  gap: 4px;
  font-size: 0.9rem;
}

.upload-text strong {
  color: var(--text-h);
  font-size: 0.95rem;
}

.upload-text span {
  color: var(--text);
}

.success-text {
  color: #10b981 !important;
}

.error-text-box {
  margin-top: 10px;
  padding: 10px;
  border-radius: 6px;
  background: rgba(239, 68, 68, 0.1);
  border: 1px solid rgba(239, 68, 68, 0.3);
  color: #ef4444;
  font-size: 0.875rem;
}

/* Spinner 动画 */
.spinner {
  width: 28px;
  height: 28px;
  border: 3px solid rgba(255, 255, 255, 0.1);
  border-radius: 50%;
  border-top-color: var(--accent);
  animation: spin 1s ease-in-out infinite;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

/* 提示框 */
.alert-box {
  display: flex;
  gap: 12px;
  padding: 14px 18px;
  border-radius: 8px;
  font-size: 0.875rem;
  line-height: 1.5;
  text-align: left;
}

.alert-box.info {
  background: rgba(59, 130, 246, 0.08);
  border: 1px solid rgba(59, 130, 246, 0.3);
  color: #3b82f6;
}

.alert-icon {
  flex-shrink: 0;
  margin-top: 2px;
}

/* 映射列表容器 */
.mappings-container {
  border: 1px solid var(--border);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.01);
  overflow: hidden;
}

.mappings-header-row {
  display: flex;
  background: rgba(255, 255, 255, 0.03);
  border-bottom: 1px solid var(--border);
  padding: 10px 16px;
  font-size: 0.825rem;
  font-weight: 600;
  color: var(--text-h);
  text-align: left;
}

.mapping-item {
  display: flex;
  align-items: center;
  padding: 12px 16px;
  border-bottom: 1px solid var(--border);
  gap: 14px;
}

.mapping-item:last-child {
  border-bottom: none;
}

/* 列宽分配 */
.col-ref {
  width: 90px;
  flex-shrink: 0;
}

.col-label {
  width: 160px;
  flex-shrink: 0;
}

.col-type {
  width: 170px;
  flex-shrink: 0;
}

.col-instruction {
  flex-grow: 1;
  text-align: left;
}

.col-actions {
  width: 50px;
  flex-shrink: 0;
  display: flex;
  justify-content: center;
}

.text-center {
  text-align: center;
}

.uppercase {
  text-transform: uppercase;
}

.form-control {
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid var(--border);
  color: var(--text-h);
  padding: 8px 12px;
  border-radius: 6px;
  font-size: 0.9rem;
  outline: none;
  width: 100%;
  box-sizing: border-box;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.form-control:hover {
  border-color: rgba(255, 255, 255, 0.15);
  background: rgba(255, 255, 255, 0.05);
}

.form-control:focus {
  border-color: var(--accent);
  background: rgba(255, 255, 255, 0.06);
  box-shadow: 0 0 0 3px var(--accent-bg);
}

.instruction-input {
  font-size: 0.825rem;
}

.text-muted {
  font-size: 0.825rem;
  color: var(--text);
}

/* 按钮切换状态 */
.type-toggle {
  display: flex;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 3px;
  width: 100%;
  box-sizing: border-box;
}

.toggle-btn {
  flex: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  border: none;
  background: transparent;
  color: var(--text);
  font-size: 0.8rem;
  padding: 5px 8px;
  cursor: pointer;
  border-radius: 4px;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.toggle-btn.active {
  background: linear-gradient(135deg, #a855f7 0%, #6366f1 100%);
  color: white;
  box-shadow: 0 2px 8px rgba(168, 85, 247, 0.25);
}

.toggle-btn.active span {
  font-weight: 600;
}

.toggle-btn:hover:not(.active) {
  color: var(--text-h);
  background: rgba(255, 255, 255, 0.03);
}

/* 删除按钮样式与交互 */
.btn-delete {
  background: rgba(239, 68, 68, 0.05);
  border: 1px solid rgba(239, 68, 68, 0.15);
  color: #ef4444;
  cursor: pointer;
  padding: 8px;
  border-radius: 6px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.btn-delete:hover:not(:disabled) {
  background: rgba(239, 68, 68, 0.15);
  border-color: #ef4444;
  transform: scale(1.05);
}

.btn-delete:active:not(:disabled) {
  transform: scale(0.95);
}

.btn-delete:disabled {
  opacity: 0.3;
  cursor: not-allowed;
  background: transparent;
  border-color: transparent;
}

/* 添加按钮行 */
.add-mapping-btn-row {
  padding: 12px 16px;
  background: rgba(255, 255, 255, 0.01);
  display: flex;
  justify-content: flex-start;
}

/* 虚线边框大按钮 */
.btn-outline {
  background: rgba(255, 255, 255, 0.02);
  border: 1px dashed var(--border);
  color: var(--text-h);
  padding: 10px 18px;
  font-size: 0.875rem;
  border-radius: 8px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.btn-outline:hover {
  border-color: var(--accent);
  color: var(--text-h);
  background: var(--accent-bg);
  box-shadow: 0 4px 12px rgba(168, 85, 247, 0.1);
  transform: translateY(-1px);
}

.btn-outline:active {
  transform: translateY(0);
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

@media (max-width: 900px) {
  .mappings-header-row {
    display: none;
  }
  .mapping-item {
    flex-wrap: wrap;
    gap: 8px;
    padding: 16px;
  }
  .col-ref, .col-label, .col-type {
    width: calc(50% - 4px);
  }
  .col-instruction {
    width: 100%;
  }
  .col-actions {
    width: 100%;
    justify-content: flex-end;
  }
}
</style>
