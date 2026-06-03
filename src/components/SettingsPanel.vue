<script setup lang="ts">
/**
 * @fileoverview API 配置设置面板组件。
 * 提供 API Key、代理接口、模型及系统设定（Prompt）的配置与保存功能。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref } from 'vue';
import { geminiConfig } from '../store';
import { Settings, ShieldAlert, Eye, EyeOff, Save, Check, ChevronDown, ChevronUp } from 'lucide-vue-next';

// 控制配置面板的展开与收起状态
// 若本地已存在 API Key，则默认折叠以节省主面板空间；若未配置则展开引导用户输入
const isCollapsed = ref(!!geminiConfig.apiKey);

// 控制 API Key 显隐状态
const showKey = ref(false);
// 控制保存反馈动效
const savedFeedback = ref(false);

// 模型选择项
const modelSelect = ref(
  ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-1.5-pro'].includes(geminiConfig.model)
    ? geminiConfig.model
    : 'custom'
);

/**
 * 保存配置并触发短暂成功勾选效果
 */
function handleSave() {
  savedFeedback.value = true;
  setTimeout(() => {
    savedFeedback.value = false;
  }, 2000);
}

/**
 * 模型选择变更回调
 */
function handleModelSelectChange() {
  if (modelSelect.value !== 'custom') {
    geminiConfig.model = modelSelect.value;
  }
  handleSave();
}
</script>

<template>
  <div class="glass-card settings-panel">
    <!-- 点击头部可控制配置面板的折叠与展开 -->
    <div class="card-header" @click="isCollapsed = !isCollapsed" style="cursor: pointer; user-select: none;">
      <Settings class="header-icon" />
      <h2>AI 模型与配置设定</h2>
      <div class="collapse-trigger">
        <span class="status-summary" v-if="isCollapsed">
          模型: {{ geminiConfig.model }} | {{ geminiConfig.proxyUrl ? '代理已启用' : '官方直连' }}
        </span>
        <ChevronDown v-if="isCollapsed" :size="16" class="arrow-icon" />
        <ChevronUp v-else :size="16" class="arrow-icon" />
      </div>
    </div>

    <transition name="collapse">
      <div class="card-body" v-show="!isCollapsed">
      <!-- 隐私警示框 -->
      <div class="alert-box warning">
        <ShieldAlert class="alert-icon" />
        <div class="alert-content">
          <strong>隐私安全提示：</strong>
          本系统为纯前端实现，您的 API Key 和绩效数据仅会存在浏览器本地 `localStorage` 或直接发送至 Gemini 官方接口，绝不会上传至任何中转服务器，请放心使用。
        </div>
      </div>

      <!-- 表单配置项 -->
      <div class="form-group">
        <label for="apiKey">
          Gemini API Key
          <span class="required">*</span>
        </label>
        <div class="input-with-icon">
          <input
            id="apiKey"
            :type="showKey ? 'text' : 'password'"
            v-model="geminiConfig.apiKey"
            placeholder="AIzaSy..."
            class="form-control"
            @change="handleSave"
          />
          <button type="button" class="icon-btn" @click="showKey = !showKey" title="显示/隐藏密钥">
            <Eye v-if="!showKey" :size="18" />
            <EyeOff v-else :size="18" />
          </button>
        </div>
        <span class="help-text">
          请输入您的 Google AI Studio 密钥。您可以前往
          <a href="https://aistudio.google.com/" target="_blank" class="link-btn">Google AI Studio</a>
          免费获取。
        </span>
      </div>

      <div class="form-row">
        <div class="form-group flex-1">
          <label for="proxyUrl">API 代理网关地址</label>
          <input
            id="proxyUrl"
            type="url"
            v-model="geminiConfig.proxyUrl"
            placeholder="https://generativelanguage.googleapis.com"
            class="form-control"
            @change="handleSave"
          />
          <span class="help-text">国内用户如果无法直连，可填写自定义反向代理网关。若为空则默认直连官方服务。</span>
          <!-- 安全防范警示：警告用户关于第三方中转代理的数据安全风险 -->
          <span class="help-text warning-text" v-if="geminiConfig.proxyUrl">
            ⚠️ 警告：使用非官方代理会经过其服务器，请确保代理可信，以防泄露您的 API 密钥及敏感绩效数据。
          </span>
        </div>
        <div class="form-group flex-1">
          <label for="modelSelect">AI 模型版本</label>
          <select id="modelSelect" v-model="modelSelect" class="form-control select-control" @change="handleModelSelectChange">
            <option value="gemini-2.5-flash">Gemini 2.5 Flash (推荐，新版低成本)</option>
            <option value="gemini-2.5-pro">Gemini 2.5 Pro (推理更强，适合复杂指标)</option>
            <option value="gemini-1.5-pro">Gemini 1.5 Pro (旧版强推理)</option>
            <option value="custom">自定义模型名称...</option>
          </select>
          <span class="help-text">推荐选用最新的 2.5 Flash，响应极其迅速。</span>
        </div>
      </div>

      <!-- 自定义模型名称输入 -->
      <div class="form-row animate-fade-in" v-if="modelSelect === 'custom'">
        <div class="form-group flex-1">
          <label for="customModel">自定义模型标示符</label>
          <input
            id="customModel"
            type="text"
            v-model="geminiConfig.model"
            placeholder="例如 gemini-3.5-flash"
            class="form-control"
            @change="handleSave"
          />
          <span class="help-text">手动输入最新的 Gemini 模型名称，例如 gemini-3.5-flash 。</span>
        </div>
      </div>

      <div class="form-group">
        <label for="systemInstruction">全局考核基准 (System Prompt)</label>
        <textarea
          id="systemInstruction"
          v-model="geminiConfig.systemInstruction"
          rows="3"
          placeholder="例如：你是一个专业、严谨的企业HR绩效评定官..."
          class="form-control textarea-control"
          @change="handleSave"
        ></textarea>
        <span class="help-text">用于定义 AI 在评估过程中的角色、语气以及遵循的通用企业评级规则。</span>
      </div>

      <div class="form-actions">
        <button type="button" class="btn btn-primary btn-save" @click="handleSave">
          <Check v-if="savedFeedback" :size="18" />
          <Save v-else :size="18" />
          <span>{{ savedFeedback ? '已自动保存' : '保存配置' }}</span>
        </button>
      </div>
    </div>
    </transition>
  </div>
</template>

<style scoped>
.settings-panel {
  margin-bottom: 24px;
  animation: fadeIn 0.5s ease-out;
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

/* 警示提示框 */
.alert-box {
  display: flex;
  gap: 12px;
  padding: 14px 18px;
  border-radius: 8px;
  font-size: 0.875rem;
  line-height: 1.5;
}

.alert-box.warning {
  background: rgba(245, 158, 11, 0.08);
  border: 1px solid rgba(245, 158, 11, 0.3);
  color: #d97706;
}

.alert-icon {
  flex-shrink: 0;
  margin-top: 2px;
}

/* 表单布局 */
.form-group {
  display: flex;
  flex-direction: column;
  gap: 6px;
  text-align: left;
}

.form-row {
  display: flex;
  gap: 16px;
}

.flex-1 {
  flex: 1;
}

@media (max-width: 768px) {
  .form-row {
    flex-direction: column;
  }
}

label {
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--text-h);
}

.required {
  color: #ef4444;
  margin-left: 2px;
}

.form-control {
  background: rgba(255, 255, 255, 0.05);
  border: 1px solid var(--border);
  color: var(--text-h);
  padding: 10px 14px;
  border-radius: 6px;
  font-size: 0.95rem;
  outline: none;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  box-sizing: border-box;
}

.form-control:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px var(--accent-bg);
  background: rgba(255, 255, 255, 0.08);
}

.select-control {
  appearance: none;
  background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%239ca3af' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: right 14px center;
  background-size: 16px;
  padding-right: 40px;
}

.textarea-control {
  resize: vertical;
  font-family: inherit;
}

.input-with-icon {
  position: relative;
  display: flex;
  align-items: center;
}

.input-with-icon .form-control {
  width: 100%;
  padding-right: 44px;
}

.icon-btn {
  position: absolute;
  right: 6px;
  background: transparent;
  border: none;
  color: var(--text);
  cursor: pointer;
  padding: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
  transition: color 0.2s, background-color 0.2s;
}

.icon-btn:hover {
  color: var(--text-h);
  background: rgba(255, 255, 255, 0.1);
}

.help-text {
  font-size: 0.775rem;
  color: var(--text);
  margin-top: 2px;
}

.link-btn {
  color: var(--accent);
  text-decoration: none;
  font-weight: 500;
  border-bottom: 1px dashed var(--accent-border);
  transition: color 0.2s;
}

.link-btn:hover {
  color: var(--text-h);
  border-bottom-color: var(--text-h);
}

.form-actions {
  display: flex;
  justify-content: flex-end;
  margin-top: 8px;
}

.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  font-size: 0.9rem;
  font-weight: 600;
  padding: 10px 20px;
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

.btn-save {
  min-width: 150px;
}

/* 折叠展开触发器及动画效果 */
.collapse-trigger {
  margin-left: auto;
  display: flex;
  align-items: center;
  gap: 10px;
}

.status-summary {
  font-size: 0.775rem;
  color: var(--text);
  font-weight: 400;
  background: rgba(255, 255, 255, 0.03);
  padding: 4px 10px;
  border-radius: 20px;
  border: 1px solid var(--border);
}

.arrow-icon {
  color: var(--text);
  transition: transform 0.3s;
}

/* 代理风险警告提示文字样式 */
.warning-text {
  color: #f59e0b !important;
  margin-top: 4px;
  display: block;
}

/* 展开收起过渡动画 */
.collapse-enter-active,
.collapse-leave-active {
  transition: max-height 0.3s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.3s ease-out;
  overflow: hidden;
  max-height: 500px;
}

.collapse-enter-from,
.collapse-leave-to {
  max-height: 0;
  opacity: 0;
  padding-top: 0 !important;
  padding-bottom: 0 !important;
  margin-top: 0 !important;
  margin-bottom: 0 !important;
}

/* 下拉框选项美化，确保在各类暗色背景浏览器中清晰易读 */
.select-control option {
  background-color: #1e293b;
  color: #f8fafc;
  padding: 10px;
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
