<script setup lang="ts">
/**
 * @fileoverview API 配置设置面板组件。
 * 提供多大模型（Google Gemini、DeepSeek、通义千问、ChatGPT、智谱 AI 及自定义 URL）的统一配置、服务商切换、独立 Key 记忆、动态模型拉取与系统设定。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref, computed, watch } from "vue";
import { geminiConfig, performanceMonth, availableModels, LLM_PROVIDERS, type LLMProvider } from "../store";
import { fetchAvailableModels } from "../utils/geminiHelper";
import { Settings, ShieldAlert, Eye, EyeOff, Save, Check, ChevronDown, ChevronUp, RefreshCw, XCircle, ExternalLink, Globe, Cpu, RotateCcw } from "lucide-vue-next";

// 控制配置面板的展开与收起状态
// 默认展开以直观呈现 Google Gemini 及多模型配置信息，满足“默认展示 Gemini”要求
const isCollapsed = ref(false);

// 控制 API Key 显隐明文状态
const showKey = ref(false);

// 控制保存成功反馈动画提示
const savedFeedback = ref(false);

// 模型请求加载状态
const isLoadingModels = ref(false);
const loadModelsError = ref<string | null>(null);
const loadModelsSuccess = ref(false);

// 当前选中的预置服务商详情
const currentPreset = computed(() => {
  return LLM_PROVIDERS[geminiConfig.provider] || LLM_PROVIDERS.gemini;
});

// 判断当前接口地址是否为自定义修改后的地址
const isCustomUrl = computed(() => {
  if (!currentPreset.value.defaultBaseUrl) return !!geminiConfig.baseUrl;
  return !!geminiConfig.baseUrl && geminiConfig.baseUrl.trim() !== currentPreset.value.defaultBaseUrl.trim();
});

// 服务商列表（用于生成选择网格）
const providerList = computed(() => {
  return Object.values(LLM_PROVIDERS);
});

// 初始化模型下拉框选中项
const initialSelect = () => {
  const currentModel = geminiConfig.model || currentPreset.value.defaultModel || 'gemini-2.5-flash';
  if (currentPreset.value.presetModels.includes(currentModel)) {
    return currentModel;
  }
  if (availableModels.value.includes(currentModel)) {
    return currentModel;
  }
  if (geminiConfig.provider === 'gemini') {
    return 'gemini-2.5-flash';
  }
  return currentModel || 'custom';
};
const modelSelect = ref(initialSelect());

// 监听当前模型变化，同步更新下拉框
watch(
  () => geminiConfig.model,
  (newModel) => {
    if (currentPreset.value.presetModels.includes(newModel) || availableModels.value.includes(newModel)) {
      modelSelect.value = newModel;
    } else {
      modelSelect.value = "custom";
    }
  },
);

/**
 * 保存配置并触发反馈动效
 */
function handleSave() {
  geminiConfig.proxyUrl = geminiConfig.baseUrl;
  savedFeedback.value = true;
  setTimeout(() => {
    savedFeedback.value = false;
  }, 2000);
}

/**
 * 切换大模型服务商
 * 自动保存前一个服务商的 Key 和 URL，恢复已记忆的新服务商信息或预置默认值
 * @param providerId 目标服务商标识
 */
function selectProvider(providerId: LLMProvider) {
  if (geminiConfig.provider === providerId) return;

  // 1. 记忆当前服务商的 Key 与 Base URL，避免切换后用户重新输入
  if (geminiConfig.provider) {
    if (!geminiConfig.providerKeys) geminiConfig.providerKeys = {};
    if (!geminiConfig.providerUrls) geminiConfig.providerUrls = {};
    geminiConfig.providerKeys[geminiConfig.provider] = geminiConfig.apiKey;
    geminiConfig.providerUrls[geminiConfig.provider] = geminiConfig.baseUrl;
  }

  // 2. 切换当前服务商
  geminiConfig.provider = providerId;
  const targetPreset = LLM_PROVIDERS[providerId] || LLM_PROVIDERS.gemini;

  // 3. 恢复新服务商对应的记忆 Key 或清空
  geminiConfig.apiKey = geminiConfig.providerKeys?.[providerId] || "";

  // 4. 恢复新服务商对应的记忆 URL 或赋予默认 Base URL
  const rememberedUrl = geminiConfig.providerUrls?.[providerId];
  geminiConfig.baseUrl = rememberedUrl !== undefined ? rememberedUrl : targetPreset.defaultBaseUrl;
  geminiConfig.proxyUrl = geminiConfig.baseUrl;

  // 5. 设置新服务商推荐的模型
  if (targetPreset.defaultModel) {
    geminiConfig.model = targetPreset.defaultModel;
    modelSelect.value = targetPreset.defaultModel;
  } else {
    modelSelect.value = "custom";
  }

  // 6. 重置状态提示
  loadModelsError.value = null;
  loadModelsSuccess.value = false;

  handleSave();
}

/**
 * 重置当前服务商接口基础地址至官方默认预置
 */
function resetBaseUrl() {
  if (currentPreset.value.defaultBaseUrl) {
    geminiConfig.baseUrl = currentPreset.value.defaultBaseUrl;
    geminiConfig.proxyUrl = geminiConfig.baseUrl;
    handleSave();
  }
}

/**
 * 模型选择变更回调
 */
function handleModelSelectChange() {
  if (modelSelect.value !== "custom") {
    geminiConfig.model = modelSelect.value;
  }
  handleSave();
}

/**
 * 依据当前服务商和 API Key / Base URL 动态拉取可用模型列表
 */
async function fetchModels() {
  if (!geminiConfig.apiKey && geminiConfig.provider !== "custom") {
    loadModelsError.value = "请先填写对应的 API Key 再拉取模型列表。";
    return;
  }

  isLoadingModels.value = true;
  loadModelsError.value = null;
  loadModelsSuccess.value = false;

  try {
    const models = await fetchAvailableModels(geminiConfig);
    availableModels.value = models;
    loadModelsSuccess.value = true;

    // 若拉取到的模型列表包含当前模型，则保持；否则若列表非空且当前未指定，则选中第一个
    if (models.includes(geminiConfig.model)) {
      modelSelect.value = geminiConfig.model;
    } else if (models.length > 0 && !currentPreset.value.presetModels.includes(geminiConfig.model)) {
      geminiConfig.model = models[0];
      modelSelect.value = models[0];
    }

    setTimeout(() => {
      loadModelsSuccess.value = false;
    }, 3500);
  } catch (err: any) {
    console.error("动态拉取模型列表失败:", err);
    loadModelsError.value = err.message || "网络连接失败或跨域被拦截";
  } finally {
    isLoadingModels.value = false;
  }
}
</script>

<template>
  <div class="glass-card settings-panel">
    <!-- 面板标题栏（支持折叠/展开） -->
    <div class="card-header" @click="isCollapsed = !isCollapsed" style="cursor: pointer; user-select: none">
      <Settings class="header-icon" />
      <h2>大模型与 API 配置设定</h2>
      <div class="collapse-trigger">
        <span class="status-summary" v-if="isCollapsed">
          服务商: {{ currentPreset.name }} | 模型: {{ geminiConfig.model || "未设定" }} |
          {{ isCustomUrl ? "自定义端点" : currentPreset.isDirectGemini ? "官方直连" : "默认端点" }}
        </span>
        <ChevronDown v-if="isCollapsed" :size="16" class="arrow-icon" />
        <ChevronUp v-else :size="16" class="arrow-icon" />
      </div>
    </div>

    <transition name="collapse">
      <div class="card-body" v-show="!isCollapsed">
        <!-- 隐私安全提示条 -->
        <div class="alert-box warning">
          <ShieldAlert class="alert-icon" />
          <div class="alert-content">
            <strong>隐私安全保障：</strong>
            本系统为纯前端实现，您的 API Key 和绩效数据仅保存在浏览器本地
            <code>localStorage</code>
            ，且只直接发送至您所选定的 AI 服务商或指定的反代网关，绝不经由任何第三方应用后端，请放心使用。
          </div>
        </div>

        <!-- 1. 服务商选择器（网格卡片式） -->
        <div class="form-group">
          <label class="section-label">
            <Cpu :size="16" class="label-icon" />
            选择大模型服务商
          </label>
          <div class="providers-grid">
            <div v-for="provider in providerList" :key="provider.id" class="provider-card" :class="{ active: geminiConfig.provider === provider.id }" @click="selectProvider(provider.id)">
              <div class="provider-card-header">
                <span class="provider-name">{{ provider.name }}</span>
                <span class="provider-badge">{{ provider.badge }}</span>
              </div>
              <div class="provider-desc">
                {{ provider.id === "gemini" ? "Google AI Studio 直连" : provider.id === "deepseek" ? "DeepSeek-V3 / R1 推理" : provider.id === "qwen" ? "阿里百炼千问大模型" : provider.id === "openai" ? "官方 GPT-4o / o3 系列" : provider.id === "zhipu" ? "智谱 GLM-4 旗舰大模型" : "兼容 OpenAI 协议任意端点" }}
              </div>
              <div class="active-indicator" v-if="geminiConfig.provider === provider.id">
                <Check :size="14" />
              </div>
            </div>
          </div>
        </div>

        <!-- 2. API Key 与 基础接口地址配置行 -->
        <div class="form-row">
          <!-- API Key 输入框 -->
          <div class="form-group flex-1">
            <div class="label-with-action">
              <label for="apiKey">
                {{ currentPreset.name }} API Key
                <span class="required" v-if="geminiConfig.provider !== 'custom'">*</span>
              </label>
              <a v-if="currentPreset.keyUrl" :href="currentPreset.keyUrl" target="_blank" class="link-btn external-link" title="前往服务商官方控制台获取 API Key">
                <span>获取 Key</span>
                <ExternalLink :size="12" />
              </a>
            </div>
            <div class="input-with-icon">
              <input id="apiKey" :type="showKey ? 'text' : 'password'" v-model="geminiConfig.apiKey" :placeholder="currentPreset.keyPlaceholder" class="form-control" @change="handleSave" />
              <button type="button" class="icon-btn" @click="showKey = !showKey" :title="showKey ? '隐藏密钥' : '显示明文密钥'">
                <Eye v-if="!showKey" :size="18" />
                <EyeOff v-else :size="18" />
              </button>
            </div>
            <span class="help-text">
              {{ currentPreset.keyHelpText }}
              <a v-if="currentPreset.keyUrl" :href="currentPreset.keyUrl" target="_blank" class="link-btn">官方控制台 ↗</a>
            </span>
          </div>

          <!-- Base URL 输入框 -->
          <div class="form-group flex-1">
            <div class="label-with-action">
              <label for="baseUrl">接口基础地址 (Base URL)</label>
              <button type="button" class="text-action-btn" v-if="isCustomUrl && currentPreset.defaultBaseUrl" @click="resetBaseUrl" title="恢复到该服务商官方默认接口地址">
                <RotateCcw :size="12" />
                <span>恢复默认</span>
              </button>
            </div>
            <div class="input-with-icon">
              <input id="baseUrl" type="url" v-model="geminiConfig.baseUrl" :placeholder="currentPreset.defaultBaseUrl || 'https://api.example.com/v1'" class="form-control" @change="handleSave" />
              <Globe class="input-inner-icon" :size="16" />
            </div>
            <span class="help-text">
              <template v-if="geminiConfig.provider === 'gemini'">支持官方直连；若国内网络受限可填写反代地址。</template>
              <template v-else-if="geminiConfig.provider === 'custom'">支持 OneAPI、NewAPI、本地 Ollama（如 http://localhost:11434/v1）等兼容端点。</template>
              <template v-else>默认直连服务商官方 API。支持填写您专属的反向代理或中转网关。</template>
            </span>
            <span class="help-text warning-text" v-if="isCustomUrl">⚠️ 当前正在使用自定义接口地址，请确保目标代理服务安全可靠。</span>
          </div>
        </div>

        <!-- 3. 模型版本选择与加载行 -->
        <div class="form-row">
          <div class="form-group flex-1">
            <label for="modelSelect">AI 模型版本</label>
            <div class="model-select-wrapper">
              <select id="modelSelect" v-model="modelSelect" class="form-control select-control" @change="handleModelSelectChange">
                <!-- 预置推荐模型 -->
                <optgroup label="服务商推荐模型" v-if="currentPreset.presetModels.length > 0">
                  <option v-for="m in currentPreset.presetModels" :key="m" :value="m">
                    {{ m }}
                  </option>
                </optgroup>

                <!-- 动态拉取到的模型列表 -->
                <optgroup label="动态拉取的可用模型" v-if="availableModels.length > 0">
                  <option v-for="m in availableModels" :key="m" :value="m">
                    {{ m }}
                  </option>
                </optgroup>

                <!-- 当前已配且不在上述列表中的模型 -->
                <optgroup label="当前配置模型" v-if="geminiConfig.model && geminiConfig.model !== 'custom' && !currentPreset.presetModels.includes(geminiConfig.model) && !availableModels.includes(geminiConfig.model)">
                  <option :value="geminiConfig.model">{{ geminiConfig.model }}</option>
                </optgroup>

                <!-- 自定义模型选项 -->
                <option value="custom">✍️ 手动输入自定义模型名称...</option>
              </select>

              <button type="button" class="btn btn-outline" :disabled="isLoadingModels" @click="fetchModels" title="向当前 API 接口发送查询，获取所有可调用的模型列表">
                <RefreshCw :size="16" :class="{ 'animate-spin': isLoadingModels }" />
                <span>{{ isLoadingModels ? "拉取中..." : "拉取模型" }}</span>
              </button>
            </div>

            <!-- 拉取结果状态提示 -->
            <span class="help-text text-success-desc animate-fade-in" v-if="loadModelsSuccess" style="color: #10b981; display: flex; align-items: center; gap: 4px; margin-top: 4px">
              <Check :size="14" />
              <span>成功拉取并更新了 {{ availableModels.length }} 个可用模型！</span>
            </span>
            <span class="help-text text-danger-desc animate-fade-in" v-if="loadModelsError" style="color: #f87171; display: flex; align-items: center; gap: 4px; margin-top: 4px">
              <XCircle :size="14" />
              <span>加载失败: {{ loadModelsError }}</span>
            </span>
          </div>

          <!-- 自定义模型标识符输入（当选中自定义时展示） -->
          <div class="form-group flex-1 animate-fade-in" v-if="modelSelect === 'custom'">
            <label for="customModel">自定义模型标识符 (Model ID)</label>
            <input id="customModel" type="text" v-model="geminiConfig.model" placeholder="例如 deepseek-chat、qwen-max 或 gpt-4o" class="form-control" @change="handleSave" />
            <span class="help-text">请填写模型在服务商接口中的精确模型名称代码。</span>
          </div>
        </div>

        <!-- 4. 考核周期月份配置 -->
        <div class="form-row">
          <div class="form-group flex-1">
            <label for="performanceMonth">绩效考核月份</label>
            <input id="performanceMonth" type="month" v-model="performanceMonth" class="form-control" @change="handleSave" />
            <span class="help-text">选择要评定的绩效年月。AI 导出的工作目标建议时间、Excel 里的标题/考核周期，以及生成的文件名均将自动适配该月份。</span>
          </div>
        </div>

        <!-- 5. 全局系统提示词 -->
        <div class="form-group">
          <label for="systemInstruction">全局考核基准 (System Prompt)</label>
          <textarea id="systemInstruction" v-model="geminiConfig.systemInstruction" rows="3" placeholder="例如：你是一个专业、严谨的企业HR绩效评定官..." class="form-control textarea-control" @change="handleSave"></textarea>
          <span class="help-text">用于定义 AI 在评估过程中的角色、语气以及遵循的通用企业评级规则。</span>
        </div>

        <!-- 保存按钮 -->
        <div class="form-actions">
          <button type="button" class="btn btn-primary btn-save" @click="handleSave">
            <Check v-if="savedFeedback" :size="18" />
            <Save v-else :size="18" />
            <span>{{ savedFeedback ? "已自动保存" : "保存配置" }}</span>
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

/* 服务商选择网格 */
.section-label {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 0.95rem;
  font-weight: 600;
  color: var(--text-h);
  margin-bottom: 4px;
}

.label-icon {
  color: var(--accent);
}

.providers-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(210px, 1fr));
  gap: 12px;
}

.provider-card {
  position: relative;
  display: flex;
  flex-direction: column;
  padding: 14px 16px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid var(--border);
  border-radius: 10px;
  cursor: pointer;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  overflow: hidden;
  user-select: none;
}

.provider-card:hover {
  background: rgba(255, 255, 255, 0.06);
  border-color: rgba(168, 85, 247, 0.4);
  transform: translateY(-2px);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
}

.provider-card.active {
  background: linear-gradient(135deg, rgba(168, 85, 247, 0.18) 0%, rgba(99, 102, 241, 0.18) 100%);
  border-color: var(--accent);
  box-shadow: 0 4px 20px rgba(168, 85, 247, 0.25);
}

.provider-card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 6px;
}

.provider-name {
  font-size: 0.95rem;
  font-weight: 600;
  color: var(--text-h);
}

.provider-badge {
  font-size: 0.7rem;
  font-weight: 500;
  padding: 2px 7px;
  border-radius: 12px;
  background: rgba(255, 255, 255, 0.08);
  color: var(--text);
  white-space: nowrap;
}

.provider-card.active .provider-badge {
  background: rgba(168, 85, 247, 0.3);
  color: #f3e8ff;
  border: 1px solid rgba(168, 85, 247, 0.4);
}

.provider-desc {
  font-size: 0.775rem;
  color: var(--text);
  line-height: 1.4;
}

.active-indicator {
  position: absolute;
  top: 6px;
  right: 6px;
  color: var(--accent);
  display: flex;
  align-items: center;
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

.label-with-action {
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
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
  width: 100%;
}

.form-control:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px var(--accent-bg);
  background: rgba(255, 255, 255, 0.08);
}

.model-select-wrapper {
  display: flex;
  gap: 8px;
  width: 100%;
}

.select-control {
  appearance: none;
  background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%239ca3af' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: right 14px center;
  background-size: 16px;
  padding-right: 40px;
  flex: 1;
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
  padding-right: 44px;
}

.input-inner-icon {
  position: absolute;
  right: 14px;
  color: var(--text);
  pointer-events: none;
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
  transition:
    color 0.2s,
    background-color 0.2s;
}

.icon-btn:hover {
  color: var(--text-h);
  background: rgba(255, 255, 255, 0.1);
}

.text-action-btn {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  background: transparent;
  border: none;
  color: var(--accent);
  font-size: 0.775rem;
  cursor: pointer;
  padding: 2px 6px;
  border-radius: 4px;
  transition: all 0.2s;
}

.text-action-btn:hover {
  background: var(--accent-bg);
  color: #f3e8ff;
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
  display: inline-flex;
  align-items: center;
  gap: 3px;
}

.link-btn:hover {
  color: var(--text-h);
  border-bottom-color: var(--text-h);
}

.external-link {
  font-size: 0.775rem;
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
  padding: 10px 18px;
  border-radius: 8px;
  cursor: pointer;
  border: 1px solid rgba(255, 255, 255, 0.08);
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  background: rgba(255, 255, 255, 0.03);
  color: var(--text-h);
  backdrop-filter: blur(4px);
  white-space: nowrap;
}

.btn-outline:hover:not(:disabled) {
  border-color: var(--accent);
  background: rgba(168, 85, 247, 0.1);
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
  transition:
    max-height 0.35s cubic-bezier(0.4, 0, 0.2, 1),
    opacity 0.3s ease-out;
  overflow: hidden;
  max-height: 800px;
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
.select-control option,
.select-control optgroup {
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

.animate-spin {
  animation: spin 1.5s linear infinite;
}

@keyframes spin {
  from {
    transform: rotate(0deg);
  }
  to {
    transform: rotate(360deg);
  }
}
</style>
