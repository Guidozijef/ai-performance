<script setup lang="ts">
/**
 * @fileoverview App.vue 主入口组件。
 * 提供“正式员工 JX1.6 绩效套表智能填报”专区（默认）和“通用自定义模板映射回写”双模式。
 * 包含全局配置、响应式布局、高档玻璃拟态暗黑主题设定。
 * 遵循 Google TS/Vue 编码标准，包含详细中文注释。
 */

import { ref } from 'vue';
import SettingsPanel from './components/SettingsPanel.vue';
import TemplateConfig from './components/TemplateConfig.vue';
import BatchGenerator from './components/BatchGenerator.vue';
import FormalGenerator from './components/FormalGenerator.vue';
import { Sparkles, FileSpreadsheet, Sliders } from 'lucide-vue-next';

// 记录当前活跃的模式页签：formal - 正式员工特定套表智能填报；generic - 通用模板单元格坐标映射
const activeTab = ref<'formal' | 'generic'>('formal');
</script>

<template>
  <header class="app-header">
    <div class="header-container">
      <div class="logo-area">
        <div class="logo-icon-bg">
          <Sparkles class="logo-icon animate-pulse" />
        </div>
        <div class="logo-text">
          <h1>AI 智能员工绩效回写助手</h1>
          <p>基于 Gemini 与 ExcelJS 的纯前端批量绩效生成与格式保留导出工具</p>
        </div>
      </div>
      
      <!-- 双模式切换 Tab 导航 -->
      <nav class="mode-tabs">
        <button 
          type="button" 
          class="tab-btn" 
          :class="{ active: activeTab === 'formal' }"
          @click="activeTab = 'formal'"
        >
          <FileSpreadsheet :size="16" />
          <span>正式员工 JX1.6 专区</span>
        </button>
        <button 
          type="button" 
          class="tab-btn" 
          :class="{ active: activeTab === 'generic' }"
          @click="activeTab = 'generic'"
        >
          <Sliders :size="16" />
          <span>自定义模板映射模式</span>
        </button>
      </nav>
    </div>
  </header>

  <main class="app-main animate-fade-in">
    <!-- 步骤一：API 配置设定（全局通用） -->
    <section class="step-section">
      <div class="step-badge">系统设置</div>
      <SettingsPanel />
    </section>

    <!-- 模式二：正式员工 JX1.6 套表专区 (默认激活) -->
    <template v-if="activeTab === 'formal'">
      <section class="step-section">
        <div class="step-badge">核心工作台</div>
        <FormalGenerator />
      </section>
    </template>

    <!-- 模式一：通用坐标映射回写模式 -->
    <template v-else-if="activeTab === 'generic'">
      <section class="step-section">
        <div class="step-badge">步骤 2</div>
        <TemplateConfig />
      </section>

      <section class="step-section">
        <div class="step-badge">步骤 3</div>
        <BatchGenerator />
      </section>
    </template>
  </main>

  <footer class="app-footer">
    <p>© 2026 AI Performance Helper. 完全基于浏览器本地运行，切实保障企业数据隐私与密钥安全。</p>
  </footer>
</template>

<style>
/* 全局样式规范与布局定义 */
:root {
  --sans: 'Outfit', system-ui, 'Segoe UI', Roboto, sans-serif;
  --bg-gradient: radial-gradient(circle at top, #1e1b4b 0%, #0f172a 100%);
  --border: rgba(255, 255, 255, 0.08);
  --card-bg: rgba(30, 41, 59, 0.4);
  --text-h: #f8fafc;
  --text: #94a3b8;
  --accent: #a855f7;
  --accent-bg: rgba(168, 85, 247, 0.12);
  --accent-border: rgba(168, 85, 247, 0.4);
  --shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
}

body {
  margin: 0;
  padding: 0;
  font-family: var(--sans);
  background: #0f172a;
  background-image: var(--bg-gradient);
  color: var(--text);
  min-height: 100vh;
  -webkit-font-smoothing: antialiased;
}

#app {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
}

/* 顶部 Header */
.app-header {
  border-bottom: 1px solid var(--border);
  background: rgba(15, 23, 42, 0.6);
  backdrop-filter: blur(12px);
  position: sticky;
  top: 0;
  z-index: 50;
}

/* 顶部导航条容器，设为最大宽度 1780px 以匹配宽屏自适应大面板 */
.header-container {
  max-width: 1780px;
  margin: 0 auto;
  padding: 14px 24px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 16px;
}

.logo-area {
  display: flex;
  align-items: center;
  gap: 16px;
}

.logo-icon-bg {
  width: 42px;
  height: 42px;
  background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%);
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 4px 14px rgba(168, 85, 247, 0.4);
}

.logo-icon {
  color: white;
  width: 22px;
  height: 22px;
}

.logo-text {
  text-align: left;
}

.logo-text h1 {
  margin: 0;
  font-size: 1.35rem;
  font-weight: 700;
  color: var(--text-h);
  letter-spacing: -0.5px;
}

.logo-text p {
  margin: 2px 0 0;
  font-size: 0.8rem;
  color: var(--text);
}

/* 页签切换导航 */
.mode-tabs {
  display: flex;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 4px;
  gap: 4px;
}

.tab-btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  border: none;
  background: transparent;
  color: var(--text);
  font-size: 0.85rem;
  font-weight: 600;
  padding: 8px 16px;
  cursor: pointer;
  border-radius: 6px;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
}

.tab-btn:hover:not(.active) {
  color: var(--text-h);
  background: rgba(255, 255, 255, 0.03);
}

.tab-btn.active {
  background: linear-gradient(135deg, #a855f7 0%, #6366f1 100%);
  color: white;
  box-shadow: 0 4px 12px rgba(168, 85, 247, 0.25);
}

/* 主体容器，调整至 1780px 宽屏以适配 100% 满宽表格和更宽的输入交互区域 */
.app-main {
  max-width: 1780px;
  width: 100%;
  margin: 0 auto;
  padding: 32px 24px;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  gap: 36px;
  flex-grow: 1;
}

/* 步骤分段样式 */
.step-section {
  position: relative;
}

.step-badge {
  position: absolute;
  top: -12px;
  left: 24px;
  z-index: 10;
  background: linear-gradient(135deg, #a855f7 0%, #6366f1 100%);
  color: white;
  font-size: 0.725rem;
  font-weight: 700;
  padding: 4px 10px;
  border-radius: 20px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  box-shadow: 0 2px 8px rgba(168, 85, 247, 0.3);
}

/* 玻璃卡片通用样式 */
.glass-card {
  background: var(--card-bg);
  border: 1px solid var(--border);
  backdrop-filter: blur(16px);
  border-radius: 12px;
  box-shadow: var(--shadow);
  overflow: hidden;
  transition: border-color 0.3s, box-shadow 0.3s;
}

.glass-card:hover {
  border-color: rgba(168, 85, 247, 0.2);
  box-shadow: 0 12px 40px -10px rgba(0, 0, 0, 0.6), 0 0 20px rgba(168, 85, 247, 0.03);
}

/* 页脚 */
.app-footer {
  text-align: center;
  padding: 24px;
  font-size: 0.8rem;
  color: var(--text);
  border-top: 1px solid var(--border);
  background: rgba(15, 23, 42, 0.4);
}

/* 动画 */
.animate-pulse {
  animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
}

.animate-fade-in {
  animation: fadeIn 0.4s ease-out;
}

@keyframes pulse {
  0%, 100% {
    opacity: 1;
  }
  50% {
    opacity: .5;
  }
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}

@media (max-width: 768px) {
  .header-container {
    flex-direction: column;
    align-items: flex-start;
  }
  .mode-tabs {
    width: 100%;
  }
  .tab-btn {
    flex: 1;
    justify-content: center;
  }
}
</style>
