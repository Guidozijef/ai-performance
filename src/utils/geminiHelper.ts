/**
 * @fileoverview Gemini API 客户端调用工具。
 * 封装了对 Google Gemini API 的直接 HTTPS 请求，支持结构化 JSON 输出、自定义代理 API 网关、自定义模型和系统提示词。
 * 遵循 Google TypeScript 编码标准。
 */

import type { CellMapping, PerformanceTask } from "./excelHelper";

/**
 * Gemini 客户端请求配置
 */
export interface GeminiConfig {
  /** API 密钥 */
  apiKey: string;
  /** 备用 API 代理网关地址（例如 https://api.example.com ），若为空则使用官方域名 */
  proxyUrl?: string;
  /** 调用的模型名称，默认使用 gemini-1.5-flash */
  model?: string;
  /** 全局系统提示词，用于设定 AI 的角色和整体考核基准 */
  systemInstruction?: string;
}

/**
 * AI 生成结果单条数据定义
 */
export interface AIGeneratedResult {
  /** 单元格坐标与生成内容的映射，例如 { "C8": 92, "A10": "表现优秀" } */
  cellData: Record<string, string | number>;
  /** 生成过程中的提示信息或错误信息 */
  message?: string;
  /** 是否成功 */
  success: boolean;
}

/**
 * 调用 Gemini API 生成单个员工的绩效字段内容
 *
 * @param config Gemini API 客户端配置
 * @param employeeInput 员工基本数据及工作描述（键值对）
 * @param aiMappings 需要 AI 生成的单元格映射列表
 * @returns Promise<AIGeneratedResult> 返回生成的单元格数据对象
 */
export async function generateEmployeePerformance(config: GeminiConfig, employeeInput: Record<string, string | number>, aiMappings: CellMapping[]): Promise<AIGeneratedResult> {
  const { apiKey, proxyUrl, model = "gemini-1.5-flash", systemInstruction } = config;

  if (!apiKey) {
    return {
      cellData: {},
      message: "未配置 API Key，请先在设置中填写 Gemini API Key。",
      success: false,
    };
  }

  if (aiMappings.length === 0) {
    return {
      cellData: {},
      message: "未配置任何需要 AI 生成的单元格映射。",
      success: false,
    };
  }

  // 1. 构建 Prompt，明确告知模型输出格式及字段要求
  const inputDataStr = Object.entries(employeeInput)
    .map(([key, val]) => `- ${key}: ${val}`)
    .join("\n");

  const fieldsDemandStr = aiMappings.map((m) => `- 单元格 [${m.cellRef}] (期望生成的内容: ${m.label}): ${m.aiInstruction || "请结合员工表现生成合适的内容"}`).join("\n");

  const prompt = `
你是一个专业的 HR 绩效考评专家。请根据以下员工的基础信息及工作表现，生成对应的绩效考核内容。

【员工基础信息与工作表现】：
${inputDataStr}

【需要你生成的单元格数据及对应要求】：
${fieldsDemandStr}

【生成规则】：
1. 必须完全依照上述需要生成的单元格字段进行生成，不要遗漏任何一个指定的单元格。
2. 评分或数字字段：必须只返回数值（如 90, 85），不要包含任何单位或文字说明。
3. 文本或评语字段：要求语言专业、符合 HR 绩效考核的严谨口吻，切合该员工的具体工作描述，避免过度空泛的夸赞。
4. 你必须返回一个标准的 JSON 对象，其键名是单元格坐标（如 "C8", "A10"），其值是对应的生成内容。
5. 仅返回 JSON 字符串本身，严禁在 JSON 前后包裹任何 Markdown 标记（例如 \`\`\`json ... \`\`\`），以确保代码能够直接进行 JSON.parse 解析。
`;

  // 2. 双轨协议选择：官方直连使用 Gemini 原生 API（避免 CORS 跨域拦截），代理使用 OpenAI 协议
  const isDirectGemini = !proxyUrl || proxyUrl.trim().includes("generativelanguage.googleapis.com");

  let requestUrl = "";
  let payload: any = {};
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };

  if (isDirectGemini) {
    // 直连官方：使用原生 REST 接口（已由谷歌开启 CORS 许可）
    requestUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    payload = {
      contents: [
        {
          parts: [
            {
              text: prompt,
            },
          ],
        },
      ],
      generationConfig: {
        responseMimeType: "application/json",
        temperature: 0.7,
      },
      ...(systemInstruction?.trim()
        ? {
            systemInstruction: {
              parts: [{ text: systemInstruction }],
            },
          }
        : {}),
    };
  } else {
    // 代理网关：采用标准 OpenAI chat/completions 与 Bearer Auth
    let baseUrl = proxyUrl.trim().replace(/\/$/, "");
    if (baseUrl.endsWith("/v1") || baseUrl.includes("/v1/") || baseUrl.includes("/v1beta")) {
      requestUrl = `${baseUrl}/chat/completions`;
    } else {
      requestUrl = `${baseUrl}/v1/chat/completions`;
    }
    headers["Authorization"] = `Bearer ${apiKey}`;

    // 在本地开发环境下，通过 Vite 动态代理规避代理服务可能包含的 CORS 限制
    if (import.meta.env.DEV) {
      try {
        const urlObj = new URL(requestUrl);
        const origin = urlObj.origin;
        const pathname = urlObj.pathname.replace(/\/$/, "");
        requestUrl = `/cors-proxy${pathname}`.replace(/\/+/g, "/");
        headers["X-Target-Url"] = origin;
      } catch (e) {
        // 降级处理：若 URL 解析失败，保留原路径直接发起请求
        console.warn("CORS Proxy URL 解析失败，降级直连:", e);
        const originalUrl = requestUrl;
        requestUrl = "/cors-proxy";
        headers["X-Target-Url"] = originalUrl;
      }
    }

    payload = {
      model,
      messages: [...(systemInstruction?.trim() ? [{ role: "system", content: systemInstruction }] : []), { role: "user", content: prompt }],
      temperature: 0.7,
    };
  }

  try {
    const response = await fetch(requestUrl, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP 错误！状态码: ${response.status}, 详情: ${errorText}`);
    }

    const resData = await response.json();
    let generatedText = "";

    if (isDirectGemini) {
      generatedText = resData.candidates?.[0]?.content?.parts?.[0]?.text;
    } else {
      generatedText = resData.choices?.[0]?.message?.content;
    }

    if (!generatedText) {
      throw new Error("模型未返回有效的内容。");
    }

    // 4. 解析返回的 JSON 数据
    const cleanText = generatedText
      .trim()
      .replace(/^```json/, "")
      .replace(/```$/, "")
      .trim();
    const cellData = JSON.parse(cleanText) as Record<string, string | number>;

    // 校验返回的数据是否包含我们需要的单元格
    const validatedData: Record<string, string | number> = {};
    for (const mapping of aiMappings) {
      const ref = mapping.cellRef;
      if (ref in cellData) {
        validatedData[ref] = cellData[ref];
      } else {
        // 若 AI 遗漏了某个字段，给予默认提示或空白
        validatedData[ref] = "";
      }
    }

    return {
      cellData: validatedData,
      success: true,
    };
  } catch (error: any) {
    console.error("API 调用异常:", error);
    return {
      cellData: {},
      message: error.message || "调用 API 时发生未知错误。",
      success: false,
    };
  }
}

/**
 * 正式员工绩效智能生成响应定义
 */
export interface FormalPerformanceResult {
  name: string;
  position: string;
  tasks: PerformanceTask[];
  success: boolean;
  message?: string;
}

/**
 * 调用 Gemini API 为正式员工生成月度绩效任务列表
 *
 * @param config Gemini API 客户端配置
 * @param name 员工姓名
 * @param position 员工岗位
 * @param lastMonthPerformance 上个月绩效内容 (文本)
 * @param thisMonthWorkContent 本月工作安排及内容 (文本)
 * @param targetMonth 考核月份，格式为 YYYY-MM
 * @returns Promise<FormalPerformanceResult> 返回生成的员工姓名、岗位及绩效任务数组
 */
export async function generateFormalPerformance(
  config: GeminiConfig,
  name: string,
  position: string,
  lastMonthPerformance: string,
  thisMonthWorkContent: string,
  targetMonth: string, // 考核月份，格式为 YYYY-MM
): Promise<FormalPerformanceResult> {
  const { apiKey, proxyUrl, model = "gemini-1.5-flash", systemInstruction } = config;

  if (!apiKey) {
    return {
      name,
      position,
      tasks: [],
      success: false,
      message: "未配置 API Key，请先在设置中填写 Gemini API Key。",
    };
  }

  // 1. 安全解析传入的考核月份，防范非法输入的注入攻击，并动态推导当月最后一天
  let year = 2026;
  let month = 6;
  let lastDay = 30;
  if (targetMonth && typeof targetMonth === "string") {
    const parts = targetMonth.split("-");
    if (parts.length === 2) {
      const parsedYear = parseInt(parts[0], 10);
      const parsedMonth = parseInt(parts[1], 10);
      if (!isNaN(parsedYear) && !isNaN(parsedMonth)) {
        year = parsedYear;
        month = parsedMonth;
        lastDay = new Date(year, month, 0).getDate();
      }
    }
  }

  // 2. 构建大模型提示词，要求严密、科学、可量化，坚决杜绝任何主观或模糊的文字描述
  const prompt = `
你是一个资深的企业人力资源（HR）绩效考评专家。
请根据员工的基本信息、上个月填写的绩效内容、以及这个月的一些工作内容/工作安排，为该员工自动生成“${year}年${month}月的工作绩效计划表（即 Excel 模板中红色的任务部分）”。

【员工基本信息】：
- 被考核人姓名: ${name}
- 岗位名称: ${position}

【员工上个月绩效内容回顾】：
${lastMonthPerformance || "无上月绩效数据"}

【员工本月工作内容与计划安排】：
${thisMonthWorkContent || "无本月工作安排"}

【绩效计划制定规则】：
1. 必须生成【至少 4 个】工作任务项目（通常为 4-5 个），以完整科学地考核该员工的工作。即使本月输入的工作安排内容较少，你也要根据其岗位和项目对其进行合理拆解、细化出至少 4 项，绝对不能少于 4 项！
2. 任务行的指标分配与权重分配约束（总权重必须刚好等于 100%，即小数 1.0）：
   - 【等级与类型严格绑定】：指标等级包含“核心战略任务”、“重要关键任务”、“常规执行任务”或“辅助零散任务”。
     * 如果指标等级为常规任务（即等级为“常规执行任务”或“辅助零散任务”），那么其对应的指标类型（type）必须是“CPI”。
     * 如果指标等级为核心或重要任务（即等级为“核心战略任务”或“重要关键任务”），其对应的指标类型（type）必须是“KPI”。
   - 【指标权重约束】：每个任务分配的权重必须是 5% 的倍数（如以 0.3, 0.25, 0.2, 0.15, 0.1, 0.05 等小数形式表示）。
   - 【重要限制：每一项指标权重最大绝对不能超过 30%（即 weight <= 0.30）】。
   - 【重要限制：“重要关键任务”的分配比例（权重）最低不能低于 20%（即 weight >= 0.20，且最大不超过 0.30）】。
   - 所有任务的权重之和必须由你平摊并计算，确保生成的 JSON 中 tasks 列表中所有 weight 字段相加之和【精准等于 1.0】。如果不够 100%，你必须增加更多符合限制的项目以达到 100%。
3. 【核心极重要规则：质量目标与质量标准必须绝对客观可量化，严禁任何主观描述，且必须分条列出序号】：
   - 质量目标（quality_target）和质量标准（quality_standard）的内容必须采用分条列出序号格式（如：'1. [具体项1]\n2. [具体项2]\n3. [具体项3]'），不得写成一整段话，每条都必须有精确的量化边界或对应的违规扣分触发条件。
   - 【严禁使用主观描述及模糊形容词】。禁止出现任何诸如：“积极主动”、“认真细致”、“配合度高”、“高效率”、“及时完成”、“保证系统稳定”、“正常运行”、“无重大故障”、“工作量饱满”、“沟通良好”等字眼。
   - 【量化替换方案参考】：
     * “保证系统稳定/正常运行” ➡️ 替换为：“1.系统月度可用性（SLA）达到99.9%\n2.发生严重级别（一级/二级）线上事故次数为0次\n3.监控系统报警响应延迟在5分钟内”
     * “配合好、沟通积极” ➡️ 替换为：“1.跨部门接口及业务联调进度偏差在1天以内\n2.联调阶段未因单方开发质量原因导致项目整体阻塞达2小时以上”
     * “及时完成、无延期” ➡️ 替换为：“1.在${year}年${month}月${lastDay}日前按原型要求完成代码交付并通过测试验收\n2.交付延期偏差为0天”
   - 质量目标（quality_target）必须是针对指标名称-解释说明的约束和限定，通常为3-4条说明，例如：
     * 按照产品原型要求完成功能开发
     * 提测用例通过率达100%
     * 功能开发无遗漏点
     * 满意度100%
   - 质量标准（quality_standard）必须是针对质量目标所做出的【扣分细则/不得分细则】：
     * 每一个质量目标序号，都必须能在质量标准中找到对应的扣分规则。
     * 比如目标是：“1.提测用例通过率达100%”，对应的标准应该是：“1.提测用例通过率每低于100%一个百分点扣2分”。
     * 比如目标是：“2.满意度100%”，对应的标准应该是：“2.收到测试/产品/市场侧投诉扣3分/次”。
4. 【精简聚焦，拒绝冗余】：
   - 生成的每一项质量目标（quality_target）和质量标准（quality_standard）应当简明扼要，直击工作核心，避免字数堆砌或冗余描述。
   - 每个任务的质量目标和标准数量严格控制在 2-3 条以内，确保清晰直观。不要生成长篇大论。
5. 【严格继承与参考】：
   - 必须优先参考【员工上个月绩效内容回顾】中各项工作的专业技术表达（例如：提测用例通过率、严重级别BUG、代码审计通过率等）和评分的扣分尺度。
   - 如果上个月的绩效内容存在，必须严格继承并保留其表达习惯、指标风格和逻辑链条，确保前后两个月考核指标的连续性。
   - 若无上月数据，请按照上述第3条量化方案示例生成。

6. 任务字段对应定义：
   - type: 指标类型，通常为 "KPI" 或 "CPI"。
   - level: 指标等级，可从 "核心战略任务"、"重要关键任务"、"常规执行任务"、"辅助零散任务" 中选择。
   - weight: 权重数值（小数形式，如 0.3 代表 30%）。
   - category: 所属板块，例如 "系统开发"、"系统维护"、"系统优化" 等。
   - description: 解释说明，即任务的具体背景或重要工作内容（如：无人机V2.1.8版本迭代开发）。
   - time_target: 时间目标，如 "${year}年${month}月${lastDay - 1}日"。
   - count_target: 数量目标，如果没有则填写 "/"。
   - quality_target: 质量目标（务必使用分条序号格式，客观可量化，禁止主观判断）。
   - time_standard: 时间标准，如果没有则填写 "/"，或者写如 "每延迟一天扣 3 分，逾期 3 天及以上该项不得分"。
   - count_standard: 数量标准，如果没有则填写 "/"。
   - quality_standard: 质量标准（务必使用分条序号格式，客观可量化，针对质量目标给出具体扣分细则，禁止主观判断）。

【输出格式要求】：
你必须返回一个合法的 JSON 对象，其属性包括 name、position 以及 tasks 数组。
严禁在 JSON 前后包裹任何 Markdown 标记（例如 \`\`\`json ... \`\`\`），以确保代码能够直接进行 JSON.parse 解析。

JSON 结构样例：
{
  "name": "${name}",
  "position": "${position}",
  "tasks": [
    {
      "type": "KPI",
      "level": "重要关键任务",
      "weight": 0.3,
      "category": "系统开发",
      "description": "手持甲烷探测器功能开发...",
      "time_target": "${year}年${month}月${lastDay - 1}日",
      "count_target": "/",
      "quality_target": "1.按产品原型要求完成APP侧XX功能开发与接口对接，提测用例通过率达100%\\n2.配合后端进行功能联调且无堵塞\\n3.内外部客户满意度达100%",
      "time_standard": "每延迟一天扣3分，累计3天及以上不得分",
      "count_standard": "/",
      "quality_standard": "1.未按产品原型要求完成该项任务不得分\\n2.提测后发生严重级别BUG（一级、二级）未在24小时内解决扣5分/次\\n3.联调阶段因开发质量问题导致联调延期扣2分/天\\n4.已完成功能收到内外部有效投诉扣2分/次"
    }
  ]
}
`;

  // 3. 双轨协议选择：官方直连使用 Gemini 原生 API（避免 CORS 跨域拦截），代理使用 OpenAI 协议
  const isDirectGemini = !proxyUrl || proxyUrl.trim().includes("generativelanguage.googleapis.com");

  let requestUrl = "";
  let payload: any = {};
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };

  if (isDirectGemini) {
    // 直连官方：使用原生 REST 接口（已由谷歌开启 CORS 许可）
    requestUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    payload = {
      contents: [
        {
          parts: [
            {
              text: prompt,
            },
          ],
        },
      ],
      generationConfig: {
        responseMimeType: "application/json",
        temperature: 0.4,
      },
      ...(systemInstruction?.trim()
        ? {
            systemInstruction: {
              parts: [{ text: systemInstruction }],
            },
          }
        : {}),
    };
  } else {
    // 代理网关：采用标准 OpenAI chat/completions 与 Bearer Auth
    let baseUrl = proxyUrl.trim().replace(/\/$/, "");
    if (baseUrl.endsWith("/v1") || baseUrl.includes("/v1/") || baseUrl.includes("/v1beta")) {
      requestUrl = `${baseUrl}/chat/completions`;
    } else {
      requestUrl = `${baseUrl}/v1/chat/completions`;
    }
    headers["Authorization"] = `Bearer ${apiKey}`;

    // 在本地开发环境下，通过 Vite 动态代理规避代理服务可能包含的 CORS 限制
    if (import.meta.env.DEV) {
      try {
        const urlObj = new URL(requestUrl);
        const origin = urlObj.origin;
        const pathname = urlObj.pathname.replace(/\/$/, "");
        requestUrl = `/cors-proxy${pathname}`.replace(/\/+/g, "/");
        headers["X-Target-Url"] = origin;
      } catch (e) {
        // 降级处理：若 URL 解析失败，保留原路径直接发起请求
        console.warn("CORS Proxy URL 解析失败，降级直连:", e);
        const originalUrl = requestUrl;
        requestUrl = "/cors-proxy";
        headers["X-Target-Url"] = originalUrl;
      }
    }

    payload = {
      model,
      messages: [...(systemInstruction?.trim() ? [{ role: "system", content: systemInstruction }] : []), { role: "user", content: prompt }],
      temperature: 0.4,
    };
  }

  try {
    const response = await fetch(requestUrl, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP 错误！状态码: ${response.status}, 详情: ${errorText}`);
    }

    const resData = await response.json();
    let generatedText = "";

    if (isDirectGemini) {
      generatedText = resData.candidates?.[0]?.content?.parts?.[0]?.text;
    } else {
      generatedText = resData.choices?.[0]?.message?.content;
    }

    if (!generatedText) {
      throw new Error("模型未返回有效的内容。");
    }

    const cleanText = generatedText
      .trim()
      .replace(/^```json/, "")
      .replace(/```$/, "")
      .trim();
    const parsedData = JSON.parse(cleanText);

    let tasks: PerformanceTask[] = parsedData.tasks || [];

    // 1. 规范化指标等级与指标类型（常规/辅助任务对应 CPI，重要/核心对应 KPI）
    tasks.forEach((t) => {
      if (!["核心战略任务", "重要关键任务", "常规执行任务", "辅助零散任务"].includes(t.level)) {
        t.level = "常规执行任务";
      }
      if (t.level === "常规执行任务" || t.level === "辅助零散任务") {
        t.type = "CPI";
      } else {
        t.type = "KPI";
      }
    });

    // 2. 格式化权重值，并将其舍入为最近的 5% (0.05) 倍数
    tasks.forEach((t) => {
      let w = parseFloat(t.weight as any);
      if (isNaN(w)) w = 0.2;
      if (w > 1) w = w / 100;
      w = Math.round(w * 20) / 20; // 舍入到 0.05
      t.weight = w;
    });

    // 3. 约束每一项权重的边界：最大不超过 30% (0.30)；重要关键任务最低不能低于 20% (0.20)
    tasks.forEach((t) => {
      if (t.level === "重要关键任务") {
        if (t.weight < 0.2) t.weight = 0.2;
      }
      if (t.weight > 0.3) {
        t.weight = 0.3;
      }
    });

    // 4. 如果不够 4 个指标项目，进行补齐，因为单项最大 30%，要凑够 100% 至少需要 4 项
    while (tasks.length < 4) {
      tasks.push({
        type: "CPI",
        level: "常规执行任务",
        weight: 0.1,
        category: "日常事务",
        description: "常规日常工作支持",
        time_target: "当月月底前",
        count_target: "/",
        quality_target: "1.按时保质完成上级交办的常规事务\\n2.部门内协作响应时间不超过1小时",
        time_standard: "每延迟一天扣2分",
        count_standard: "/",
        quality_standard: "1.未按要求交付且无提前说明扣5分/次\\n2.受到跨部门协作投诉核实扣2分/次",
      });
    }

    // 5. 循环微调以 5% (0.05) 为增量，确保权重之和严格等于 1.0 (100%)，且均满足边界条件
    let iterations = 0;
    while (iterations < 100) {
      let sum = tasks.reduce((acc, t) => acc + t.weight, 0);
      sum = Math.round(sum * 100) / 100; // 避免浮点精度问题

      if (Math.abs(sum - 1.0) < 0.001) {
        break;
      }

      if (sum < 1.0) {
        // 总权重不足，且单项最大 30%：增加能够增加的项
        let increased = false;
        for (let t of tasks) {
          if (t.weight + 0.05 <= 0.3001) {
            t.weight = Math.round((t.weight + 0.05) * 100) / 100;
            increased = true;
            break;
          }
        }
        // 如果所有已有项都达到了 30% 且依旧总权重不够（例如 3 项 30% 共 90%），则必须新增一项
        if (!increased) {
          tasks.push({
            type: "CPI",
            level: "常规执行任务",
            weight: 0.1,
            category: "日常事务",
            description: "追加常规业务支持",
            time_target: "当月月底前",
            count_target: "/",
            quality_target: "1.保质保量完成部门安排的常规任务",
            time_standard: "/",
            count_standard: "/",
            quality_standard: "1.未达标工作受到组内提醒扣1分/次",
          });
        }
      } else {
        // 总权重超出：减少能够减少的项
        let decreased = false;
        for (let t of tasks) {
          const minLimit = t.level === "重要关键任务" ? 0.2 : 0.05;
          if (t.weight - 0.05 >= minLimit - 0.001) {
            t.weight = Math.round((t.weight - 0.05) * 100) / 100;
            decreased = true;
            break;
          }
        }
        // 如果在约束下无法继续减少，则打破约束微调非重要任务
        if (!decreased) {
          for (let t of tasks) {
            if (t.level !== "重要关键任务" && t.weight > 0.05) {
              t.weight = Math.round((t.weight - 0.05) * 100) / 100;
              decreased = true;
              break;
            }
          }
        }
        // 最终兜底
        if (!decreased) {
          tasks[0].weight = Math.round((tasks[0].weight - 0.05) * 100) / 100;
        }
      }
      iterations++;
    }

    return {
      name: parsedData.name || name,
      position: parsedData.position || position,
      tasks,
      success: true,
    };
  } catch (error: any) {
    console.error("正式员工绩效生成异常:", error);
    return {
      name,
      position,
      tasks: [],
      success: false,
      message: error.message || "生成正式员工绩效时发生未知错误。",
    };
  }
}
