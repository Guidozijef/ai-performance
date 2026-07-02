import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";
import { IncomingMessage, ServerResponse, request as httpRequest } from "http";
import { request as httpsRequest } from "https";

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    {
      name: "dynamic-cors-proxy",
      configureServer(server) {
        // 在 Vite 开发服务器上注册 /cors-proxy 中间件，
        // 利用 Node.js 原生 http/https.request 进行流式管道转发，
        // 完全避免 fetch 的自动解压缩及 Header 长度不匹配问题。
        server.middlewares.use((req: IncomingMessage, res: ServerResponse, next: () => void) => {
          if (!req.url?.startsWith("/cors-proxy")) {
            return next();
          }

          // 从请求头中提取真实的目标 API 地址（Origin 部分）
          const targetOrigin = ((req.headers["x-target-url"] as string) || "").trim();
          if (!targetOrigin) {
            res.statusCode = 400;
            res.end("Missing X-Target-Url header");
            return;
          }

          // 去除 /cors-proxy 前缀，保留完整的 API 路径（如 /compatible-mode/v1/chat/completions）
          const apiPath = req.url.replace(/^\/cors-proxy/, "") || "/";

          // 拼装完整的目标 URL
          let targetUrl: URL;
          try {
            targetUrl = new URL(apiPath, targetOrigin);
          } catch (e) {
            res.statusCode = 400;
            res.end(`Invalid target URL: ${targetOrigin}${apiPath}`);
            return;
          }

          // 构建转发请求的 Headers，过滤掉浏览器专属字段
          const forwardHeaders: Record<string, string | string[] | undefined> = { ...req.headers };
          delete forwardHeaders["host"];
          delete forwardHeaders["x-target-url"];
          delete forwardHeaders["origin"];
          delete forwardHeaders["referer"];
          delete forwardHeaders["connection"];
          // 设置正确的 Host 头部
          forwardHeaders["host"] = targetUrl.host;

          // 选择 http 或 https 模块发起请求
          const isHttps = targetUrl.protocol === "https:";
          const doRequest = isHttps ? httpsRequest : httpRequest;

          const proxyReq = doRequest(
            {
              hostname: targetUrl.hostname,
              port: targetUrl.port || (isHttps ? 443 : 80),
              path: targetUrl.pathname + targetUrl.search,
              method: req.method,
              headers: forwardHeaders,
            },
            (proxyRes) => {
              // 将目标服务器的响应状态码和 Headers 原封不动地透传回浏览器
              res.writeHead(proxyRes.statusCode || 500, proxyRes.headers);
              // 使用流式管道直接转发响应体（不解压、不缓存、不修改）
              proxyRes.pipe(res);
            },
          );

          proxyReq.on("error", (err) => {
            console.error("CORS Proxy Error:", err);
            if (!res.headersSent) {
              res.statusCode = 502;
              res.end(`CORS Proxy Error: ${err.message}`);
            }
          });

          // 将浏览器的请求体通过流式管道直接转发给目标服务器
          req.pipe(proxyReq);
        });
      },
    },
  ],
});
