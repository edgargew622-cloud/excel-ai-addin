import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";
import type { ClientRequest } from "node:http";
import devCerts from "office-addin-dev-certs";
import { isAllowedHost, isAllowedOrigin, isLoopbackAddress } from "./server/src/localOnly";

export default defineConfig(async ({ command }) => {
  // Dev-сертификаты нужны только Vite dev server. Production build не должен
  // обращаться к хранилищу сертификатов/устанавливать их.
  const devHttps = command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;

  const localOnlyPlugin: Plugin = {
    name: "local-only",
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        if (!isLoopbackAddress(req.socket.remoteAddress) || !isAllowedHost(req.headers.host, 3100)) {
          res.statusCode = 403;
          res.end("Local access only");
          return;
        }
        if (req.url?.startsWith("/api") && !isAllowedOrigin(req.headers.origin, 3100)) {
          res.statusCode = 403;
          res.end("Foreign Origin");
          return;
        }
        next();
      });
    }
  };

  return {
    root: ".",
    plugins: [react(), localOnlyPlugin],
    server: {
      // Excel WebView may resolve localhost to either IPv4 or IPv6.
      // Binding to the IPv6 wildcard keeps both loopback variants reachable.
      host: "::",
      watch: { ignored: ["**/releases/**", "**/server/releases/**"] },
      // Рабочий сервер занимает 3000 и раздаёт собранную панель вместе с /api.
      // Разработка живёт на отдельном порту, поэтому обе панели доступны
      // одновременно и их поведение можно сравнить на одной книге.
      port: 3100,
      strictPort: true,
      https: devHttps,
      proxy: {
        // API берём у рабочего сервера: он единственный держит ключи и
        // провайдеров. Отдельного API для разработки панели не нужно.
        "/api": {
          target: "https://127.0.0.1:3000",
          changeOrigin: true,
          secure: false,
          // Рабочий сервер отклоняет сторонний Origin: панель и API у него
          // одного происхождения. Наш dev-прокси и есть тот же самый источник,
          // поэтому подменяем Origin на целевой, а не ослабляем проверку на
          // сервере ради разработки.
          configure(proxy: {
            on: (event: "proxyReq", listener: (proxyReq: ClientRequest) => void) => void;
          }) {
            proxy.on("proxyReq", (proxyReq) => {
              proxyReq.setHeader("origin", "https://localhost:3000");
            });
          }
        }
      }
    },
    build: {
      outDir: "dist",
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, "taskpane.html"),
          commands: resolve(__dirname, "commands.html")
        }
      }
    }
  };
});
