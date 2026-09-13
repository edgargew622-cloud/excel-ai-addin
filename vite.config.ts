import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";
import devCerts from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  // Dev-сертификаты нужны только Vite dev server. Production build не должен
  // обращаться к хранилищу сертификатов/устанавливать их.
  const devHttps = command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;

  const localOnlyPlugin: Plugin = {
    name: "local-only",
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const address = req.socket.remoteAddress;
        if (address === "127.0.0.1" || address === "::1" || address === "::ffff:127.0.0.1") {
          next();
          return;
        }
        res.statusCode = 403;
        res.end("Local access only");
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
      port: 3000,
      strictPort: true,
      https: devHttps,
      proxy: {
        "/api": {
          target: "https://127.0.0.1:3001",
          changeOrigin: true,
          secure: false
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
