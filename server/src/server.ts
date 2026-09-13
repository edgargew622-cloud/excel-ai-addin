import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import express from "express";
import https from "node:https";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import devCerts from "office-addin-dev-certs";
import { availableProviders, getProvider } from "./providers.js";
import { serializeMessages, type InternalMessage } from "./protocol.js";
import { isLoopbackAddress, isAllowedOrigin } from "./localOnly.js";

// Путь к .env задаётся относительно этого файла, а не рабочего каталога:
// запуск из другой папки не должен молча менять конфигурацию. И из src/, и из
// собранного dist/ этот относительный путь ведёт в один и тот же server/.env.
const envPath = fileURLToPath(new URL("../.env", import.meta.url));
const envLoaded = !dotenv.config({ path: envPath }).error;

const PORT = Number(process.env.PORT ?? 3000);

// Корень проекта относительно собранного server/dist/ и относительно server/src/.
const projectRoot = fileURLToPath(new URL("../../", import.meta.url));
const distRoot = join(projectRoot, "dist");

const APP_ID = "excel-ai-addin";
const buildVersion = (() => {
  try {
    const pkg = JSON.parse(readFileSync(join(projectRoot, "package.json"), "utf8"));
    return String(pkg.version ?? "0.0.0");
  } catch {
    return "unknown";
  }
})();
const startedAt = new Date().toISOString();

const app = express();
app.disable("x-powered-by");

// Единственная настоящая граница доступа: запросы только с этого компьютера.
// Раньше эта проверка жила в dev-плагине Vite и не действовала для собранной
// панели. Теперь она в рабочем сервере и распространяется на всё, включая
// статику.
app.use((req, res, next) => {
  if (isLoopbackAddress(req.socket.remoteAddress)) return next();
  res.status(403).type("text/plain").send("Local access only");
});

// Панель и API теперь одного происхождения, поэтому разрешать сторонние
// источники не нужно — их нужно отклонять. Собственный запрос либо не имеет
// Origin, либо имеет свой.
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (isAllowedOrigin(origin, PORT)) return next();
  res.status(403).json({ error: { message: "Сторонний Origin отклонён." } });
});

app.use(express.json({ limit: "8mb" }));

// Простейший локальный rate limit без внешней зависимости. Для публичного
// deployment замените на Redis/reverse-proxy limiter.
const WINDOW_MS = 60_000;
const MAX_REQUESTS = Number(process.env.RATE_LIMIT_PER_MINUTE ?? 60);
const buckets = new Map<string, { startedAt: number; count: number }>();
app.use("/api", (req, res, next) => {
  const now = Date.now();
  const key = req.ip || req.socket.remoteAddress || "local";
  const current = buckets.get(key);
  if (!current || now - current.startedAt >= WINDOW_MS) {
    buckets.set(key, { startedAt: now, count: 1 });
    return next();
  }
  current.count += 1;
  if (current.count > MAX_REQUESTS) {
    return res.status(429).json({ error: { message: "Слишком много запросов. Повторите через минуту." } });
  }
  next();
});

// Идентификация нужна супервизору: занятый порт может принадлежать другой
// программе, и тогда её нельзя ни считать своим экземпляром, ни завершать.
app.get("/api/health", (_req, res) =>
  res.json({ ok: true, app: APP_ID, version: buildVersion, pid: process.pid, startedAt })
);
app.get("/api/providers", (_req, res) => res.json(availableProviders()));

app.post("/api/chat", async (req, res) => {
  const { provider: providerId, model, messages, tools } = req.body ?? {};

  const provider = getProvider(String(providerId ?? ""));
  if (!provider) {
    return res.status(400).json({ error: { message: `Провайдер "${providerId}" неизвестен или отключён.` } });
  }

  const apiKey = process.env[provider.envKey]?.trim();
  if (!apiKey) {
    return res.status(400).json({ error: { message: `Не задан ${provider.envKey} в server/.env.` } });
  }
  if (!Array.isArray(messages) || !messages.length) {
    return res.status(400).json({ error: { message: "messages пуст." } });
  }
  if (!Array.isArray(tools)) {
    return res.status(400).json({ error: { message: "tools должен быть массивом." } });
  }

  const selectedModel = String(model || provider.defaultModel);
  if (!provider.models.includes(selectedModel)) {
    return res.status(400).json({
      error: { message: `Модель "${selectedModel}" не разрешена для провайдера ${provider.label}.` }
    });
  }

  const upstream = new AbortController();
  const abortUpstream = () => {
    if (!upstream.signal.aborted) upstream.abort();
  };

  // req.close нельзя использовать как признак ухода клиента: он возникает и
  // после штатного получения request body. Aborted request и преждевременно
  // закрытый RESPONSE — корректные сигналы отмены upstream.
  req.once("aborted", abortUpstream);
  const onResponseClose = () => {
    if (!res.writableEnded) abortUpstream();
  };
  res.once("close", onResponseClose);

  try {
    const wireMessages = serializeMessages(messages as InternalMessage[], provider.id);
    const requestBody = {
      model: selectedModel,
      messages: wireMessages,
      tools,
      stream: true,
      ...(provider.id === "deepseek"
        ? { thinking: { type: "enabled" }, reasoning_effort: "high" }
        : provider.id === "openai"
          ? { tool_choice: "auto", reasoning_effort: "none" }
        : { tool_choice: "auto" })
    };

    const r = await fetch(`${provider.baseURL}/chat/completions`, {
      method: "POST",
      signal: upstream.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
        ...(provider.headers ?? {})
      },
      body: JSON.stringify(requestBody)
    });

    if (!r.ok || !r.body) {
      const text = await r.text().catch(() => "");
      console.error(`[${provider.id}] ${r.status}: ${text.slice(0, 500)}`);
      return res.status(r.status).json({
        error: { message: `${provider.label} вернул ${r.status}. ${text.slice(0, 300)}` }
      });
    }

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders?.();

    const reader = r.body.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (res.destroyed) {
        abortUpstream();
        break;
      }
      res.write(Buffer.from(value));
    }
    if (!res.writableEnded && !res.destroyed) res.end();
  } catch (e: any) {
    if (e?.name === "AbortError") {
      if (!res.writableEnded && !res.destroyed) res.end();
      return;
    }
    console.error(e);
    if (res.headersSent) {
      if (!res.destroyed) {
        res.write(`data: ${JSON.stringify({ error: { message: String(e?.message ?? e) } })}\n\n`);
        res.end();
      }
      return;
    }
    return res.status(500).json({ error: { message: String(e?.message ?? e) } });
  } finally {
    req.off("aborted", abortUpstream);
    res.off("close", onResponseClose);
  }
});

// Раздаём строго каталог сборки. Исходники, server/.env и сертификаты в него
// не попадают по построению: express.static не выходит за пределы корня.
app.use(
  express.static(distRoot, {
    index: false,
    dotfiles: "deny",
    setHeaders: (res) => res.setHeader("Cache-Control", "no-store")
  })
);

app.use((req, res) => {
  if (req.path.startsWith("/api/")) {
    return res.status(404).json({ error: { message: `Нет эндпоинта ${req.path}.` } });
  }
  res
    .status(404)
    .type("text/plain")
    .send(
      `Файл ${req.path} не найден в сборке. Панель собрана? Выполните npm run build.`
    );
});

// Office webview не разрешает mixed content, поэтому нужен HTTPS.
const { cert, key } = await devCerts.getHttpsServerOptions();

// Excel резолвит localhost то в IPv4, то в IPv6. Слушаем оба адреса петли
// вместо привязки к «всем интерфейсам»: так наружу не открывается ничего,
// а не открывается-и-отклоняется проверкой.
const LOOPBACKS = ["127.0.0.1", "::1"];
let listening = 0;
for (const host of LOOPBACKS) {
  const server = https.createServer({ cert, key }, app);
  server.on("error", (error: NodeJS.ErrnoException) => {
    if (error.code === "EADDRINUSE") {
      console.error(`Порт ${PORT} на ${host} уже занят. Второй экземпляр не запускается.`);
      process.exit(10);
    }
    if (error.code === "EAFNOSUPPORT" || error.code === "EADDRNOTAVAIL") {
      console.warn(`Адрес ${host} недоступен в этой системе, пропускаем.`);
      return;
    }
    console.error(`Ошибка прослушивания ${host}:${PORT}: ${error.message}`);
    process.exit(11);
  });
  server.listen(PORT, host, () => {
    listening += 1;
    console.log(`Слушаю https://${host === "::1" ? "[::1]" : host}:${PORT}`);
    if (listening === 1) {
      const ready = availableProviders().map((p) => p.id);
      console.log(`Панель: https://localhost:${PORT}/taskpane.html (сборка ${buildVersion})`);
      console.log(`Статика: ${distRoot}`);
      console.log(envLoaded ? `Конфигурация: ${envPath}` : `Конфигурация не найдена: ${envPath}`);
      console.log(ready.length ? `Ключи найдены: ${ready.join(", ")}` : "Ключей нет — заполните server/.env");
    }
  });
}
