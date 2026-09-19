import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import express from "express";
import https from "node:https";
import { appendFileSync, existsSync, mkdirSync, readFileSync, renameSync, statSync } from "node:fs";
import { join, resolve } from "node:path";
import { format } from "node:util";
import devCerts from "office-addin-dev-certs";
import { availableProviders, getProvider, providerBaseURL, providerKey, providerReady } from "./providers.js";
import { serializeMessages, type InternalMessage } from "./protocol.js";
import { nextRouteAfterRejection, rememberRoute, routeFor, type OpenAiRoute } from "./openaiRoute.js";
import { buildResponsesBody, ResponsesTranslator, translateResponsesChunk, type ChatTool } from "./responsesApi.js";
import { isLoopbackAddress, isAllowedOrigin, isAllowedHost } from "./localOnly.js";
import { registerBackupRoutes } from "./backupRoutes.js";
import { MetricsStore, UsageScanner, formatMetricLine } from "./usageMetrics.js";

// Выпуск запускается из отдельного каталога, но конфигурация остаётся общей.
const projectRoot = process.env.EXCEL_AI_PROJECT_ROOT
  ? resolve(process.env.EXCEL_AI_PROJECT_ROOT)
  : fileURLToPath(new URL("../../", import.meta.url));
const envPath = join(projectRoot, "server", ".env");
const envLoaded = !dotenv.config({ path: envPath }).error;

// Долгоживущий процесс пишет сразу в ограниченный журнал. Перенаправление
// stdout в файл супервизором не подходит: пока процесс жив, файл не ротируется.
const logDir = join(projectRoot, "logs");
const logPath = join(logDir, "app.log");
const logLimit = 5 * 1024 * 1024;
const originalLog = console.log.bind(console);
const originalError = console.error.bind(console);
const originalWarn = console.warn.bind(console);
function appendLog(level: string, args: unknown[]): void {
  try {
    mkdirSync(logDir, { recursive: true });
    if (existsSync(logPath) && statSync(logPath).size >= logLimit) {
      for (let n = 2; n >= 1; n--) {
        const from = `${logPath}.${n}`;
        if (existsSync(from)) renameSync(from, `${logPath}.${n + 1}`);
      }
      renameSync(logPath, `${logPath}.1`);
    }
    appendFileSync(logPath, `[${new Date().toISOString()}] ${level} ${format(...args)}\n`, "utf8");
  } catch (error) {
    originalError("Не удалось записать журнал приложения:", error);
  }
}
console.log = (...args: unknown[]) => { appendLog("INFO", args); originalLog(...args); };
console.error = (...args: unknown[]) => { appendLog("ERROR", args); originalError(...args); };
// Без этого предупреждения видны только в консоли: в журнале их не было вовсе.
console.warn = (...args: unknown[]) => { appendLog("WARN", args); originalWarn(...args); };

const PORT = Number(process.env.PORT ?? 3000);

const distRoot = process.env.PANEL_DIST_DIR
  ? resolve(projectRoot, process.env.PANEL_DIST_DIR)
  : join(projectRoot, "dist");
if (!existsSync(join(distRoot, "taskpane.html"))) {
  throw new Error(`Рабочая сборка панели не найдена: ${distRoot}`);
}

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
// The running process pins a validated release even while dist/ is rebuilt.
const release = process.env.EXCEL_AI_RELEASE_ID || "development-dist";

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

app.use((req, res, next) => {
  if (isAllowedHost(req.headers.host, PORT)) return next();
  res.status(403).json({ error: { message: "Сторонний Host отклонён." } });
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
  res.json({ ok: true, app: APP_ID, version: buildVersion, release, pid: process.pid, startedAt })
);
app.get("/api/providers", (_req, res) => res.json(availableProviders()));

// Измерение расхода: только числа, имена провайдера и модели. Содержимому
// книги в записи метрики взяться неоткуда — см. usageMetrics.ts.
const metrics = new MetricsStore();
app.get("/api/metrics", (_req, res) => res.json({ summary: metrics.summary(), recent: metrics.all().slice(-25) }));

app.post("/api/chat", async (req, res) => {
  const { provider: providerId, model, messages, tools } = req.body ?? {};

  const provider = getProvider(String(providerId ?? ""));
  if (!provider) {
    return res.status(400).json({ error: { message: `Провайдер "${providerId}" неизвестен или отключён.` } });
  }

  const apiKey = providerKey(provider);
  if (!providerReady(provider)) {
    return res.status(400).json({
      error: {
        message: provider.keyOptional && provider.baseURLEnv
          ? `Не задан ${provider.baseURLEnv} в server/.env.`
          : `Не задан ${provider.envKey} в server/.env.`
      }
    });
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

    const chatBody = (effort: string | null) => ({
      model: selectedModel,
      messages: wireMessages,
      tools,
      stream: true,
      // Без явной просьбы OpenAI-совместимые провайдеры usage в потоке
      // не присылают, и измерять расход было бы нечем.
      stream_options: { include_usage: true },
      ...(provider.id === "deepseek"
        ? { thinking: { type: "enabled" }, reasoning_effort: "high" }
        : provider.id === "openai"
          ? { tool_choice: "auto", ...(effort === null ? {} : { reasoning_effort: effort }) }
        : { tool_choice: "auto" })
    });

    /** Повтор при обрыве соединения.
     *
     * У undici предел установки соединения — десять секунд, и менять его без
     * новой зависимости нельзя. При проверках DeepSeek дважды не уложился
     * и ронял задачу целиком. Повтор безопасен: ответ ещё не начинался,
     * сообщения провайдеру не доставлены. */
    const sendWithRetry = async (route: OpenAiRoute, attempts = 3): Promise<Response> => {
      for (let attempt = 1; ; attempt++) {
        try {
          return await send(route);
        } catch (error: any) {
          const connectionLost = error?.cause?.code === "UND_ERR_CONNECT_TIMEOUT" ||
            error?.cause?.code === "ECONNRESET" ||
            error?.cause?.code === "ETIMEDOUT";
          if (!connectionLost || attempt >= attempts || upstream.signal.aborted) throw error;
          console.warn(`[${provider.id}] соединение не установилось (${error?.cause?.code}), попытка ${attempt + 1} из ${attempts}`);
          await new Promise((resolve) => setTimeout(resolve, 500 * attempt));
        }
      }
    };

    const send = (route: OpenAiRoute) => fetch(
      `${providerBaseURL(provider)}${route.api === "responses" ? "/responses" : "/chat/completions"}`,
      {
        method: "POST",
        signal: upstream.signal,
        headers: {
          "Content-Type": "application/json",
          ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {}),
          ...(provider.headers ?? {})
        },
        body: JSON.stringify(
          route.api === "responses"
            ? buildResponsesBody(selectedModel, messages as InternalMessage[], tools as ChatTool[])
            : chatBody(route.effort)
        )
      }
    );

    // Ограничения зависят от модели: какое значение reasoning_effort она примет
    // и доступны ли ей функции на chat/completions. Вместо списка моделей
    // в коде маршрут выясняется из отказов провайдера и запоминается. Повтор
    // безопасен: ответ ещё не начинался, дублировать нечего.
    let route = provider.id === "openai" ? routeFor(provider.id, selectedModel) : { api: "chat" as const, effort: null };
    const tried: OpenAiRoute[] = [];
    const requestStartedAt = Date.now();
    let firstByteAt = 0;
    let responseBytes = 0;
    const usageScanner = new UsageScanner();
    let r = await sendWithRetry(route);

    while (!r.ok && provider.id === "openai" && tried.length < 3) {
      const text = await r.text().catch(() => "");
      const retry = nextRouteAfterRejection(r.status, text, route, tried);
      if (!retry) {
        console.error(`[${provider.id}] upstream HTTP ${r.status}`);
        return res.status(r.status).json({
          error: { message: `${provider.label} вернул ${r.status}. ${text.slice(0, 300)}` }
        });
      }
      console.warn(`[${provider.id}] ${selectedModel}: ${retry.reason}; повтор через ${retry.route.api}, reasoning_effort=${retry.route.effort ?? "не отправляем"}`);
      tried.push(route);
      route = retry.route;
      r = await sendWithRetry(route);
    }

    if (r.ok && provider.id === "openai" && tried.length > 0) rememberRoute(provider.id, selectedModel, route);

    if (!r.ok || !r.body) {
      const text = await r.text().catch(() => "");
      console.error(`[${provider.id}] upstream HTTP ${r.status}`);
      return res.status(r.status).json({
        error: { message: `${provider.label} вернул ${r.status}. ${text.slice(0, 300)}` }
      });
    }

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders?.();

    const reader = r.body.getReader();

    // Поток chat/completions панель понимает как есть. Поток /v1/responses
    // переводим здесь: устроен он иначе, но остальное приложение об этом
    // интерфейсе знать не должно.
    if (route.api === "responses") {
      const translator = new ResponsesTranslator();
      const decoder = new TextDecoder();
      let buffered = "";
      const emit = (pieces: string[]) => { for (const piece of pieces) res.write(piece); };
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        if (res.destroyed) {
          abortUpstream();
          break;
        }
        if (!firstByteAt) firstByteAt = Date.now();
        responseBytes += value.byteLength;
        const text = decoder.decode(value, { stream: true });
        usageScanner.push(text);
        buffered += text;
        const lines = buffered.split(/\r?\n/);
        buffered = lines.pop() ?? "";
        for (const line of lines) emit(translateResponsesChunk(translator, line));
      }
      buffered += decoder.decode();
      if (buffered.trim()) emit(translateResponsesChunk(translator, buffered));
      if (!res.destroyed) emit(translator.finalizeIfUnfinished());
      if (!res.writableEnded && !res.destroyed) res.end();
    } else {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        if (res.destroyed) {
          abortUpstream();
          break;
        }
        if (!firstByteAt) firstByteAt = Date.now();
        responseBytes += value.byteLength;
        // Поток отдаётся байт в байт; сканер только читает копию ради usage.
        usageScanner.push(new TextDecoder().decode(value, { stream: true }));
        res.write(Buffer.from(value));
      }
      if (!res.writableEnded && !res.destroyed) res.end();
    }

    const usage = usageScanner.finish();
    const metric = {
      provider: provider.id,
      model: selectedModel,
      api: route.api,
      requestBytes: Buffer.byteLength(JSON.stringify(
        route.api === "responses"
          ? buildResponsesBody(selectedModel, messages as InternalMessage[], tools as ChatTool[])
          : chatBody(route.effort)
      )),
      responseBytes,
      firstByteMs: firstByteAt ? firstByteAt - requestStartedAt : 0,
      totalMs: Date.now() - requestStartedAt,
      ok: true,
      attempts: tried.length + 1,
      ...(usage ? { usage } : {}),
      at: new Date(requestStartedAt).toISOString()
    };
    metrics.record(metric);
    console.log(formatMetricLine(metric));
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

registerBackupRoutes(app, projectRoot);

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
