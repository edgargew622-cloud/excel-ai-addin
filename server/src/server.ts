import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import express from "express";
import cors from "cors";
import https from "node:https";
import devCerts from "office-addin-dev-certs";
import { availableProviders, getProvider } from "./providers.js";
import { serializeMessages, type InternalMessage } from "./protocol.js";

// Путь к .env задаётся относительно этого файла, а не рабочего каталога:
// запуск из другой папки не должен молча менять конфигурацию. И из src/, и из
// собранного dist/ этот относительный путь ведёт в один и тот же server/.env.
const envPath = fileURLToPath(new URL("../.env", import.meta.url));
const envLoaded = !dotenv.config({ path: envPath }).error;

const app = express();
app.disable("x-powered-by");
app.use(express.json({ limit: "8mb" }));

const allowedOrigins = (process.env.ALLOWED_ORIGINS ?? "https://localhost:3000")
  .split(",")
  .map((x) => x.trim())
  .filter(Boolean);

app.use(
  cors({
    origin: allowedOrigins,
    credentials: false
  })
);

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

app.get("/api/health", (_req, res) => res.json({ ok: true }));
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

const port = Number(process.env.PORT ?? 3001);
const host = process.env.HOST?.trim() || "127.0.0.1";

// Локальный backend нужен по HTTPS: Office webview не разрешает mixed content.
const { cert, key } = await devCerts.getHttpsServerOptions();
https.createServer({ cert, key }, app).listen(port, host, () => {
  const ready = availableProviders().map((p) => p.id);
  console.log(`Прокси на https://${host}:${port}`);
  console.log(envLoaded ? `Конфигурация: ${envPath}` : `Конфигурация не найдена: ${envPath}`);
  console.log(ready.length ? `Ключи найдены: ${ready.join(", ")}` : "Ключей нет — заполните server/.env");
});
