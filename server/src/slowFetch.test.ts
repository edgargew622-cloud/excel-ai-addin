import test from "node:test";
import assert from "node:assert/strict";
import http from "node:http";
import type { AddressInfo } from "node:net";
import { fetchWithoutHeaderTimeout } from "./slowFetch.js";

async function withServer(handler: http.RequestListener, run: (url: string) => Promise<void>) {
  const server = http.createServer(handler);
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  try {
    await run(`http://127.0.0.1:${(server.address() as AddressInfo).port}`);
  } finally {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
  }
}

test("posts the body and streams the answer back like fetch", async () => {
  await withServer((req, res) => {
    let body = "";
    req.on("data", (chunk) => { body += chunk; });
    req.on("end", () => {
      res.writeHead(200, { "Content-Type": "text/event-stream" });
      res.write(`data: ${req.headers["content-type"]} ${body}\n\n`);
      res.end("data: [DONE]\n\n");
    });
  }, async (url) => {
    const response = await fetchWithoutHeaderTimeout(`${url}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{\"model\":\"qwen3:8b\"}"
    });
    assert.equal(response.status, 200);
    assert.equal(response.headers.get("content-type"), "text/event-stream");
    assert.equal(await response.text(), "data: application/json {\"model\":\"qwen3:8b\"}\n\ndata: [DONE]\n\n");
  });
});

test("an error status is returned, not thrown", async () => {
  await withServer((_req, res) => { res.writeHead(500); res.end("model failed to load"); }, async (url) => {
    const response = await fetchWithoutHeaderTimeout(url, { method: "POST", body: "{}" });
    assert.equal(response.ok, false);
    assert.equal(await response.text(), "model failed to load");
  });
});

test("the panel's cancel stops a request that is still waiting for headers", async () => {
  await withServer(() => { /* модель «читает» и не отвечает */ }, async (url) => {
    const controller = new AbortController();
    const pending = fetchWithoutHeaderTimeout(url, { method: "POST", body: "{}", signal: controller.signal });
    setTimeout(() => controller.abort(), 50);
    await assert.rejects(pending, (error: any) => error?.name === "AbortError");
  });
});

test("nobody listening is reported like fetch does, with the cause code", async () => {
  await assert.rejects(fetchWithoutHeaderTimeout("http://127.0.0.1:9/"), (error: any) =>
    error instanceof TypeError && (error.cause as any)?.code === "ECONNREFUSED");
});
