import test from "node:test";
import assert from "node:assert/strict";
import http from "node:http";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { AddressInfo } from "node:net";
import { catalogDavHandler, catalogDavPort } from "./catalogDav.js";

async function withCatalog(run: (request: (method: string, path: string, headers?: Record<string, string>) => Promise<{ status: number; body: string; headers: http.IncomingHttpHeaders }>) => Promise<void>, enabled: () => boolean = () => true) {
  const root = mkdtempSync(join(tmpdir(), "catalog-"));
  writeFileSync(join(root, "manifest.xml"), "<OfficeApp>токен</OfficeApp>");
  writeFileSync(join(root, ".hidden"), "секрет");
  mkdirSync(join(root, "sub"));
  writeFileSync(join(root, "sub", "inner.xml"), "внутри");
  const server = http.createServer();
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = (server.address() as AddressInfo).port;
  server.on("request", catalogDavHandler(root, port, enabled));
  const request = (method: string, path: string, headers: Record<string, string> = {}) =>
    new Promise<{ status: number; body: string; headers: http.IncomingHttpHeaders }>((resolve, reject) => {
      const req = http.request({ host: "127.0.0.1", port, path, method, headers: { Host: `localhost:${port}`, ...headers } }, (res) => {
        let body = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => { body += chunk; });
        res.on("end", () => resolve({ status: res.statusCode ?? 0, body, headers: res.headers }));
      });
      req.on("error", reject);
      req.end();
    });
  try {
    await run(request);
  } finally {
    server.close();
    rmSync(root, { recursive: true, force: true });
  }
}

test("Windows WebDAV client sees the catalog folder and reads the manifest", async () => {
  await withCatalog(async (request) => {
    const options = await request("OPTIONS", "/catalog");
    assert.equal(options.status, 200);
    assert.equal(options.headers.dav, "1");

    const listing = await request("PROPFIND", "/catalog", { Depth: "1" });
    assert.equal(listing.status, 207);
    assert.match(listing.body, /<D:href>\/catalog\/manifest\.xml<\/D:href>/);
    assert.doesNotMatch(listing.body, /hidden|sub/, "только обычные файлы прямо в каталоге");

    const self = await request("PROPFIND", "/catalog", { Depth: "0" });
    assert.doesNotMatch(self.body, /manifest/);

    const root = await request("PROPFIND", "/", { Depth: "1" });
    assert.match(root.body, /<D:href>\/catalog\/<\/D:href>/);

    const file = await request("GET", "/catalog/manifest.xml");
    assert.equal(file.status, 200);
    assert.equal(file.body, "<OfficeApp>токен</OfficeApp>");
    assert.equal(file.headers["access-control-allow-origin"], undefined, "чужим страницам ответ не виден");
  });
});

test("nothing outside the catalog files is served", async () => {
  await withCatalog(async (request) => {
    for (const path of ["/catalog/.hidden", "/catalog/sub/inner.xml", "/catalog/sub", "/catalog/..%2F..%2Fwindows%2Fwin.ini", "/catalog/%2e%2e/secret", "/other", "/catalog/%E0%A4%A"]) {
      const res = await request("GET", path);
      assert.ok(res.status === 404 || res.status === 400, `${path}: ${res.status}`);
    }
  });
});

test("the catalog is read-only", async () => {
  await withCatalog(async (request) => {
    for (const method of ["PUT", "DELETE", "MKCOL", "MOVE", "COPY", "PROPPATCH", "LOCK", "POST"]) {
      assert.equal((await request(method, "/catalog/manifest.xml")).status, 405, method);
    }
  });
});

test("a page that rebinds its DNS to this port does not get the manifest", async () => {
  await withCatalog(async (request) => {
    assert.equal((await request("GET", "/catalog/manifest.xml", { Host: "evil.example:3080" })).status, 403);
    assert.equal((await request("PROPFIND", "/catalog", { Host: "evil.example" })).status, 403);
  });
});

test("port comes from server/.env and can be switched off", () => {
  assert.equal(catalogDavPort(undefined), 3080);
  assert.equal(catalogDavPort("3181"), 3181);
  assert.equal(catalogDavPort("off"), null);
  assert.equal(catalogDavPort("не порт"), 3080);
});

test("without the registration marker the catalog is not served at all", async () => {
  let marker = false;
  await withCatalog(async (request) => {
    // Регистрация через C$: манифест с токеном по WebDAV не виден никому.
    assert.equal((await request("GET", "/catalog/manifest.xml")).status, 404);
    assert.equal((await request("PROPFIND", "/catalog", { Depth: "1" })).status, 404);
    marker = true;
    assert.equal((await request("GET", "/catalog/manifest.xml")).status, 200);
  }, () => marker);
});
