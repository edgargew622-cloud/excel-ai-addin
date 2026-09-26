import test from "node:test";
import assert from "node:assert/strict";
import { compareVersions, parseVersion, RELEASES_API, UpdateChecker } from "./updateCheck.js";

const reply = (body: unknown, status = 200) => async () => new Response(JSON.stringify(body), { status });

test("versions are read from tags and compared as numbers", () => {
  assert.deepEqual(parseVersion("v1.0.10"), [1, 0, 10]);
  assert.equal(parseVersion("latest"), null);
  assert.ok(compareVersions([1, 0, 10], [1, 0, 9]) > 0);
  assert.equal(compareVersions([1, 2, 0], [1, 2, 0]), 0);
});

test("a newer release is reported with its page", async () => {
  const checker = new UpdateChecker("1.0.3", reply({ tag_name: "v1.0.4", html_url: "https://github.com/edgargew622-cloud/excel-ai-addin/releases/tag/v1.0.4" }));
  const info = await checker.check();
  assert.equal(info.newer, true);
  assert.equal(info.latest, "1.0.4");
  assert.match(info.url ?? "", /releases\/tag\/v1\.0\.4$/);
});

test("the same or an older release is not news; a foreign link is dropped", async () => {
  assert.equal((await new UpdateChecker("1.0.3", reply({ tag_name: "v1.0.3" })).check()).newer, false);
  const info = await new UpdateChecker("1.0.3", reply({ tag_name: "v2.0.0", html_url: "https://evil.example/x" })).check();
  assert.equal(info.newer, true);
  assert.equal(info.url, undefined);
});

test("GitHub is asked at most once a day; no network changes nothing", async () => {
  let calls = 0;
  let time = 0;
  const checker = new UpdateChecker("1.0.3", async (url) => { calls += 1; assert.equal(url, RELEASES_API); return new Response(JSON.stringify({ tag_name: "v1.0.3" })); }, () => time);
  await checker.check();
  time += 60 * 60_000;
  await checker.check();
  assert.equal(calls, 1);
  time += 24 * 60 * 60_000;
  await checker.check();
  assert.equal(calls, 2);
  const offline = await new UpdateChecker("1.0.3", async () => { throw new Error("нет сети"); }).check();
  assert.deepEqual(offline, { current: "1.0.3", checked: false });
});

test("the check can be switched off and then asks nobody", async () => {
  let calls = 0;
  const info = await new UpdateChecker("1.0.3", async () => { calls += 1; return new Response("{}"); }, Date.now, true).check();
  assert.equal(info.disabled, true);
  assert.equal(calls, 0);
});
