import { before, beforeEach, after, test } from "node:test";
import assert from "node:assert/strict";
import { randomBytes, createHash } from "node:crypto";
import { generateKeyPair, SignJWT, exportJWK, createLocalJWKSet, type JWTPayload } from "jose";
import { poolFor, id, sha256, validateLocalDatabase, type Context } from "./core.js";
import { migrate } from "./db.js";
import { seedDevelopment } from "./seed.js";
import { DevKeyringProtector, CredentialStore } from "./secrets.js";
import { EntraAuth, SessionIdentity } from "./auth.js";
import { Tasks } from "./tasks.js";
import { Ledger } from "./ledger.js";
import { ChatService } from "./service.js";
import { FakeAgentProvider, parseChat } from "./providers.js";
import { DeepSeekPilot, type PilotPricing } from "./deepseekPilot.js";
import { createCommercialApp } from "./app.js";

const url = process.env.COMMERCIAL_TEST_DATABASE_URL;
if (!url || !new URL(url).pathname.endsWith("_c1_test")) throw new Error("An isolated *_c1_test database is required; the commercial fixture schema is reset.");
const pool = poolFor(url), tasks = new Tasks(pool), ledger = new Ledger(pool);
const protector = new DevKeyringProtector(new Map([["test", randomBytes(32)]]), "test");
const credentials = new CredentialStore(pool, protector);
const config = { tenant: "00000000-0000-4000-8000-000000000001", clientId: "00000000-0000-4000-8000-000000000002", clientSecret: "OIDC_SECRET_SENTINEL", origin: "https://localhost:3100" };
let key: Awaited<ReturnType<typeof generateKeyPair>>, a: Context, b: Context;
const issuer = `https://login.microsoftonline.com/${config.tenant}/v2.0`;
const objectId = "00000000-0000-4000-8000-000000000003";
let fixtureToken: string;
const body = (taskId: string) => ({ provider: "fake", model: "c2-pilot", keyMode: "hosted", taskId, messages: [{ role: "user", content: "synthetic request" }] });
before(async () => { await migrate(pool); key = await generateKeyPair("RS256"); });
beforeEach(async () => {
  await pool.query("TRUNCATE commercial.users,commercial.organizations,commercial.budget_accounts,commercial.auth_flows,commercial.audit_events CASCADE");
  const first = await seedDevelopment(pool), second = await seedDevelopment(pool);
  a = { ...first, role: "owner" }; b = { ...second, role: "owner" }; fixtureToken = first.token;
  await pool.query("UPDATE commercial.users SET external_subject=$1 WHERE id=$2", [`${issuer}|${objectId}`, a.userId]);
  await pool.query("UPDATE commercial.budget_accounts SET cap=10000,max_inflight=200");
});
after(async () => { await pool.end(); });
async function rejection(promise: Promise<unknown>, code: string) { await assert.rejects(promise, (e: any) => e.code === code); }

async function flow(overrides: JWTPayload = {}, badSignature = false) {
  let fetches = 0, pkce = false;
  const jwk = await exportJWK(key.publicKey); jwk.kid = "pilot";
  let jwt = "";
  const auth = new EntraAuth(pool, protector, config, (async (_url, init) => {
    fetches++;
    const params = init!.body as URLSearchParams;
    pkce = createHash("sha256").update(params.get("code_verifier")!).digest("base64url") === challenge;
    assert.equal(params.get("client_secret"), config.clientSecret);
    return Response.json({ id_token: jwt, access_token: "NEVER_FORWARD_PROVIDER_ACCESS_TOKEN" });
  }) as typeof fetch, createLocalJWKSet({ keys: [jwk] }));
  const proof = randomBytes(32).toString("base64url");
  const started = await auth.start(sha256(proof));
  const state = new URL(started.dialogUrl).searchParams.get("state")!;
  const authorize = new URL(await auth.authorize(state));
  const challenge = authorize.searchParams.get("code_challenge")!;
  assert.equal(authorize.searchParams.get("code_challenge_method"), "S256");
  assert.equal(authorize.searchParams.get("redirect_uri"), `${config.origin}/api/v1/auth/callback`);
  const signKey = badSignature ? (await generateKeyPair("RS256")).privateKey : key.privateKey;
  jwt = await new SignJWT({ nonce: authorize.searchParams.get("nonce"), tid: config.tenant, oid: objectId, ...overrides })
    .setProtectedHeader({ alg: "RS256", kid: "pilot" }).setIssuer(issuer).setAudience(config.clientId).setSubject("pilot-subject").setIssuedAt().setExpirationTime("5m").sign(signKey);
  // Explicit signed claim overrides after defaults for negative issuer/audience/expiry tests.
  if (Object.keys(overrides).some(k => ["iss", "aud", "exp"].includes(k))) {
    jwt = await new SignJWT({ iss: issuer, aud: config.clientId, sub: "pilot-subject", iat: Math.floor(Date.now()/1000), exp: Math.floor(Date.now()/1000)+300,
      nonce: authorize.searchParams.get("nonce"), tid: config.tenant, oid: objectId, ...overrides }).setProtectedHeader({ alg: "RS256", kid: "pilot" }).sign(signKey);
  }
  return { auth, state, proof, fetches: () => fetches, pkce: () => pkce };
}

test("PostgreSQL query/encoded/duplicate connection overrides fail without constructing a network connection", () => {
  for (const suffix of ["?host=example.invalid", "?%68ost=example.invalid", "?host=127.0.0.1&host=example.invalid", "?hostaddr=203.0.113.1", "?port=1", "?service=x", "?sslmode=disable", "#host=example.invalid"]) {
    assert.throws(() => validateLocalDatabase(`postgresql://u:p@127.0.0.1/test${suffix}`));
    assert.throws(() => poolFor(`postgresql://u:p@127.0.0.1/test${suffix}`));
  }
  validateLocalDatabase("postgresql://u:p%3Fhost%3Dx@127.0.0.1:55432/test");
});

test("Entra signed identity + state + nonce + PKCE + parent proof create one revocable server session", async () => {
  const f = await flow();
  await rejection(f.auth.callback(randomBytes(32).toString("base64url"), "code"), "invalid_auth_state");
  assert.equal(f.fetches(), 0);
  const ticket = await f.auth.callback(f.state, "synthetic-code");
  assert.equal(f.pkce(), true);
  await rejection(f.auth.callback(f.state, "synthetic-code"), "invalid_auth_state");
  await rejection(f.auth.redeem(ticket, randomBytes(32).toString("base64url")), "invalid_auth_ticket");
  const session = await f.auth.redeem(ticket, f.proof), identity = new SessionIdentity(pool);
  assert.equal(await identity.verify(session.token), a.userId);
  await rejection(identity.verify(fixtureToken), "unauthorized");
  await rejection(f.auth.redeem(ticket, f.proof), "invalid_auth_ticket");
  await pool.query("UPDATE commercial.auth_sessions SET revoked=true");
  await rejection(identity.verify(session.token), "unauthorized");
  const dump = JSON.stringify((await pool.query("SELECT * FROM commercial.auth_sessions")).rows);
  assert.ok(!dump.includes(session.token));
  assert.ok(!dump.includes("NEVER_FORWARD"));
});

test("forged signature, issuer, audience, expiry, nonce and tenant cannot authenticate", async () => {
  for (const [claims, bad] of [[{}, true], [{iss:"https://attacker.invalid"},false], [{aud:"other-app"},false], [{exp:1},false], [{nonce:"wrong"},false], [{tid:"other-tenant"},false]] as Array<[JWTPayload,boolean]>) {
    const f = await flow(claims, bad);
    await rejection(f.auth.callback(f.state, "code"), "invalid_identity_token");
  }
  assert.equal((await pool.query("SELECT count(*)::int n FROM commercial.auth_sessions")).rows[0].n, 0);
});

test("expired state/ticket/session and inactive/uninvited identity fail closed", async () => {
  const expired = await flow(); await pool.query("UPDATE commercial.auth_flows SET expires_at=now()-interval '1 second'");
  await rejection(expired.auth.callback(expired.state,"code"),"invalid_auth_state");
  const uninvited = await flow(); await pool.query("UPDATE commercial.users SET active=false WHERE id=$1",[a.userId]);
  await rejection(uninvited.auth.callback(uninvited.state,"code"),"pilot_invitation_required");
  await pool.query("UPDATE commercial.users SET active=true");
  const f = await flow(), ticket = await f.auth.callback(f.state,"code");
  await pool.query("UPDATE commercial.auth_redemptions SET expires_at=now()-interval '1 second'");
  await rejection(f.auth.redeem(ticket,f.proof),"invalid_auth_ticket");
  const g = await flow(), session = await g.auth.redeem(await g.auth.callback(g.state,"code"),g.proof);
  await pool.query("UPDATE commercial.auth_sessions SET expires_at=now()-interval '1 second'");
  await rejection(new SessionIdentity(pool).verify(session.token),"unauthorized");
});

test("all steps share durable task spend; concurrent duplicate dispatch is single and tenants cannot read tasks", async () => {
  const task = await tasks.create(a,"hosted"), provider = new FakeAgentProvider();
  const service = new ChatService(ledger,credentials,new Map([["fake",provider]])), key=id();
  await Promise.all(Array.from({length:100},()=>service.chat(a,body(task.id),key,new AbortController().signal)));
  const second = await service.chat(a,body(task.id),id(),new AbortController().signal);
  assert.equal(second.request.state,"settled"); assert.equal(provider.calls,2);
  const state = await tasks.get(a,task.id);
  assert.equal(state.steps,2); assert.equal(state.spent,"6"); assert.equal(state.reserved,"0"); assert.equal(state.requests.length,2);
  await rejection(tasks.get(b,task.id),"task_not_found");
  await rejection(service.chat(b,body(task.id),id(),new AbortController().signal),"task_not_found");
});

test("task cost, step, time, mode and cancellation gates stop calls before dispatch", async () => {
  const provider = new FakeAgentProvider(), service = new ChatService(ledger,credentials,new Map([["fake",provider]]));
  const task = await tasks.create(a,"hosted");
  await pool.query("UPDATE commercial.tasks SET cap=10 WHERE id=$1",[task.id]);
  await service.chat(a,body(task.id),id(),new AbortController().signal);
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"task_budget_exhausted");
  await pool.query("UPDATE commercial.tasks SET cap=200,max_steps=1 WHERE id=$1",[task.id]);
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"task_step_limit");
  await pool.query("UPDATE commercial.tasks SET max_steps=12,deadline=now()-interval '1 second' WHERE id=$1",[task.id]);
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"task_time_limit");
  await pool.query("UPDATE commercial.tasks SET deadline=now()+interval '1 minute' WHERE id=$1",[task.id]);
  await tasks.finish(a,task.id,"cancelled");
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"task_stopped");
  assert.equal(provider.calls,1);
});

test("task continuation rechecks membership, subscription and mode; BYOK uses separate task service units", async () => {
  const task=await tasks.create(a,"hosted"),fake=new FakeAgentProvider(),service=new ChatService(ledger,credentials,new Map([["fake",fake]]));
  await pool.query("UPDATE commercial.subscriptions SET status='suspended' WHERE org_id=$1",[a.orgId]);
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"entitlement_required");
  await pool.query("UPDATE commercial.subscriptions SET status='active' WHERE org_id=$1",[a.orgId]);
  await pool.query("UPDATE commercial.memberships SET role='viewer' WHERE org_id=$1 AND user_id=$2",[a.orgId,a.userId]);
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"role_cannot_spend");
  await pool.query("UPDATE commercial.memberships SET role='owner' WHERE org_id=$1 AND user_id=$2",[a.orgId,a.userId]);
  const credential=await credentials.create(a,"fake","synthetic-byok-test-key");
  await rejection(service.chat(a,{...body(task.id),keyMode:"byok",credentialId:credential.id},id(),new AbortController().signal),"task_mode_mismatch");
  const byok=await tasks.create(a,"byok");
  await service.chat(a,{...body(byok.id),keyMode:"byok",credentialId:credential.id},id(),new AbortController().signal);
  assert.equal((await tasks.get(a,byok.id)).spent,"1");
  assert.equal((await tasks.get(a,task.id)).spent,"0");
  const money=await pool.query("SELECT sum(spent)::text spent FROM commercial.budget_accounts WHERE meter='usd_micro'");assert.equal(money.rows[0].spent,"0");
});

test("scripted provider carries enable-tools and Excel results across the full metered task protocol",async()=>{
  const task=await tasks.create(a,"hosted"),fake=new FakeAgentProvider(),service=new ChatService(ledger,credentials,new Map([["fake",fake]]));
  const initial=body(task.id);
  const first=await service.chat(a,{...initial,toolNames:["enable_tools"]},id(),new AbortController().signal);
  const calls="toolCalls" in first?first.toolCalls!:[];assert.equal(calls[0].name,"enable_tools");
  const next={...initial,toolNames:["reconcile_ranges"],messages:[...initial.messages,{role:"assistant",content:first.text,tool_calls:calls},{role:"tool",tool_call_id:calls[0].id,content:'{"ok":true}'}]};
  const second=await service.chat(a,next,id(),new AbortController().signal);
  const reconciliation="toolCalls" in second?second.toolCalls!:[];assert.equal(reconciliation[0].name,"reconcile_ranges");
  const final=await service.chat(a,{...next,messages:[...next.messages,{role:"assistant",content:second.text,tool_calls:reconciliation},{role:"tool",tool_call_id:reconciliation[0].id,content:'{"ok":true,"executionState":"verified"}'}]},id(),new AbortController().signal);
  assert.equal("finishReason" in final?final.finishReason:undefined,"stop");
  const result=await tasks.get(a,task.id);assert.equal(result.steps,3);assert.equal(result.spent,"9");assert.equal(result.requests.length,3);assert.equal(fake.calls,3);
});

test("uncertain task preserves funds across a new pool and blocks later steps; pre-dispatch cancellation releases", async () => {
  const task = await tasks.create(a,"hosted");
  const reservation = await ledger.reserve(a,{idempotencyKey:id(),bodyHash:"synthetic",provider:"fake",model:"c2-pilot",mode:"hosted",providerBound:10n,pricingVersion:"fake",taskId:task.id});
  await ledger.dispatch(a,reservation.request.id);
  await pool.query("UPDATE commercial.usage_requests SET updated_at=now()-interval '1 hour'");
  const restarted=poolFor(url!); try { await new Ledger(restarted).recoverStale(); } finally { await restarted.end(); }
  assert.equal((await tasks.get(a,task.id)).reserved,"10");
  const service=new ChatService(ledger,credentials,new Map([["fake",new FakeAgentProvider()]]));
  await rejection(service.chat(a,body(task.id),id(),new AbortController().signal),"task_request_pending");
  await ledger.settle(a,reservation.request.id,3n,{promptTokens:1,completionTokens:1});
  assert.equal((await tasks.get(a,task.id)).spent,"3");
  const aborted=new AbortController();aborted.abort();
  await assert.rejects(service.chat(a,body(task.id),id(),aborted.signal));
  assert.equal((await tasks.get(a,task.id)).reserved,"0");
});

test("C2 HTTP uses signed session only, requires task/origin, emits standard tool SSE and metadata only on replay", async () => {
  const f=await flow(), session=await f.auth.redeem(await f.auth.callback(f.state,"code"),f.proof);
  const logs: unknown[]=[];
  const {app}=createCommercialApp({pool,identity:new SessionIdentity(pool),credentials,allowFake:true,c2:true,panelOrigin:config.origin,auth:f.auth,adapters:new Map([["fake",new FakeAgentProvider()]]),log:m=>logs.push(m)});
  const server=app.listen(0,"127.0.0.1");await new Promise<void>(r=>server.once("listening",r));
  const address=server.address() as {port:number},base=`http://127.0.0.1:${address.port}/api/v1`;
  const post=(path:string,body:unknown,token=session.token,key=id(),origin=config.origin)=>fetch(base+path,{method:"POST",headers:{Authorization:`Bearer ${token}`,Origin:origin,"Content-Type":"application/json","Idempotency-Key":key},body:JSON.stringify(body)});
  try {
    assert.equal((await post(`/orgs/${a.orgId}/tasks`,{keyMode:"hosted"},fixtureToken)).status,401);
    assert.equal((await post(`/orgs/${a.orgId}/tasks`,{keyMode:"hosted"},session.token,id(),"https://foreign.invalid")).status,403);
    const task=await (await post(`/orgs/${a.orgId}/tasks`,{keyMode:"hosted"})).json() as {id:string};
    const chat={...body(task.id),toolNames:["enable_tools"]},dedup=id();
    const text=await (await post(`/orgs/${a.orgId}/chat`,chat,session.token,dedup)).text();
    assert.match(text,/tool_calls/); assert.match(text,/enable_tools/); assert.match(text,/\[DONE\]/);
    const replay=await (await post(`/orgs/${a.orgId}/chat`,chat,session.token,dedup)).text();
    assert.match(replay,/request_state_only/); assert.doesNotMatch(replay,/arguments/);
    assert.equal((await post(`/orgs/${b.orgId}/chat`,chat)).status,404);
    assert.equal((await post(`/orgs/${a.orgId}/chat`,{...chat,taskId:undefined})).status,400);
    assert.equal((await post(`/orgs/${a.orgId}/chat`,{...chat,price:0})).status,400);
    await post("/auth/logout",{});
    assert.equal((await post(`/orgs/${a.orgId}/tasks`,{keyMode:"hosted"})).status,401);
    assert.doesNotMatch(JSON.stringify(logs),/OIDC_SECRET|synthetic request|Bearer|NEVER_FORWARD/);
    assert.ok(!JSON.stringify(logs).includes(session.token));
  } finally { await new Promise<void>(r=>server.close(()=>r())); }
});

test("single real-model adapter is disabled by default; contract fake verifies fixed URL, secret source, pricing and one attempt",async()=>{
  const price:PilotPricing={model:"approved-pilot-model",version:"synthetic-price-v1",inputMicroUsdPerMillion:1_000_000n,outputMicroUsdPerMillion:2_000_000n,maxOutputTokens:1024,validUntil:Date.now()+60_000};
  let calls=0; const hosted={resolve:async()=>"PLATFORM_SECRET_SENTINEL"};
  const fetcher=(async(url,init)=>{calls++; assert.equal(url,"https://api.deepseek.com/chat/completions"); assert.equal(init!.redirect,"error");
    assert.equal((init!.headers as Record<string,string>).Authorization,"Bearer PLATFORM_SECRET_SENTINEL");
    assert.equal(JSON.parse(init!.body as string).max_tokens,1024);
    return Response.json({choices:[{finish_reason:"stop",message:{content:"synthetic"}}],usage:{prompt_tokens:3,completion_tokens:2}});
  }) as typeof fetch;
  const input=parseChat({provider:"deepseek",model:price.model,keyMode:"hosted",messages:[{role:"user",content:"synthetic"}]});
  const disabled=new DeepSeekPilot(price,hosted,false,fetcher);
  await rejection(disabled.run(input,{signal:new AbortController().signal,attemptId:id(),quote:disabled.quote(input)}),"real_spend_disabled");assert.equal(calls,0);
  const adapter=new DeepSeekPilot(price,hosted,true,fetcher),quote=adapter.quote(input);
  const result=await adapter.run(input,{signal:new AbortController().signal,attemptId:id(),quote});assert.equal(result.cost,7n);assert.ok(quote.maxCost>=7n);assert.equal(calls,1);
  const missing=new DeepSeekPilot(price,hosted,true,(async()=>Response.json({choices:[{finish_reason:"stop",message:{content:"synthetic"}}]})) as typeof fetch);
  assert.equal((await missing.run(input,{signal:new AbortController().signal,attemptId:id(),quote:missing.quote(input)})).cost,null);
  const byok=new DeepSeekPilot(price,{resolve:async()=>{throw new Error("must not resolve hosted");}},true,(async(_url,init)=>{assert.equal((init!.headers as Record<string,string>).Authorization,"Bearer TENANT_KEY");return Response.json({choices:[{finish_reason:"stop",message:{content:"synthetic"}}]});}) as typeof fetch);
  await byok.run({...input,keyMode:"byok",credentialId:id()},{signal:new AbortController().signal,attemptId:id(),credential:"TENANT_KEY",quote});
});
