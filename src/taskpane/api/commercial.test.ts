import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { CommercialClient, C2_AVAILABLE } from "./commercial";
import { toolsForApi } from "../../excel/toolSchemas";
import { runAgent } from "../../agent/loop";

const org = "00000000-0000-4000-8000-000000000001";
const task = "00000000-0000-4000-8000-000000000002";
const initialContext = { workbook:{sessionId:"synthetic",documentUrl:"PRIVATE_WORKBOOK_PATH",identityConfirmed:true}, activeSheet:{name:"Sheet1",id:"sheet-1"}, selection:{address:"Sheet1!A1"}, readAt:"2026-11-09T00:00:00Z" } as any;

test("commercial mode defaults off; generated server catalog matches existing host tool schemas", () => {
  assert.equal(C2_AVAILABLE,false);
  const prev=(globalThis as any).Office;
  try {
    (globalThis as any).Office={context:{requirements:{isSetSupported:()=>true}}};
    const expected=[...toolsForApi(false,false,false),...toolsForApi(false,false,false,new Set()).filter(t=>t.function.name==="enable_tools")];
    const actual=JSON.parse(readFileSync(new URL("../../../server/src/commercial/toolCatalog.json",import.meta.url),"utf8"));
    assert.deepEqual(actual,expected);
    assert.ok(actual.some((t:any)=>t.function.name==="reconcile_ranges"));
  } finally { (globalThis as any).Office=prev; }
});

test("commercial agent transport sends tools/results under one task, reuses terminal SSE parsing, never reads BYOK storage",async(t)=>{
  const previousStorage=(globalThis as any).localStorage;
  (globalThis as any).localStorage={getItem:()=>{throw new Error("BYOK_STORAGE_MUST_NOT_BE_READ");}};
  t.after(()=>{(globalThis as any).localStorage=previousStorage;});
  const requests:any[]=[];
  const client=new CommercialClient((async(url,init)=>{
    if(String(url).endsWith("/auth/redeem"))return Response.json({token:"memory-session"});
    if(String(url).endsWith("/me"))return Response.json({organizations:[{id:org,name:"synthetic",role:"owner"}]});
    requests.push({url,init,body:JSON.parse(init!.body as string)});
    const delta=requests.length===1?{tool_calls:[{index:0,id:"call-1",function:{name:"list_sheets",arguments:"{}"}}]}:{content:"synthetic complete"};
    return new Response(`data: ${JSON.stringify({choices:[{delta,finish_reason:requests.length===1?"tool_calls":"stop"}]})}\n\ndata: [DONE]\n\n`);
  }) as typeof fetch);
  await client.redeem("ticket","proof");
  const chat=client.chatForTask(task);
  const first=await chat({provider:"fake",model:"c2-pilot",messages:[{role:"user",content:"synthetic"}],tools:[{type:"function",function:{name:"list_sheets",parameters:{}}}],onDelta:()=>undefined});
  assert.equal(first.toolCalls[0].name,"list_sheets");
  await chat({provider:"fake",model:"c2-pilot",messages:[{role:"assistant",content:"",tool_calls:first.toolCalls},{role:"tool",tool_call_id:"call-1",content:'{"ok":true,"workbook":{"documentUrl":"PRIVATE_WORKBOOK_PATH"}}'}],tools:[],onDelta:()=>undefined});
  assert.equal(requests.length,2);
  assert.ok(requests.every(r=>r.body.taskId===task&&r.body.keyMode==="hosted"&&r.init.credentials==="omit"));
  assert.notEqual(requests[0].init.headers["Idempotency-Key"],requests[1].init.headers["Idempotency-Key"]);
  assert.deepEqual(requests[0].body.toolNames,["list_sheets"]);
  assert.equal(requests[0].body.tools,undefined);
  assert.equal(requests[1].body.messages[1].role,"tool");
  assert.doesNotMatch(JSON.stringify(requests),/PRIVATE_WORKBOOK_PATH/);
  assert.throws(()=>client.selectOrganization("foreign"));
});

test("replay metadata and truncated SSE cannot become executable tool calls",async()=>{
  for(const wire of [
    'data: {"error":{"message":"State only, no replay"}}\n\ndata: [DONE]\n\n',
    'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"write","function":{"name":"set_range_values","arguments":"{}"}}]}}]}\n\n'
  ]){
    const client=new CommercialClient((async(url)=>String(url).endsWith("/auth/redeem")?Response.json({token:"memory"}):String(url).endsWith("/me")?Response.json({organizations:[{id:org,name:"fixture",role:"owner"}]}):new Response(wire)) as typeof fetch);
    await client.redeem("ticket","proof");
    await assert.rejects(client.chatForTask(task)({provider:"fake",model:"c2-pilot",messages:[{role:"user",content:"fixture"}],tools:[],onDelta:()=>undefined}));
  }
});

test("commercial agent disables automatic network retries, retaining the existing executor",async()=>{
  let calls=0,tools=0;
  await assert.rejects(runAgent({provider:"fake",model:"c2-pilot",history:[{role:"user",content:"synthetic"}],initialContext,
    disableNetworkRetries:true,localServer:false,omitWorkbookLocation:true,chat:async opts=>{calls++;assert.doesNotMatch(JSON.stringify(opts.messages),/PRIVATE_WORKBOOK_PATH/);throw new Error("Failed to fetch");},
    hooks:{onDelta:()=>undefined,onStepEnd:()=>undefined,onToolEvent:()=>{tools++;},confirm:async()=>{throw new Error("must not confirm");}}}));
  assert.equal(calls,1);assert.equal(tools,0);
});

test("logout clears the in-memory session even during outage; a new panel instance needs sign-in",async()=>{
  const client=new CommercialClient((async(url)=>String(url).endsWith("/auth/redeem")?Response.json({token:"memory"}):String(url).endsWith("/me")?Response.json({organizations:[{id:org,name:"fixture",role:"owner"}]}):new Response("offline",{status:503})) as typeof fetch);
  await client.redeem("ticket","proof");assert.equal(client.signedIn,true);
  await assert.rejects(client.logout());assert.equal(client.signedIn,false);
  assert.equal(new CommercialClient().signedIn,false);
  await assert.rejects(client.createTask(),/Microsoft/);
  const expired=new CommercialClient((async(url)=>String(url).endsWith("/auth/redeem")?Response.json({token:"memory"}):String(url).endsWith("/me")?Response.json({organizations:[{id:org,name:"fixture",role:"owner"}]}):new Response("expired",{status:401})) as typeof fetch);
  await expired.redeem("ticket","proof");await assert.rejects(expired.budgets());assert.equal(expired.signedIn,false);
});

test("Office dialog rejects a message from another origin before ticket redemption",async(t)=>{
  const prevOffice=(globalThis as any).Office,prevLocation=Object.getOwnPropertyDescriptor(globalThis,"location");
  Object.defineProperty(globalThis,"location",{configurable:true,value:{origin:"https://localhost:3100"}});
  let redeemed=0,closed=0;
  (globalThis as any).Office={AsyncResultStatus:{Succeeded:"ok"},EventType:{DialogEventReceived:"event",DialogMessageReceived:"message"},context:{ui:{
    displayDialogAsync:(_url:any,_opts:any,callback:any)=>callback({status:"ok",value:{close:()=>{closed++;},addEventHandler:(type:string,handler:any)=>{
      if(type==="message")queueMicrotask(()=>handler({origin:"https://attacker.invalid",message:JSON.stringify({type:"amai-auth",ticket:"a".repeat(43)})}));
    }}})
  }}};
  t.after(()=>{(globalThis as any).Office=prevOffice;if(prevLocation)Object.defineProperty(globalThis,"location",prevLocation);else delete (globalThis as any).location;});
  const client=new CommercialClient((async(url)=>{
    if(String(url).endsWith("/auth/start"))return Response.json({dialogUrl:"https://localhost:3100/api/v1/auth/dialog?state=fixture"});
    redeemed++;throw new Error("must not redeem");
  }) as typeof fetch);
  await assert.rejects(client.login());assert.equal(redeemed,0);assert.equal(closed,1);assert.equal(client.signedIn,false);
});
