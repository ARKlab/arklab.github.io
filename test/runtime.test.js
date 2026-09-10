import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {build} from 'esbuild';

const compiled=await build({entryPoints:['src/runtime.js'],bundle:true,write:false,format:'iife',define:{__SITE_URL__:JSON.stringify('https://example.com/'),__APP_VERSION__:JSON.stringify('test')},plugins:[{
  name:'provider-fixtures',setup(builder){
    builder.onResolve({filter:/^(@azure\/msal-browser|\.\/network\.js)$/},args=>({path:args.path,namespace:'fixture'}));
    builder.onLoad({filter:/.*/,namespace:'fixture'},args=>({contents:args.path==='@azure/msal-browser'
      ? 'export const createNestablePublicClientApplication=fixtures.create; export class InteractionRequiredAuthError extends Error {}'
      : 'export const fetchBundle=async()=>fixtures.bundle; export const jsonRequest=fixtures.graph;'}));
  }
}]});
const bundle={enabled:true,revision:'test',deployment:{clientId:'00000000-0000-0000-0000-000000000001',tenantId:'00000000-0000-0000-0000-000000000002'},branding:JSON.parse(await readFile('branding.json','utf8')),templates:{full:await readFile('templates/full.html','utf8')},assets:{ark:{filename:'ark.png',base64:'eA=='},artesian:{filename:'artesian.png',base64:'eA=='}}};

function runtime({persistent=false}={}) {
  const calls={requests:[],sets:[],warnings:[],details:[],clears:0,completed:0,graph:0,popups:0};
  const handlers={};
  const sender={emailAddress:'alex@ark-energy.eu',displayName:'Alex Example'};
  const ok=(callback,value)=>callback({status:'succeeded',value});
  const item={from:{getAsync:callback=>ok(callback,sender)},getComposeTypeAsync:callback=>ok(callback,{composeType:'newMail'}),getAttachmentsAsync:callback=>ok(callback,[]),addFileAttachmentFromBase64Async:(_bytes,_name,_options,callback)=>ok(callback,'attachment'),body:{getTypeAsync:callback=>ok(callback,'html'),setSignatureAsync:(html,_options,callback)=>{calls.sets.push(html);ok(callback);}},notificationMessages:{replaceAsync:(_key,value)=>calls.warnings.push(value.message),removeAsync:()=>calls.clears++}};
  runInNewContext(compiled.outputFiles[0].text,{URL,clearTimeout,setTimeout:(fn,ms)=>setTimeout(fn,ms===1000?0:ms),console:{warn:text=>calls.details.push(text)},Office:{onReady(){},actions:{associate:(name,handler)=>{handlers[name]=handler;}},context:{requirements:{isSetSupported:()=>true},mailbox:{item,userProfile:sender}}},fixtures:{bundle,
    create:async()=>({acquireTokenSilent:async request=>{calls.requests.push(request);if(persistent||calls.requests.length===1)throw Object.assign(new Error('PRIVATE_PROVIDER_PAYLOAD'),{errorCode:'7000024'});return {accessToken:'SYNTHETIC_TOKEN',account:{tenantId:bundle.deployment.tenantId}};},acquireTokenPopup:()=>{calls.popups++;assert.fail('No background popup');}}),
    graph:async()=>{calls.graph++;return {displayName:sender.displayName,mail:sender.emailAddress,jobTitle:'Director'};}
  }});
  return {calls,run:()=>handlers.arkOnCompose({completed:()=>calls.completed++})};
}

test('automatic compose silently recovers, inserts once and clears the old warning',async()=>{
  const {run,calls}=runtime();
  await run();
  assert.equal(calls.requests.length,2);
  assert.equal(calls.requests[1].forceRefresh,true);
  assert.equal(calls.sets.length,1);
  assert.ok(calls.sets[0].includes('Alex Example'));
  assert.equal(calls.warnings.length,0);
  assert.equal(calls.clears,1);
  assert.equal(calls.completed,1);
  assert.equal(calls.popups,0);
});

test('repeated broker failure preserves the signature and logs only safe support references',async()=>{
  const {run,calls}=runtime({persistent:true});
  await run();
  assert.equal(calls.requests.length,2);
  assert.equal(calls.sets.length,0);
  assert.equal(calls.graph,0);
  assert.equal(calls.warnings.length,1);
  assert.ok(calls.details[0].includes('AADSTS7000024'));
  assert.ok(calls.details[0].includes('Token attempts: 2'));
  assert.equal(calls.details[0].includes('PRIVATE_PROVIDER_PAYLOAD'),false);
  assert.equal(calls.details[0].includes('SYNTHETIC_TOKEN'),false);
  assert.equal(calls.completed,1);
});
