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
const bundle={enabled:true,revision:'test',deployment:{clientId:'00000000-0000-0000-0000-000000000001',tenantId:'00000000-0000-0000-0000-000000000002'},branding:JSON.parse(await readFile('branding.json','utf8')),templates:{full:await readFile('templates/full.html','utf8'),reply:await readFile('templates/reply.html','utf8')},assets:{ark:{filename:'ark.png',base64:'eA=='},artesian:{filename:'artesian.png',base64:'eA=='}}};

function runtime({persistent=false,senderAddress='alex@ark-energy.eu',senderName='Alex Example',composeType='newMail',proxyAddresses=[],platform='Mac',sessionFailure,signatureFailure}={}) {
  const calls={requests:[],sets:[],warnings:[],details:[],clears:0,completed:0,graph:0,popups:0,attachments:0};
  const handlers={};
  const mailboxProfile={emailAddress:'alex@ark-energy.eu',displayName:'Alex Example'};
  const sender={emailAddress:senderAddress,displayName:senderName};
  const ok=(callback,value)=>callback({status:'succeeded',value});
  const item={from:{getAsync:callback=>ok(callback,sender)},getComposeTypeAsync:callback=>ok(callback,{composeType}),getAttachmentsAsync:callback=>ok(callback,[]),addFileAttachmentFromBase64Async:(_bytes,_name,_options,callback)=>ok(callback,'attachment'),body:{getTypeAsync:callback=>ok(callback,'html'),setSignatureAsync:(html,_options,callback)=>{calls.sets.push(html);ok(callback);}},notificationMessages:{replaceAsync:(_key,value)=>calls.warnings.push(value.message),removeAsync:()=>calls.clears++}};
  if (['Android','iOS'].includes(platform)) {
    const session=new Map();
    item.getAttachmentsAsync=()=>assert.fail('Unsupported mobile API');item.body.getTypeAsync=()=>assert.fail('Unsupported mobile API');
    // null models a failed callback with no error code; undefined keeps normal reads.
    item.sessionData={getAsync:(key,cb)=>sessionFailure!==undefined?cb({status:'failed',error:sessionFailure===null?{}:{code:sessionFailure}}):session.has(key)?ok(cb,session.get(key)):cb({status:'failed',error:{code:9050}}),setAsync:(key,value,cb)=>{session.set(key,value);ok(cb);}};
    item.addFileAttachmentFromBase64Async=(_bytes,_name,_options,cb)=>{calls.attachments++;ok(cb,'attachment');};
  }
  if(signatureFailure) item.body.setSignatureAsync=(_html,_options,cb)=>cb({status:'failed',error:{code:signatureFailure}});
  runInNewContext(compiled.outputFiles[0].text,{URL,clearTimeout,setTimeout:(fn,ms)=>setTimeout(fn,ms===1000?0:ms),console:{warn:text=>calls.details.push(text)},Office:{onReady(){},actions:{associate:(name,handler)=>{handlers[name]=handler;}},context:{platform,requirements:{isSetSupported:()=>true},mailbox:{item,userProfile:mailboxProfile}}},fixtures:{bundle,
    create:async()=>({acquireTokenSilent:async request=>{calls.requests.push(request);if(persistent||calls.requests.length===1)throw Object.assign(new Error('PRIVATE_PROVIDER_PAYLOAD'),{errorCode:'7000024'});return {accessToken:'SYNTHETIC_TOKEN',account:{tenantId:bundle.deployment.tenantId}};},acquireTokenPopup:()=>{calls.popups++;assert.fail('No background popup');}}),
    graph:async()=>{calls.graph++;return {displayName:mailboxProfile.displayName,mail:mailboxProfile.emailAddress,jobTitle:'Director',businessPhones:['+353 83 111 2222'],proxyAddresses};}
  }});
  return {calls,sender,run:(name='arkOnCompose')=>handlers[name]({completed:()=>calls.completed++})};
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
test('mobile event runtime signs in silently and reuses logos on a subsequent From event',async()=>{
  for (const platform of ['Android','iOS']) {
    const {run,calls}=runtime({platform,senderAddress:'alex@ark-energy.it',proxyAddresses:['smtp:alex@ark-energy.it']});
    await run();await run('arkOnFromChanged');
    assert.equal(calls.sets.length,2);assert.equal(calls.attachments,2);assert.equal(calls.completed,2);
    assert.equal(calls.warnings.length,0);assert.equal(calls.popups,0);
    assert.ok(calls.sets.every(html=>html.includes('mailto:alex@ark-energy.it')&&html.includes('Director')));
    assert.ok(calls.requests.every(r=>r.loginHint==='alex@ark-energy.eu'&&r.scopes.join(',')==='User.Read'));
  }
});
test('mobile authentication failure directs recovery to a received message',async()=>{
  const {run,calls}=runtime({platform:'Android',persistent:true});await run();
  assert.equal(calls.sets.length,0);assert.equal(calls.attachments,0);assert.equal(calls.completed,1);
  assert.ok(calls.warnings[0].includes('received email'));assert.equal(calls.popups,0);
});
test('mobile Office failures expose a short numeric support reference without provider payloads',async()=>{
  for(const [options,reference] of [
    [{sessionFailure:9051},'session-read/OUTLOOK_9051'],
    [{signatureFailure:5001},'signature-write/OUTLOOK_5001'],
    [{signatureFailure:9999999999},'signature-write/OUTLOOK_9999999999'],
    [{sessionFailure:null},'session-read'],
    [{sessionFailure:'SECRET_EMPLOYEE_ADDRESS'},'session-read']
  ]) {
    const {run,calls}=runtime({platform:'Android',...options});await run();
    assert.equal(calls.sets.length,0);assert.equal(calls.completed,1);assert.equal(calls.warnings.length,1);
    assert.ok(calls.warnings[0].includes('('+reference+')'));
    assert.ok(calls.warnings[0].includes('share this reference with IT'));
    if(options.sessionFailure===null) {
      assert.equal(calls.warnings[0].includes('OUTLOOK_FAILED'),false);
      assert.ok(calls.details[0].includes('Result: OUTLOOK_FAILED'));
    }
    assert.ok(calls.warnings[0].length<=150);assert.ok(calls.details[0].includes('Step: '));
    assert.equal((calls.warnings.join()+calls.details.join()).includes('SECRET_'),false);
  }
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
test('From-change insertion uses the verified .it alias while signing in as the mailbox owner',async()=>{
  const {run,calls}=runtime({senderAddress:'alex@ark-energy.it',proxyAddresses:['SMTP:alex@ark-energy.eu','smtp:alex@ark-energy.it']});
  await run('arkOnFromChanged');
  assert.equal(calls.sets.length,1);assert.equal(calls.warnings.length,0);
  assert.ok(calls.sets[0].includes('mailto:alex@ark-energy.it'));
  assert.ok(calls.sets[0].includes('Director'));assert.ok(calls.sets[0].includes('+353&#160;83&#160;111&#160;2222'));
  assert.ok(calls.requests.every(request=>request.loginHint==='alex@ark-energy.eu'&&request.scopes.join(',')==='User.Read'));
  assert.equal(calls.popups,0);assert.equal(calls.completed,1);
});

test('desktop events render a shared From name and address without borrowing the delegate profile',async()=>{
  for(const platform of ['Mac','PC','OfficeOnline']) for(const composeType of ['newMail','reply','forward']) {
    const {run,calls}=runtime({platform,composeType,senderAddress:'service@ark-energy.eu',senderName:'Service team'});
    await run();
    assert.equal(calls.sets.length,1);assert.equal(calls.completed,1);
    const html=calls.sets[0];
    assert.ok(html.includes('Service team'));assert.ok(html.includes('mailto:service@ark-energy.eu'));assert.ok(html.includes('Home of'));
    assert.ok(!html.includes('Alex Example'));assert.ok(!html.includes('Director'));assert.ok(!html.includes('tel:'));
    assert.equal(html.includes('cid:ark.png'),composeType==='newMail');
    assert.ok(calls.requests.every(request=>request.loginHint==='alex@ark-energy.eu'&&request.scopes.join(',')==='User.Read'));
    assert.equal(calls.popups,0);
  }
});

test('switching between shared and personal From identities restores only the matching contact details',async()=>{
  const {run,calls,sender}=runtime({senderAddress:'service@ark-energy.eu',senderName:'Service team'});
  await run();
  Object.assign(sender,{emailAddress:'alex@ark-energy.eu',displayName:'Alex Example'});
  await run('arkOnFromChanged');
  Object.assign(sender,{emailAddress:'service@ark-energy.eu',displayName:'Service team'});
  await run('arkOnFromChanged');
  assert.equal(calls.sets.length,3);assert.equal(calls.completed,3);
  assert.ok(calls.sets[1].includes('Alex Example'));assert.ok(calls.sets[1].includes('Director'));assert.ok(calls.sets[1].includes('tel:'));
  for(const html of [calls.sets[0],calls.sets[2]]) {
    assert.ok(html.includes('Service team'));assert.ok(html.includes('service@ark-energy.eu'));
    assert.ok(!html.includes('Alex Example'));assert.ok(!html.includes('Director'));assert.ok(!html.includes('tel:'));
  }
});
