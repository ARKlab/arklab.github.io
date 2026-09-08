import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {profileForSender,renderSignature,plainSignature} from '../src/render.js';
import {applySignature,completeEvent} from '../src/office-flow.js';

const branding=JSON.parse(await readFile(new URL('../branding.json',import.meta.url)));
const bundle={enabled:true,branding,revision:'test123',assets:{ark:{filename:'ark-test.png',base64:'eA=='},artesian:{filename:'artesian-test.png',base64:'eA=='}},templates:{full:await readFile(new URL('../templates/full.html',import.meta.url),'utf8'),reply:await readFile(new URL('../templates/reply.html',import.meta.url),'utf8')}};
const graph={displayName:'Alex Example',mail:'alex@ark-energy.eu',userPrincipalName:'alex@ark-energy.eu',jobTitle:'Director',businessPhones:['+353 83 111 2222'],mobilePhone:'+39 333 111 2222',officeLocation:'Dublin'};
const sender={displayName:'Alex Example',emailAddress:'alex@ark-energy.eu'};

function fakeItem({composeType='newMail',bodyType='html',existing=[]}={}) {
  const state={sets:[],attachments:[],from:sender,fromReads:0};
  const ok=(cb,value)=>cb({status:'succeeded',value});
  const item={
    from:{getAsync:cb=>{state.fromReads++;ok(cb,state.from);}},
    getComposeTypeAsync:cb=>ok(cb,{composeType}),
    getAttachmentsAsync:cb=>ok(cb,existing),
    addFileAttachmentFromBase64Async:(bytes,name,options,cb)=>{state.attachments.push({bytes,name,options});ok(cb,name);},
    body:{getTypeAsync:cb=>ok(cb,bodyType),setSignatureAsync:(html,options,cb)=>{state.sets.push({html,options});ok(cb);}}
  };
  return {item,state};
}
test('directory fields are escaped and public mobile numbers are opt-in',()=>{
  const p=profileForSender({...graph,displayName:'Alex <img src=x onerror=alert(1)>',jobTitle:'Research & Development'},sender,branding);
  const html=renderSignature(bundle,p);
  assert.ok(html.includes('&lt;img'));
  assert.ok(!html.includes('<img src=x'));
  assert.ok(html.includes('Research &amp; Development'));
  assert.ok(!html.includes('333&#160;111'));
});
test('missing directory fields produce no placeholder or empty contact rows',()=>{
  const p=profileForSender({mail:sender.emailAddress,displayName:'Alex'},sender,branding);
  const html=renderSignature(bundle,p);
  assert.ok(!html.includes('{{'));
  assert.ok(!html.includes('undefined'));
  assert.ok(!html.includes('tel:'));
  assert.equal(p.title,'');assert.equal(p.office,'');assert.deepEqual(p.phones,[]);
});
test('changing From never leaks another employee’s title or phone',()=>{
  const p=profileForSender(graph,{displayName:'Service team',emailAddress:'service@ark-energy.eu'},branding);
  assert.equal(p.directoryMatched,false);assert.equal(p.name,'Service team');assert.equal(p.title,'');assert.deepEqual(p.phones,[]);
  assert.throws(()=>profileForSender(graph,{emailAddress:'alex@example.org'},branding),/UNAPPROVED/);
});
test('mailbox display names cannot smuggle mailto headers',()=>{
  const p=profileForSender({...graph,mail:'alex?subject=oops@ark-energy.eu'},{emailAddress:'alex?subject=oops@ark-energy.eu'},branding);
  assert.ok(renderSignature(bundle,p).includes('mailto:alex%3Fsubject%3Doops@ark-energy.eu'));
});
test('signature insertion only writes the signature slot with embedded PNGs',async()=>{
  const {item,state}=fakeItem();
  const result=await applySignature({item,bundle,getGraph:async()=>graph});
  assert.equal(result.status,'applied');assert.equal(state.sets.length,1);assert.equal(state.attachments.length,2);
  assert.ok(state.attachments.every(a=>a.options.isInline));
  assert.ok(state.sets[0].html.includes('cid:ark-test.png'));
  assert.ok(state.sets[0].html.includes('Home of'));
  assert.ok(!state.sets[0].html.includes(' style='));
  assert.ok(state.sets[0].html.startsWith('<style'));
});
test('repeated insertion reuses existing managed inline images',async()=>{
  const {item,state}=fakeItem({existing:[{name:'ark-test.png',isInline:true},{name:'artesian-test.png',isInline:true}]});
  await applySignature({item,bundle,getGraph:async()=>graph});
  await applySignature({item,bundle,getGraph:async()=>graph});
  assert.equal(state.attachments.length,0);assert.equal(state.sets.length,2);
});
test('replies and forwards use the compact co-branded signature',async()=>{
  for(const composeType of ['reply','forward']) {
    const {item,state}=fakeItem({composeType});await applySignature({item,bundle,getGraph:async()=>graph});
    assert.equal(state.attachments.length,0);assert.ok(state.sets[0].html.includes('Home of Artesian'));
  }
});
test('plain text email stays plain text and needs no attachments',async()=>{
  const {item,state}=fakeItem({bodyType:'text'});await applySignature({item,bundle,getGraph:async()=>graph});
  assert.equal(state.attachments.length,0);assert.equal(state.sets[0].options.coercionType,'text');
  assert.ok(!state.sets[0].html.includes('<table'));assert.ok(state.sets[0].html.includes('Home of Artesian'));
});
test('paused deployment and directory failures preserve the existing signature',async()=>{
  const {item,state}=fakeItem();let read=false;
  await applySignature({item,bundle:{...bundle,enabled:false},getGraph:async()=>{read=true;return graph;}});
  assert.equal(read,false);assert.equal(state.sets.length,0);
  await assert.rejects(applySignature({item,bundle,getGraph:async()=>{throw new Error('AUTH');}}));
  assert.equal(state.sets.length,0);assert.equal(state.attachments.length,0);
});
test('sender switch during profile loading cancels an outdated insertion',async()=>{
  const {item,state}=fakeItem();
  await assert.rejects(applySignature({item,bundle,getGraph:async()=>{state.from={emailAddress:'other@ark-energy.eu'};return graph;}}),/SENDER_CHANGED/);
  assert.equal(state.sets.length,0);
});
test('event completion happens once, after successful insertion',async()=>{
  const {item,state}=fakeItem();let completed=0;
  await completeEvent({completed:()=>{assert.equal(state.sets.length,1);completed++;}},context=>applySignature({item,bundle,getGraph:async()=>graph,context}));
  assert.equal(completed,1);
});
test('a timed-out directory request cannot insert a signature later',async()=>{
  const {item,state}=fakeItem();let completed=0;let errors=0;
  await completeEvent({completed:()=>completed++},context=>applySignature({item,bundle,context,getGraph:()=>new Promise(resolve=>setTimeout(()=>resolve(graph),30))}),{timeoutMs:5,onError:()=>errors++});
  assert.equal(completed,1);assert.equal(errors,1);assert.equal(state.sets.length,0);assert.equal(state.attachments.length,0);
});
test('new branding revisions change both full and compact output without employee editing',()=>{
  const p=profileForSender(graph,sender,branding);
  const revised={...bundle,revision:'test456',branding:{...branding,descriptor:'Revised descriptor',relationship:'Creators of'}};
  assert.ok(renderSignature(revised,p).includes('Revised descriptor'));
  assert.ok(plainSignature(revised,p,true).includes('Creators of Artesian'));
});
test('unconfirmed Artesian destination is omitted, and invalid links fail closed',()=>{
  const p=profileForSender(graph,sender,branding);
  assert.ok(!renderSignature(bundle,p).includes('href=""'));
  assert.throws(()=>renderSignature({...bundle,branding:{...branding,artesianWebsite:'javascript:alert(1)'}},p));
});
