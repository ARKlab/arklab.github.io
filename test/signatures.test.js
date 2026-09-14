import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {profileForSender,renderSignature,plainSignature,formatPhoneNumber} from '../src/render.js';
import {applySignature,completeEvent} from '../src/office-flow.js';

// Stable fixtures keep routine company configuration edits independent of test expectations.
const branding={schemaVersion:1,descriptor:'Energy markets. Managed data services. Technology.',relationship:'Home of',arkWebsite:'https://www.ark-energy.eu/en',arkWebsiteLabel:'ark-energy.eu',artesianWebsite:'https://www.artesian.cloud/',artesianWebsiteLabel:'artesian.cloud',artesianWordmark:true,showOfficeLocation:true,includeMobilePhone:false,maxPhoneNumbers:2,compactReplies:true,approvedSenderDomains:['ark-energy.eu','ark-energy.it','artesian.cloud']};
const bundle={enabled:true,branding,revision:'test123',assets:{ark:{filename:'ark-test.png',base64:'eA=='},artesian:{filename:'artesian-test.png',base64:'eA=='}},templates:{full:await readFile(new URL('../templates/full.html',import.meta.url),'utf8'),reply:await readFile(new URL('../templates/reply.html',import.meta.url),'utf8')}};
const graph={displayName:'Alex Example',mail:'alex@ark-energy.eu',userPrincipalName:'alex@ark-energy.eu',jobTitle:'Director',businessPhones:['+353 83 111 2222'],mobilePhone:'+39 333 111 2222',officeLocation:'Dublin'};
const sender={displayName:'Alex Example',emailAddress:'alex@ark-energy.eu'};

function fakeItem({composeType='newMail',bodyType='html',existing=[],mobile=false}={}) {
  const state={sets:[],attachments:[],from:sender,fromReads:0};
  const ok=(cb,value)=>cb({status:'succeeded',value});
  const item={
    from:{getAsync:cb=>{state.fromReads++;ok(cb,state.from);}},
    getComposeTypeAsync:cb=>ok(cb,{composeType}),
    getAttachmentsAsync:cb=>ok(cb,existing),
    addFileAttachmentFromBase64Async:(bytes,name,options,cb)=>{state.attachments.push({bytes,name,options});ok(cb,name);},
    body:{getTypeAsync:cb=>ok(cb,bodyType),setSignatureAsync:(html,options,cb)=>{state.sets.push({html,options});ok(cb);}}
  };
  if (mobile) {
    state.session=new Map();
    item.getAttachmentsAsync=()=>assert.fail('Mobile must not enumerate attachments');
    item.body.getTypeAsync=()=>assert.fail('Mobile must not request the body format');
    item.sessionData={getAsync:(key,cb)=>ok(cb,state.session.get(key)),setAsync:(key,value,cb)=>{state.session.set(key,value);ok(cb);}};
  }
  return {item,state};
}

test('mobile full signatures reuse inline logos across repeated insertion and an alias switch',async()=>{
  const {item,state}=fakeItem({mobile:true});
  const getGraph=async()=>({...graph,proxyAddresses:['smtp:alex@ark-energy.it']});
  await applySignature({item,bundle,getGraph,mobile:true});
  state.from={...sender,emailAddress:'alex@ark-energy.it'};
  await applySignature({item,bundle,getGraph,mobile:true});
  assert.equal(state.attachments.length,2);assert.equal(state.sets.length,2);
  assert.ok(state.attachments.every(a=>a.options.isInline));
  assert.ok(state.sets.every(s=>s.options.coercionType==='html'&&s.html.includes('Director')&&s.html.includes('cid:ark-test.png')));
  assert.ok(state.sets[1].html.includes('mailto:alex@ark-energy.it'));
  assert.ok(!state.sets[1].html.includes('mailto:alex@ark-energy.eu'));
  assert.deepEqual([...state.session.values()],['added','added']);
  assert.ok([...state.session.keys()].every(key=>!key.includes('@')));
});
test('mobile replies and forwards use compact HTML without attachment or session APIs',async()=>{
  for (const composeType of ['reply','forward']) {
    const {item,state}=fakeItem({mobile:true,composeType});delete item.sessionData;
    await applySignature({item,bundle,getGraph:async()=>graph,mobile:true});
    assert.equal(state.attachments.length,0);
    assert.ok(!state.sets[0].html.includes('<img'));
    assert.ok(state.sets[0].html.includes('Home of'));
  }
});
test('mobile session markers are isolated to each draft',async()=>{
  for (let index=0;index<2;index++) {
    const {item,state}=fakeItem({mobile:true});
    await applySignature({item,bundle,getGraph:async()=>graph,mobile:true});
    assert.equal(state.attachments.length,2);
  }
});
test('mobile retry after partial attachment failure reuses the successful logo',async()=>{
  const {item,state}=fakeItem({mobile:true});const attach=item.addFileAttachmentFromBase64Async;
  item.addFileAttachmentFromBase64Async=(bytes,name,options,cb)=>name==='artesian-test.png'?cb({status:'failed',error:{code:'UPLOAD'}}):attach(bytes,name,options,cb);
  await assert.rejects(applySignature({item,bundle,getGraph:async()=>graph,mobile:true}),/OUTLOOK_UPLOAD/);
  assert.equal(state.sets.length,0);assert.equal(state.attachments.length,1);
  item.addFileAttachmentFromBase64Async=attach;
  await applySignature({item,bundle,getGraph:async()=>graph,mobile:true});
  assert.equal(state.attachments.length,2);assert.equal(state.sets.length,1);
});
test('missing or failed mobile session APIs preserve the current signature',async()=>{
  for (const unavailable of [true,false]) {
    const {item,state}=fakeItem({mobile:true});
    if (unavailable) delete item.sessionData;
    else item.sessionData.getAsync=(_key,cb)=>cb({status:'failed',error:{code:'SESSION'}});
    await assert.rejects(applySignature({item,bundle,getGraph:async()=>graph,mobile:true}),unavailable?/OUTLOOK_UPDATE_REQUIRED/:/OUTLOOK_SESSION/);
    assert.equal(state.sets.length,0);assert.equal(state.attachments.length,0);
  }
});
test('cancellation during mobile session retrieval cannot add logos or insert later',async()=>{
  const {item,state}=fakeItem({mobile:true});const context={cancelled:false};
  item.sessionData.getAsync=(_key,cb)=>{context.cancelled=true;cb({status:'succeeded'});};
  await assert.rejects(applySignature({item,bundle,getGraph:async()=>graph,mobile:true,context}),/CANCELLED/);
  assert.equal(state.sets.length,0);assert.equal(state.attachments.length,0);
});
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
test('a directory SMTP alias retains the owner profile and selected email in every signature form',()=>{
  const alias='alex.italia@ark-energy.it';
  const p=profileForSender({...graph,proxyAddresses:['SMTP:alex@ark-energy.eu','smtp:'+alias]},{emailAddress:alias},branding);
  assert.equal(p.directoryMatched,true);assert.equal(p.name,graph.displayName);assert.equal(p.title,graph.jobTitle);
  assert.deepEqual(p.phones,graph.businessPhones);assert.equal(p.email,alias);
  for(const compact of [false,true]) {
    const html=renderSignature(bundle,p,{compact});
    assert.ok(html.includes('mailto:'+alias));assert.ok(html.includes('Director'));
    assert.ok(html.includes('+353&#160;83&#160;111&#160;2222'));
    assert.ok(!html.includes('mailto:alex@ark-energy.eu'));
    assert.ok(plainSignature(bundle,p,compact).includes(alias));
  }
});
test('alias matching normalises SMTP case and ignores non-SMTP or unverified addresses',()=>{
  const alias='alex@ark-energy.it';
  const verified=profileForSender({...graph,proxyAddresses:[null,42,'SMTP: ALEX@ARK-ENERGY.IT ']},{emailAddress:alias},branding);
  assert.equal(verified.directoryMatched,true);
  for(const extras of [{},{proxyAddresses:null},{proxyAddresses:'smtp:'+alias},{proxyAddresses:['SIP:'+alias,'X500:'+alias,alias,'smtp:other@ark-energy.it']},{otherMails:[alias]}]) {
    const p=profileForSender({...graph,...extras},{emailAddress:alias,displayName:'Other sender'},branding);
    assert.equal(p.directoryMatched,false);assert.equal(p.title,'');assert.deepEqual(p.phones,[]);
  }
  assert.throws(()=>profileForSender({...graph,proxyAddresses:['smtp:alex@example.org']},{emailAddress:'alex@example.org'},branding),/UNAPPROVED/);
});
test('mailbox display names cannot smuggle mailto headers',()=>{
  const p=profileForSender({...graph,mail:'alex?subject=oops@ark-energy.eu'},{emailAddress:'alex?subject=oops@ark-energy.eu'},branding);
  assert.ok(renderSignature(bundle,p).includes('mailto:alex%3Fsubject%3Doops@ark-energy.eu'));
});
test('mailto links preserve literal percent sequences and encode URL delimiters',()=>{
  const email='alex%40team?subject=x&bcc=y#note@ark-energy.eu';
  const p=profileForSender({...graph,mail:email},{emailAddress:email},branding);
  const html=renderSignature(bundle,p);
  assert.ok(html.includes('href="mailto:alex%2540team%3Fsubject%3Dx%26bcc%3Dy%23note@ark-energy.eu"'));
  assert.throws(()=>renderSignature(bundle,{...p,email:'alex@team@ark-energy.eu'}),/INVALID_SENDER/);
});
test('signature insertion only writes the signature slot with embedded PNGs',async()=>{
  const {item,state}=fakeItem();
  const result=await applySignature({item,bundle,getGraph:async()=>graph});
  assert.equal(result.status,'applied');assert.equal(state.sets.length,1);assert.equal(state.attachments.length,2);
  assert.ok(state.attachments.every(a=>a.options.isInline));
  assert.ok(state.sets[0].html.includes('cid:ark-test.png'));
  assert.ok(state.sets[0].html.includes('Home of'));
  assert.ok(state.sets[0].html.includes(' style='));
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
    assert.equal(state.attachments.length,0);assert.ok(state.sets[0].html.includes('Home of'));
    assert.ok(state.sets[0].html.includes('href="https://www.artesian.cloud/"'));
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
test('switching between verified aliases still cancels insertion for the previous From address',async()=>{
  const {item,state}=fakeItem();
  await assert.rejects(applySignature({item,bundle,getGraph:async()=>{
    state.from={emailAddress:'alex@ark-energy.it'};
    return {...graph,proxyAddresses:['SMTP:alex@ark-energy.eu','smtp:alex@ark-energy.it']};
  }}),/SENDER_CHANGED/);
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

test('one ARK logo accompanies linked Artesian text in the shared signature',()=>{
  const p=profileForSender(graph,sender,branding);
  const textBundle={...bundle,branding:{...branding,artesianWordmark:false}};
  const html=renderSignature(textBundle,p);
  assert.equal((html.match(/<img /g)||[]).length,1);
  assert.ok(html.includes('alt="ARK"'));
  assert.ok(html.includes('href="https://www.artesian.cloud/"'));
  assert.ok(!html.includes('cid:artesian-test.png'));
  const withoutLink=renderSignature({...bundle,branding:{...branding,artesianWebsite:'',artesianWordmark:false}},p);
  assert.ok(withoutLink.includes('Home of Artesian'));
  assert.ok(!withoutLink.includes('href=""'));
});

test('the official small wordmark is linked in full signatures and becomes text in replies',()=>{
  const p=profileForSender(graph,sender,branding);
  const html=renderSignature(bundle,p);
  assert.equal((html.match(/<img /g)||[]).length,2);
  assert.ok(html.includes('cid:artesian-test.png'));
  assert.ok(html.includes('width="92" height="32"'));
  assert.ok(html.includes('width="64" height="20"'));
  assert.ok(!renderSignature(bundle,p,{compact:true}).includes('<img'));
});

test('the central country line replaces employee cities in HTML and plain text',()=>{
  const p=profileForSender(graph,sender,branding);
  const countries={...bundle,branding:{...branding,locationLine:'Ireland · Italy',showOfficeLocation:false}};
  for(const output of [renderSignature(countries,p),plainSignature(countries,p)]) {
    assert.ok(output.includes('Ireland · Italy'));
    assert.ok(!output.includes('Dublin'));
  }
});

test('email formatting survives removal of either style blocks or inline styles',()=>{
  const p=profileForSender(graph,sender,branding);
  for(const compact of [false,true]) {
    const html=renderSignature(bundle,p,{compact,officeCss:true});
    // The renderer prepends one stylesheet. Remove that known prefix to
    // simulate delivery behaviour; this fixture operation is not a sanitizer.
    assert.ok(html.startsWith('<style type="text/css">'));
    const styleEnd=html.indexOf('</style>');
    assert.ok(styleEnd>0);
    const withoutStyleBlock=html.slice(styleEnd+'</style>'.length);
    assert.ok(withoutStyleBlock.includes('font-family:Arial,Helvetica,sans-serif'));
    assert.ok(withoutStyleBlock.includes(compact?'color:#0074EA':'color:#102326'));
    assert.ok(withoutStyleBlock.includes('text-decoration:none'));
    assert.ok(withoutStyleBlock.includes('<strong'));
    assert.ok(withoutStyleBlock.includes('Alex Example'));
    assert.equal((withoutStyleBlock.match(/<img /g)||[]).length,compact?0:2);
    assert.ok(!withoutStyleBlock.includes('display:none'));
    const withoutInlineStyles=html.replace(/\sstyle="[^"]*"/g,'');
    assert.ok(withoutInlineStyles.includes('<style'));
    assert.ok(withoutInlineStyles.includes('font-family:Arial,Helvetica,sans-serif'));
    assert.ok(withoutInlineStyles.includes(' class="arksig_'));
    assert.ok(html.length<30000);
  }
});

test('mobile display spacing preserves callable numbers and leaves unfamiliar formats alone',()=>{
  assert.equal(formatPhoneNumber('+393939946956'),'+39 393 994 6956');
  assert.equal(formatPhoneNumber('+353831432590'),'+353 83 143 2590');
  assert.equal(formatPhoneNumber('+44 20 7123 4567 ext 12'),'+44 20 7123 4567 ext 12');
  const p=profileForSender({...graph,businessPhones:['+393939946956','+353831432590']},sender,branding);
  const html=renderSignature(bundle,p);
  assert.ok(html.includes('href="tel:+393939946956"'));
  assert.ok(html.includes('+39&#160;393&#160;994&#160;6956'));
  assert.ok(plainSignature(bundle,p).includes('+353 83 143 2590'));
});
