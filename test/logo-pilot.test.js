import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {isLogoPilotAccount,makeLogoPilotBundle} from '../src/experiments/logo-colours.js';
import {runLogoPilot,installLogoPilot} from '../src/experiments/logo-pilot.js';
import {renderSignature} from '../src/render.js';

const email='alex@ark-energy.eu';
const base={enabled:true,revision:'release-test',deployment:{},branding:JSON.parse(await readFile('branding.json','utf8')),assets:{ark:{filename:'original-ark.png',base64:'eA=='},artesian:{filename:'original-artesian.png',base64:'eA=='}},templates:{full:await readFile('templates/full.html','utf8'),reply:await readFile('templates/reply.html','utf8')}};
const assets={};
for(const [key,file] of Object.entries({arkFallback:'ark-fallback',arkWhite:'ark-white',artesianFallback:'artesian-fallback',artesianWhite:'artesian-white'})) {
  const bytes=await readFile(`public/experiments/logo-colours/${file}.png`);
  assets[key]={filename:`ark-signatures-trial-${file}-${createHash('sha256').update(bytes).digest('hex').slice(0,12)}.png`,base64:bytes.toString('base64')};
}
const profile={name:'Alex Example',email,title:'Director',phones:[],office:''};

// Use a synthetic profile for flow tests. A separate, unmocked test below
// verifies that the configured pilot address matches through real Web Crypto.
function assignFixture(t) {
  const realDigest=crypto.subtle.digest.bind(crypto.subtle);
  t.mock.method(crypto.subtle,'digest',async(algorithm,input)=>new TextDecoder().decode(input)==='ark-signature-logo-trial:'+email
    ? Uint8Array.from(Buffer.from('60e2cda6273e570f6f94f568358d16afd50331ca18f5baa29298e5d261661ac0','hex')).buffer
    : realDigest(algorithm,input));
}

function fixture({account=email,sender=email,composeType='newMail',bodyType='html',enabled=true,graphEmail=email}={}) {
  const calls={base:0,assets:0,graph:[],sets:[],attachments:[],existing:[]};
  const ok=(cb,value)=>cb({status:'succeeded',value});
  const item={
    from:{getAsync:cb=>ok(cb,{emailAddress:sender})},
    getComposeTypeAsync:cb=>ok(cb,{composeType}),
    getAttachmentsAsync:cb=>ok(cb,calls.existing),
    addFileAttachmentFromBase64Async:(data,name,options,cb)=>{calls.attachments.push({name,options});calls.existing.push({name,isInline:true});ok(cb,name);},
    body:{getTypeAsync:cb=>ok(cb,bodyType),setSignatureAsync:(html,options,cb)=>{calls.sets.push({html,options});ok(cb);}}
  };
  const input={mailbox:{userProfile:{emailAddress:account},item},fetchBase:async()=>{calls.base++;return {...base,enabled};},fetchAssets:async()=>{calls.assets++;return assets;},getGraph:async(_config,request)=>{calls.graph.push(request);return {displayName:profile.name,mail:graphEmail,userPrincipalName:graphEmail};}};
  return {calls,item,input};
}

test('configured pilot address matches using real Web Crypto',async()=>{
  assert.equal(await isLogoPilotAccount('francesco.arci@ark-energy.eu'),true);
  assert.equal(await isLogoPilotAccount(' Francesco.Arci@ARK-ENERGY.EU '),true);
  assert.equal(await isLogoPilotAccount('someone-else@ark-energy.eu'),false);
});

test('pilot account comparison normalises case and rejects other accounts',async t=>{
  assignFixture(t);
  assert.equal(await isLogoPilotAccount(' Alex@ARK-ENERGY.EU '),true);
  assert.equal(await isLogoPilotAccount('someone-else@ark-energy.eu'),false);
  assert.equal(await isLogoPilotAccount(undefined),false);
});

test('opening the panel for another account creates no controls or requests',async()=>{
  const f=fixture({account:'someone-else@ark-energy.eu'});
  assert.equal(await installLogoPilot({...f.input,document:{createElement(){assert.fail('No pilot UI');}}}),false);
  assert.equal(f.calls.base,0);assert.equal(f.calls.assets,0);assert.equal(f.calls.graph.length,0);assert.equal(f.calls.sets.length,0);
  await assert.rejects(runLogoPilot(f.input),/NOT_ASSIGNED/);
});

test('pilot render keeps normal output and compact replies unchanged',()=>{
  const before=JSON.stringify(base);
  const trial=makeLogoPilotBundle(base,assets);
  assert.equal(JSON.stringify(base),before);
  assert.equal(renderSignature(trial,profile,{compact:true,officeCss:true}),renderSignature({...base,revision:trial.revision},profile,{compact:true,officeCss:true}));
  const html=renderSignature(trial,profile,{officeCss:true});
  assert.equal((html.match(/<img /g)||[]).length,4);
  assert.ok(html.includes('@media (prefers-color-scheme:dark)'));
  assert.ok(html.includes('[data-ogsc] .arklogotrial-light'));
  assert.ok(html.includes('[data-ogsb] .arklogotrial-light'));
  assert.equal((html.match(/width="0" height="0"/g)||[]).length,2);
  assert.equal((html.match(/<!--\[if !mso\]><!-->/g)||[]).length,2);
  assert.equal(/<[^>]*\sclass="[^"]*"[^>]*\sclass=/.test(html),false);
  assert.ok(html.length<30000);
});

test('white images retain inline hiding when all style blocks are removed',()=>{
  const html=renderSignature(makeLogoPilotBundle(base,assets),profile,{officeCss:true});
  // Remove the two known leading stylesheets from this generated fixture.
  // This models a mail-client transformation; it is not an HTML sanitizer.
  let withoutStyles=html;
  for(let block=0;block<2;block++) {
    assert.ok(withoutStyles.startsWith('<style type="text/css">'));
    const end=withoutStyles.indexOf('</style>');
    assert.ok(end>0);
    withoutStyles=withoutStyles.slice(end+'</style>'.length);
  }
  assert.equal(withoutStyles.includes('<style'),false);
  for(const tag of withoutStyles.match(/<img\b[^>]*>/g).filter(tag=>tag.includes('-white-'))) {
    assert.ok(tag.includes('display:none;'));
    assert.ok(tag.includes('width="0" height="0"'));
  }
  assert.ok(withoutStyles.includes('ark-fallback'));
  assert.ok(withoutStyles.includes('artesian-fallback'));
});

test('pilot accepts only valid local PNG attachment names and a compatible template',()=>{
  assert.throws(()=>makeLogoPilotBundle(base,{...assets,arkWhite:{...assets.arkWhite,filename:'x" onerror="alert(1)'}}),/INVALID_ASSETS/);
  assert.throws(()=>makeLogoPilotBundle(base,{...assets,arkWhite:{...assets.arkWhite,base64:'<script>'}}),/INVALID_ASSETS/);
  assert.throws(()=>makeLogoPilotBundle({...base,templates:{...base.templates,full:'different template'}},assets),/TEMPLATE_CHANGED/);
});

test('new HTML trial embeds four PNGs once, and restore returns the standard design',async t=>{
  assignFixture(t);
  const f=fixture();
  await runLogoPilot(f.input);
  assert.equal(f.calls.sets.length,1);assert.equal(f.calls.attachments.length,4);
  assert.ok(f.calls.attachments.every(a=>a.options.isInline));
  assert.equal(f.calls.graph[0].interactive,false);
  await runLogoPilot(f.input);
  assert.equal(f.calls.attachments.length,4);
  const requests=f.calls.assets;
  await runLogoPilot({...f.input,restore:true});
  assert.equal(f.calls.assets,requests);
  assert.equal(f.calls.sets.length,3);
  assert.equal((f.calls.sets[2].html.match(/<img /g)||[]).length,2);
  assert.equal(f.calls.sets[2].html.includes('arklogotrial'),false);
});

test('shared sender, replies and plain text reject before any profile or artwork request',async t=>{
  assignFixture(t);
  for(const options of [{sender:'shared@ark-energy.eu'},{composeType:'reply'},{composeType:'forward'},{bodyType:'text'}]) {
    const f=fixture(options);
    await assert.rejects(runLogoPilot(f.input),/OWN_SENDER_ONLY|NEW_HTML_ONLY/);
    assert.equal(f.calls.base,0);assert.equal(f.calls.assets,0);assert.equal(f.calls.graph.length,0);assert.equal(f.calls.sets.length,0);
  }
});

test('wrong profile, sender changes and disabled deployment do not insert the trial',async t=>{
  assignFixture(t);
  const mismatch=fixture({graphEmail:'other@ark-energy.eu'});
  await assert.rejects(runLogoPilot(mismatch.input),/PROFILE_MISMATCH/);
  assert.equal(mismatch.calls.attachments.length,0);assert.equal(mismatch.calls.sets.length,0);
  const changed=fixture();
  const getGraph=changed.input.getGraph;
  changed.input.getGraph=async(...args)=>{
    changed.item.from.getAsync=cb=>cb({status:'succeeded',value:{emailAddress:'shared@ark-energy.eu'}});
    return getGraph(...args);
  };
  await assert.rejects(runLogoPilot(changed.input),/OWN_SENDER_ONLY/);
  assert.equal(changed.calls.attachments.length,0);assert.equal(changed.calls.sets.length,0);
  const paused=fixture({enabled:false});
  assert.equal((await runLogoPilot(paused.input)).status,'paused');
  assert.equal(paused.calls.assets,0);assert.equal(paused.calls.graph.length,0);assert.equal(paused.calls.sets.length,0);
});

function panelDocument() {
  const elements={};
  const make=()=>({disabled:false,children:[],handlers:{},style:{},setAttribute(){},append(...children){this.children.push(...children);for(const child of children)if(child.id)elements[child.id]=child;},addEventListener(event,handler){this.handlers[event]=handler;}});
  elements.connect=make();elements.apply=make();
  elements.signature={parentNode:{insertBefore(element){elements[element.id]=element;}}};
  return {elements,document:{createElement:make,getElementById:id=>elements[id]}};
}

test('pilot controls require a click; repeated clicks cannot create concurrent insertions',async t=>{
  assignFixture(t);
  const f=fixture();const doc=panelDocument();
  assert.equal(await installLogoPilot({...f.input,document:doc.document}),true);
  assert.equal(f.calls.base,0);assert.equal(f.calls.graph.length,0);assert.equal(f.calls.sets.length,0);
  const click=doc.elements['insert-test-logos'].handlers.click;
  await Promise.all([click(),click()]);
  assert.equal(f.calls.sets.length,1);
  assert.equal(doc.elements.connect.disabled,false);
  assert.equal(doc.elements['insert-test-logos'].disabled,false);
  assert.ok(doc.elements['logo-pilot-status'].textContent.includes('Test logos inserted'));
});

test('plain-text draft explains the restriction and restores controls without loading data',async t=>{
  assignFixture(t);
  const f=fixture({bodyType:'text'});const doc=panelDocument();
  await installLogoPilot({...f.input,document:doc.document});
  await doc.elements['insert-test-logos'].handlers.click();
  assert.match(doc.elements['logo-pilot-status'].textContent,/Plain-text messages cannot display logos/);
  assert.equal(f.calls.base,0);assert.equal(f.calls.assets,0);assert.equal(f.calls.graph.length,0);
  assert.equal(f.calls.attachments.length,0);assert.equal(f.calls.sets.length,0);
  for(const id of ['connect','apply','insert-test-logos','restore-standard-signature']) assert.equal(doc.elements[id].disabled,false);
});

test('pilot waits for normal profile loading and returns controls after failure',async t=>{
  assignFixture(t);
  const f=fixture();const doc=panelDocument();
  f.input.fetchAssets=async()=>{throw new Error('REQUEST_NETWORK_ERROR');};
  await installLogoPilot({...f.input,document:doc.document});
  doc.elements.connect.disabled=true;
  await doc.elements['insert-test-logos'].handlers.click();
  assert.equal(f.calls.base,0);
  doc.elements.connect.disabled=false;
  await doc.elements['insert-test-logos'].handlers.click();
  assert.equal(f.calls.sets.length,0);
  assert.equal(doc.elements.connect.disabled,false);
  assert.equal(doc.elements.apply.disabled,false);
  assert.ok(doc.elements['logo-pilot-status'].textContent.includes('could not complete'));
});
