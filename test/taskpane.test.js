import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {build} from 'esbuild';

const compiled=await build({
  entryPoints:['src/taskpane.js'],bundle:true,write:false,format:'iife',
  define:{__SITE_URL__:JSON.stringify('https://example.com/')},
  plugins:[{name:'profile-service-fixtures',setup(builder){
    builder.onResolve({filter:/^\.\/(auth|network)\.js$/},args=>({path:args.path,namespace:'fixture'}));
    builder.onLoad({filter:/.*/,namespace:'fixture'},args=>({contents:args.path.endsWith('auth.js')
      ? 'export const graphProfile=(...args)=>fixtures.graphProfile(...args); export const isConfigured=()=>true;'
      : 'export const fetchBundle=()=>fixtures.fetchBundle();'}));
  }}]
});
const bundle=JSON.parse(await readFile('branding.json','utf8'));
const template=await readFile('templates/full.html','utf8');

function panel({needsConsent=false,brokerFailure=false,settingsFailure=false,applyFailure=false}={}) {
  const elements=Object.fromEntries(['status','connect','apply','signature','support','diagnostics'].map(id=>[id,{
    disabled:true,textContent:'',innerHTML:'',handlers:{},
    addEventListener(event,handler){this.handlers[event]=handler;}
  }]));
  const calls=[];
  let ready;
  let fetchCount=0;
  const email='alex@ark-energy.eu';
  runInNewContext(compiled.outputFiles[0].text,{
    document:{getElementById:id=>elements[id]},
    Office:{onReady:handler=>{ready=handler;},context:{mailbox:{userProfile:{emailAddress:email},item:{from:{getAsync:callback=>callback({status:'succeeded',value:{emailAddress:email}})},body:{setSignatureAsync(){assert.fail('Opening a preview must not insert a signature.');}}}}}},
    URL,fixtures:{
      fetchBundle(){if(settingsFailure&&fetchCount++===0) throw Object.assign(new Error('REQUEST_NETWORK_ERROR'),{stage:'settings'});return this.bundle;},
      bundle:{enabled:true,deployment:{},branding:bundle,revision:'panel-test',templates:{full:template},assets:{ark:{base64:'eA=='},artesian:{base64:'eA=='}}},
      graphProfile:async(_config,request)=>{
        calls.push(request);
        if(brokerFailure&&calls.length===1) throw Object.assign(new Error('SIGN_IN_BROKER_REJECTED'),{stage:'sign-in',microsoftCode:'7000024'});
        if(applyFailure&&calls.length>1) throw new Error('SIGN_IN_UNAVAILABLE');
        if(needsConsent&&!request.interactive) throw new Error('SIGN_IN_REQUIRED');
        return {displayName:'Alex Example',mail:email,jobTitle:'Director',businessPhones:['+353 83 111 2222']};
      }
    }
  });
  return {elements,calls,ready};
}

test('opening the panel uses Outlook SSO without an interactive prompt or insertion',async()=>{
  const {elements,calls,ready}=panel();
  await ready();
  assert.equal(calls.length,1);
  assert.equal(calls[0].interactive,false);
  assert.equal(calls[0].loginHint,'alex@ark-energy.eu');
  assert.ok(elements.signature.innerHTML.includes('Alex Example'));
  assert.equal(elements.connect.textContent,'Refresh preview');
  assert.equal(elements.apply.disabled,false);
});

test('missing consent waits for an explicit click before requesting interaction',async()=>{
  const {elements,calls,ready}=panel({needsConsent:true});
  await ready();
  assert.equal(calls.length,1);
  assert.equal(calls[0].interactive,false);
  assert.equal(elements.apply.disabled,true);
  assert.equal(elements.connect.textContent,'Continue with Microsoft 365');
  assert.ok(elements.status.textContent.includes('permission or a sign-in check'));
  await elements.connect.handlers.click();
  assert.equal(calls.length,2);
  assert.equal(calls[1].interactive,true);
  assert.ok(elements.signature.innerHTML.includes('Alex Example'));
});

test('broker failure displays safe details and an explicit fresh-token retry recovers',async()=>{
  const {elements,calls,ready}=panel({brokerFailure:true});
  await ready();
  assert.equal(elements.connect.textContent,'Retry connection');
  assert.equal(elements.apply.disabled,true);
  assert.equal(elements.support.hidden,false);
  assert.ok(elements.diagnostics.textContent.includes('AADSTS7000024'));
  await elements.connect.handlers.click();
  assert.equal(calls[1].recover,true);
  assert.equal(elements.support.hidden,true);
  assert.equal(elements.apply.disabled,false);
});

test('initial settings failure leaves a working retry button',async()=>{
  const {elements,calls,ready}=panel({settingsFailure:true});
  await ready();
  assert.equal(calls.length,0);
  assert.equal(elements.connect.disabled,false);
  assert.equal(elements.apply.disabled,true);
  assert.ok(elements.status.textContent.includes('company signature settings'));
  await elements.connect.handlers.click();
  assert.equal(calls.length,1);
  assert.equal(elements.apply.disabled,false);
});

test('a manual refresh sign-in failure leaves insertion disabled until reconnection',async()=>{
  const {elements,ready}=panel({applyFailure:true});
  await ready();
  assert.equal(elements.apply.disabled,false);
  await elements.apply.handlers.click();
  assert.equal(elements.apply.disabled,true);
  assert.equal(elements.connect.disabled,false);
  assert.equal(elements.connect.textContent,'Retry connection');
  assert.equal(elements.signature.innerHTML,'');
});
