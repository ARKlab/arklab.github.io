import test from 'node:test';
import assert from 'node:assert/strict';
import {runInNewContext} from 'node:vm';
import {build} from 'esbuild';

const compiled=await build({entryPoints:['src/auth.js'],bundle:true,write:false,format:'cjs',plugins:[{
  name:'auth-fixtures',setup(builder){
    builder.onResolve({filter:/^(@azure\/msal-browser|\.\/network\.js)$/},args=>({path:args.path,namespace:'fixture'}));
    builder.onLoad({filter:/.*/,namespace:'fixture'},args=>({contents:args.path==='@azure/msal-browser'
      ? 'export const createNestablePublicClientApplication=(...args)=>fixtures.create(...args); export const InteractionRequiredAuthError=fixtures.InteractionRequiredAuthError;'
      : 'export const jsonRequest=(...args)=>fixtures.graph(...args);'}));
  }
}]});
const config={clientId:'a3ed4a3d-10ac-4015-a703-51181aa7365e',tenantId:'ab533857-aa85-4d60-bd9e-652773f2a06e'};
class InteractionRequiredAuthError extends Error {}
function service({silent,popup,graph,create,shortTimeout=false,now=()=>Date.now(),onDelay=()=>{}}={}) {
  const calls={silent:[],popup:[],graph:[],create:0};
  const token={accessToken:'PRIVATE_TOKEN',account:{tenantId:config.tenantId}};
  const module={exports:{}};
  runInNewContext(compiled.outputFiles[0].text,{module,clearTimeout,
    Date:class extends Date {static now(){return now();}},
    crypto:{randomUUID:()=> 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee'},
    setTimeout:(fn,ms)=>setTimeout(()=>{if(ms===1000)onDelay();fn();},ms===1000?0:shortTimeout&&ms===15000?5:ms),
    Office:{context:{requirements:{isSetSupported:()=>true}}},
    fixtures:{InteractionRequiredAuthError,
      create:async()=>{calls.create++;if(create) await create(calls.create);return {
        acquireTokenSilent:async request=>{calls.silent.push(request);return silent?silent(request,calls.silent.length):token;},
        acquireTokenPopup:async request=>{calls.popup.push(request);return popup?popup(request):token;}
      };},
      graph:async(...args)=>{calls.graph.push(args);return graph?graph(calls.graph.length):{displayName:'Private Name'};}
    }
  });
  return {...module.exports,calls,token};
}

test('normal SSO keeps User.Read and Graph /me with the mailbox hint',async()=>{
  const s=service();
  await s.graphProfile(config,{loginHint:'employee@ark-energy.eu'});
  assert.equal(s.calls.silent[0].scopes.join(','),'User.Read');
  assert.equal(s.calls.silent[0].loginHint,'employee@ark-energy.eu');
  assert.equal(s.calls.silent[0].forceRefresh,undefined);
  assert.ok(s.calls.graph[0][0].startsWith('https://graph.microsoft.com/v1.0/me?'));
  assert.equal(s.calls.popup.length,0);
});

test('Graph rejects a cached token: refresh once, then return the profile',async()=>{
  const s=service({graph:attempt=>{if(attempt===1) throw new Error('REQUEST_FAILED_401');return {ok:true};}});
  assert.equal((await s.graphProfile(config)).ok,true);
  assert.equal(s.calls.silent.length,2);
  assert.equal(s.calls.silent[1].forceRefresh,true);
});

test('a second Graph 401 stops, and a 403 is never retried',async()=>{
  for(const [code,count] of [['401',2],['403',1]]) {
    const s=service({graph:()=>{throw new Error('REQUEST_FAILED_'+code);}});
    await assert.rejects(s.graphProfile(config),e=>e.message==='REQUEST_FAILED_'+code&&e.stage==='profile');
    assert.equal(s.calls.silent.length,count);
    assert.equal(s.calls.popup.length,0);
  }
});

test('broker rejection preserves safe references without opening a popup or retrying indefinitely',async()=>{
  const correlationId='01a086e4-99b7-7ad8-af76-63dff13599db';
  const s=service({silent:()=>{throw Object.assign(new Error('AADSTS7000024: private provider payload PRIVATE_TOKEN employee@example.com'),{correlationId});}});
  await assert.rejects(s.graphProfile(config,{interactive:true}),e=>{
    assert.equal(e.message,'SIGN_IN_BROKER_REJECTED');
    assert.equal(e.microsoftCode,'7000024');
    assert.equal(e.correlationId,correlationId);
    assert.equal(JSON.stringify(e).includes('PRIVATE_TOKEN'),false);
    return true;
  });
  assert.equal(s.calls.silent.length,2);
  assert.equal(s.calls.silent[1].forceRefresh,true);
  assert.equal(s.calls.popup.length,0);
  assert.equal(s.calls.graph.length,0);
});

test('explicit recovery bypasses the MSAL token cache',async()=>{
  const s=service();
  await s.graphProfile(config,{interactive:true,recover:true});
  assert.equal(s.calls.silent[0].forceRefresh,true);
});

test('interaction is permitted only after an explicit click and an interaction-required error',async()=>{
  const s=service({silent:()=>{throw new InteractionRequiredAuthError('consent_required');}});
  await assert.rejects(s.graphProfile(config),{message:'SIGN_IN_REQUIRED'});
  assert.equal(s.calls.popup.length,0);
  await s.graphProfile(config,{interactive:true});
  assert.equal(s.calls.popup.length,1);
});

test('popup errors retain the same safe broker diagnostic',async()=>{
  const s=service({silent:()=>{throw new InteractionRequiredAuthError();},popup:()=>{throw new Error('AADSTS7000024: private');}});
  await assert.rejects(s.graphProfile(config,{interactive:true}),{message:'SIGN_IN_BROKER_REJECTED'});
  assert.equal(s.calls.graph.length,0);
  assert.equal(s.calls.popup.length,1);
  assert.equal(s.calls.silent.length,1);
});

test('a completed broker rejection recovers silently with one fresh-token request',async()=>{
  const s=service({silent:(_request,n)=>{
    if(n===1) throw Object.assign(new Error('broker rejected'),{errorCode:'7000024'});
    return s.token;
  }});
  const profile=await s.graphProfile(config,{loginHint:'employee@ark-energy.eu',context:{deadlineAt:Date.now()+60000}});
  assert.ok(profile.displayName);
  assert.equal(s.calls.silent.length,2);
  assert.equal(s.calls.silent[1].forceRefresh,true);
  assert.equal(s.calls.silent[1].loginHint,'employee@ark-energy.eu');
  assert.equal(s.calls.silent[1].scopes.join(','),'User.Read');
  assert.equal(s.calls.graph.length,1);
  assert.equal(s.calls.popup.length,0);
});

test('the known broker description is recognized even without an AADSTS prefix',async()=>{
  const s=service({silent:()=>{throw Object.assign(new Error('Inconsistent broker application IDs asserted by incoming credentials.'),{errorCode:'invalid_grant'});}});
  await assert.rejects(s.graphProfile(config),e=>e.message==='SIGN_IN_BROKER_REJECTED'&&e.msalCode==='invalid_grant'&&e.attempts===2&&e.correlationId==='aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee');
  assert.equal(s.calls.silent.length,2);
});

test('identified transient errors retry once; policy, consent and unknown failures do not',async()=>{
  for(const code of ['temporarily_unavailable','server_error','bridge_connection_reset','no_network_connectivity','invalid_grant','unknown_error','consent_required']) {
    const s=service({silent:()=>{throw Object.assign(new Error('PRIVATE_PROVIDER_MESSAGE'),{errorCode:code});}});
    await assert.rejects(s.graphProfile(config),e=>e.message==='SIGN_IN_UNAVAILABLE'&&!JSON.stringify(e).includes('PRIVATE_PROVIDER_MESSAGE'));
    assert.equal(s.calls.silent.length,['invalid_grant','unknown_error','consent_required'].includes(code)?1:2);
    assert.equal(s.calls.popup.length,0);
  }
});

test('acquisition recovery and Graph 401 share a total two-attempt limit',async()=>{
  const s=service({silent:(_request,n)=>{if(n===1)throw new Error('AADSTS7000024: rejected');return s.token;},graph:()=>{throw new Error('REQUEST_FAILED_401');}});
  await assert.rejects(s.graphProfile(config),e=>e.message==='REQUEST_FAILED_401'&&e.attempts===2);
  assert.equal(s.calls.silent.length,2);
  assert.equal(s.calls.graph.length,1);
});

test('cancellation during retry delay prevents a second token request',async()=>{
  const context={cancelled:false,deadlineAt:Date.now()+60000};
  const s=service({silent:()=>{throw new Error('AADSTS7000024: rejected');},onDelay:()=>{context.cancelled=true;}});
  await assert.rejects(s.graphProfile(config,{context}),{message:'REQUEST_TIMEOUT'});
  assert.equal(s.calls.silent.length,1);
  assert.equal(s.calls.graph.length,0);
});

test('the remaining event budget reserves time for Graph and insertion',async()=>{
  let now=0;
  const s=service({now:()=>now,silent:()=>{now=48000;throw new Error('AADSTS7000024: rejected');}});
  await assert.rejects(s.graphProfile(config,{context:{deadlineAt:60000}}),{message:'REQUEST_TIMEOUT'});
  assert.equal(s.calls.silent.length,1);
  assert.equal(s.calls.graph.length,0);
});

test('a late token or profile cannot continue a cancelled automatic event',async()=>{
  for(const phase of ['token','profile']) {
    const context={cancelled:false};
    const s=service({silent:()=>{if(phase==='token')context.cancelled=true;return s.token;},graph:()=>{context.cancelled=true;return {ok:true};}});
    await assert.rejects(s.graphProfile(config,{context}),{message:'REQUEST_TIMEOUT'});
    assert.equal(s.calls.graph.length,phase==='token'?0:1);
  }
});

test('an expired event never initializes the broker',async()=>{
  const s=service();
  await assert.rejects(s.graphProfile(config,{context:{deadlineAt:Date.now()-1}}),{message:'REQUEST_TIMEOUT'});
  assert.equal(s.calls.create,0);
});

test('a stalled broker times out; its late token cannot trigger a profile request',async()=>{
  let resolve;
  const s=service({shortTimeout:true,silent:()=>new Promise(r=>{resolve=r;})});
  await assert.rejects(s.graphProfile(config),{message:'SIGN_IN_TIMEOUT'});
  await assert.rejects(s.graphProfile(config,{interactive:true,recover:true}),{message:'SIGN_IN_PENDING'});
  assert.equal(s.calls.silent.length,1);
  resolve(s.token);
  await new Promise(r=>setTimeout(r,5));
  assert.equal(s.calls.graph.length,0);
  assert.equal(s.calls.popup.length,0);
});

test('a failed initialization is discarded so the next attempt can recover',async()=>{
  const s=service({create:n=>{if(n===1) throw new Error('private init failure');}});
  await assert.rejects(s.graphProfile(config),{message:'SIGN_IN_UNAVAILABLE'});
  await s.graphProfile(config);
  assert.equal(s.calls.create,2);
});

test('a token for a different tenant never reaches Graph',async()=>{
  const s=service({silent:()=>({accessToken:'PRIVATE_TOKEN',account:{tenantId:'another-tenant'}})});
  await assert.rejects(s.graphProfile(config),{message:'TENANT_MISMATCH'});
  assert.equal(s.calls.graph.length,0);
});
