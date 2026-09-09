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
function service({silent,popup,graph,create,shortTimeout=false}={}) {
  const calls={silent:[],popup:[],graph:[],create:0};
  const token={accessToken:'PRIVATE_TOKEN',account:{tenantId:config.tenantId}};
  const module={exports:{}};
  runInNewContext(compiled.outputFiles[0].text,{module,clearTimeout,
    setTimeout:(fn,ms)=>setTimeout(fn,shortTimeout&&ms===15000?5:ms),
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
  assert.equal(s.calls.silent.length,1);
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
