import test from 'node:test';
import assert from 'node:assert/strict';
import {jsonRequest,fetchBundle} from '../src/network.js';
import {errorText,signatureError,supportDetails} from '../src/errors.js';

test('HTTP timeout aborts the fetch and reports a timeout, not the resulting abort error',async t=>{
  let signal;
  t.mock.method(globalThis,'fetch',(_url,options)=>new Promise((_,reject)=>{
    signal=options.signal;
    signal.addEventListener('abort',()=>reject(new Error('abort')));
  }));
  await assert.rejects(jsonRequest('https://example.com/',{},5),{message:'REQUEST_TIMEOUT'});
  assert.equal(signal.aborted,true);
});

test('a stalled response body is bounded too',async t=>{
  let signal;
  t.mock.method(globalThis,'fetch',async(_url,options)=>{
    signal=options.signal;return {ok:true,json:()=>new Promise(()=>{})};
  });
  await assert.rejects(jsonRequest('https://example.com/',{},5),{message:'REQUEST_TIMEOUT'});
  assert.equal(signal.aborted,true);
});

test('settings failure retains its stage but never a raw network error',async t=>{
  t.mock.method(globalThis,'fetch',async()=>{throw new Error('private URL with sensitive query');});
  await assert.rejects(fetchBundle('https://example.com/'),error=>{
    assert.equal(error.stage,'settings');
    assert.equal(error.message,'REQUEST_NETWORK_ERROR');
    assert.ok(errorText(error).includes('company signature settings'));
    assert.equal(supportDetails(error).includes('sensitive'),false);
    return true;
  });
});

test('support details distinguish a broker rejection and a timeout without leaking provider messages',()=>{
  const broker=signatureError('SIGN_IN_BROKER_REJECTED','sign-in',{
    message:'AADSTS7000024: raw token and employee details',correlationId:'not-a-uuid-private-value'
  });
  assert.ok(errorText(broker).includes('Microsoft rejected'));
  assert.ok(supportDetails(broker).includes('AADSTS7000024'));
  assert.equal(supportDetails(broker).includes('employee'),false);
  assert.equal(supportDetails(broker).includes('not-a-uuid'),false);
  assert.ok(errorText(signatureError('SIGN_IN_TIMEOUT','sign-in')).includes('time limit'));
});

test('numeric and symbolic sign-in codes are retained without copying arbitrary payloads',()=>{
  for(const source of [{errorCode:'7000024'},{errorCode:7000024},{errorCode:'AADSTS7000024'}]) {
    assert.equal(signatureError('SIGN_IN_BROKER_REJECTED','sign-in',source).microsoftCode,'7000024');
  }
  const safe=signatureError('SIGN_IN_UNAVAILABLE','sign-in',{errorCode:'temporarily_unavailable',subError:'bad_token'});
  safe.attempts=2;
  assert.ok(supportDetails(safe).includes('Sign-in code: temporarily_unavailable'));
  assert.ok(supportDetails(safe).includes('Token attempts: 2'));
  const privateError=signatureError('SIGN_IN_UNAVAILABLE','sign-in',{errorCode:'private_employee_address',subError:'PRIVATE_TOKEN',message:'PRIVATE_PAYLOAD'});
  assert.equal(supportDetails(privateError).includes('private_employee'),false);
  assert.equal(supportDetails(privateError).includes('PRIVATE_'),false);
});
