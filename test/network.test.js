import test from 'node:test';
import assert from 'node:assert/strict';
import {jsonRequest,fetchBundle} from '../src/network.js';
import {errorText,signatureError,supportDetails,failureReference} from '../src/errors.js';
import {claimsFor} from '../src/graph-error.js';

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

test('Office support references retain allowlisted steps and numeric codes only',()=>{
  const error=signatureError('OUTLOOK_9050','session-read');
  assert.ok(supportDetails(error).includes('Step: session-read'));
  assert.ok(supportDetails(error).includes('Result: OUTLOOK_9050'));
  assert.equal(failureReference(error),'session-read/OUTLOOK_9050');
  for(const message of ['OUTLOOK_FAILED','OUTLOOK_UPDATE_REQUIRED','SIGN_IN_REQUIRED','REQUEST_FAILED_503']) {
    const failure=signatureError(message,'session-read');
    assert.equal(failureReference(failure),'session-read');
    assert.ok(supportDetails(failure).includes('Result: '+message));
    assert.equal(failureReference(signatureError(message,'PRIVATE_STEP')),'');
  }
  for(const message of ['OUTLOOK_PRIVATE_TOKEN','SIGN_IN_PRIVATE_TOKEN','OUTLOOK_9050 employee@example.com']) {
    const failure=signatureError(message,'PRIVATE_STEP');
    assert.equal(failureReference(failure),'');
    assert.equal(supportDetails(failure).includes('PRIVATE_'),false);
    assert.equal(supportDetails(failure).includes('employee@'),false);
  }
});

const graphUrl='https://graph.microsoft.com/v1.0/me?$select=displayName';
test('Graph diagnostics retain safe codes and request IDs, not tokens or provider messages',async t=>{
  const requestId='aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
  const claims=JSON.stringify({access_token:{test:{value:'PRIVATE_CLAIM'}}});
  t.mock.method(globalThis,'fetch',async()=>new Response(JSON.stringify({error:{code:'InvalidAuthenticationToken',message:'PRIVATE_TOKEN employee@example.org',innerError:{'request-id':requestId}}}),{status:401,headers:{'www-authenticate':`Bearer realm="", error="insufficient_claims", claims="${btoa(claims)}"`}}));
  await assert.rejects(jsonRequest(graphUrl),error=>{
    assert.equal(error.httpStatus,401);assert.equal(error.graphCode,'InvalidAuthenticationToken');
    assert.equal(error.requestId,requestId);assert.equal(claimsFor(error),claims);
    const details=supportDetails(error);
    assert.ok(details.includes('Graph request ID: '+requestId));assert.ok(details.includes('Claims challenge: usable'));
    assert.equal((JSON.stringify(error)+details).includes('PRIVATE_'),false);
    assert.equal(details.includes('employee@'),false);
    return true;
  });
});

test('claims are never accepted from settings, a different origin or a 403',async t=>{
  const claims=btoa('{"access_token":{"x":{"value":"PRIVATE"}}}');
  let status=401;
  t.mock.method(globalThis,'fetch',async()=>new Response('{}',{status,headers:{'www-authenticate':`Bearer error="insufficient_claims", claims="${claims}"`}}));
  for(const url of ['https://example.org/v1.0/me','https://graph.microsoft.com/v1.0/users','https://graph.microsoft.com.evil.example/v1.0/me']) {
    await assert.rejects(jsonRequest(url),error=>!claimsFor(error)&&!error.claimsChallenge);
  }
  status=403;
  await assert.rejects(jsonRequest(graphUrl),error=>!claimsFor(error)&&error.httpStatus===403);
});

test('malformed, oversized and mixed authentication challenges cannot become broker input',async t=>{
  const valid=btoa('{"access_token":{"test":{"value":"PRIVATE_CLAIM"}}}');
  let header;
  t.mock.method(globalThis,'fetch',async()=>new Response('{"error":{"code":"PRIVATE_CODE","message":"PRIVATE_TOKEN","innerError":{"request-id":"PRIVATE_ID"}}}',{status:401,headers:{'request-id':'PRIVATE_ID','www-authenticate':header}}));
  for(header of [
    'Basic claims="'+valid+'"',
    `Bearer error="insufficient_claims", claims="${valid}", claims="${valid}"`,
    `Bearer error="insufficient_claims", Basic claims="${valid}"`,
    `Bearer error="invalid_token", claims="${valid}"`,
    'Bearer error="insufficient_claims", claims="@@@"',
    'Bearer error="insufficient_claims", claims="'+btoa('{"__proto__":{},"access_token":{}}')+'"',
    'Bearer error="insufficient_claims", claims="'+btoa('{"access_token":{"\\u005f_proto__":{}}}')+'"',
    'Bearer error="insufficient_claims", claims="'+btoa('{"access_token":null}')+'"',
    'Bearer error="insufficient_claims", claims="'+btoa('{"access_token":{"x":"'+'a'.repeat(10000)+'"}}')+'"',
    'Bearer error="insufficient_claims", claims="'+btoa('not json')+'"'
  ]) {
    await assert.rejects(jsonRequest(graphUrl),error=>{
      assert.equal(claimsFor(error),undefined);
      assert.equal((JSON.stringify(error)+supportDetails(error)).includes('PRIVATE_'),false);
      return true;
    });
  }
});

test('an oversized error body is discarded while HTTP status and header references survive',async t=>{
  t.mock.method(globalThis,'fetch',async()=>new Response('x'.repeat(20000),{status:401,headers:{'request-id':'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee'}}));
  await assert.rejects(jsonRequest(graphUrl),error=>error.httpStatus===401&&Boolean(error.requestId)&&!error.graphCode);
});

test('a stalled error body remains bounded by the request deadline',async t=>{
  t.mock.method(globalThis,'fetch',async()=>new Response(new ReadableStream({start(){}}),{status:401}));
  await assert.rejects(jsonRequest(graphUrl,{},5),{message:'REQUEST_TIMEOUT'});
});

test('diagnostics label OfficeOnline without guessing web versus new Windows and reject unsafe versions',t=>{
  const office={context:{diagnostics:{platform:'OfficeOnline',version:'16.0.12345.123'}}};
  const previous=globalThis.Office;
  globalThis.Office=office;
  t.after(()=>{if(previous===undefined) delete globalThis.Office;else globalThis.Office=previous;});
  const error=signatureError('REQUEST_FAILED_401','profile');
  const details=supportDetails(error);
  assert.ok(details.includes('Office platform: OfficeOnline'));
  assert.ok(details.includes('Office host version: 16.0.12345.123'));
  office.context.diagnostics={platform:'PRIVATE_PLATFORM',version:'PRIVATE_VERSION'};
  assert.equal(supportDetails(error).includes('PRIVATE_'),false);
});
