import {signatureError} from './errors.js';
// Challenges stay in memory and are never enumerable or included in diagnostics.
const challenges=new WeakMap();
export function claimsFor(error) {return error&&challenges.get(error);}
export function copyClaims(source,target) {
  const claims=claimsFor(source);
  if(claims) challenges.set(target,claims);
  return target;
}

function decodeChallenge(header) {
  // Accept one unambiguous Bearer challenge. Ignore malformed/mixed schemes;
  // never use an authority or scope supplied by an HTTP header.
  if(typeof header!=='string'||header.length>16384||!/^Bearer\s/i.test(header)) return;
  let rest=header.replace(/^Bearer\s+/i,'');
  const fields=new Map();
  while(rest) {
    const match=/^([\w-]+)="([^"\\\r\n]*)"\s*(?:,\s*|$)/.exec(rest);
    if(!match||fields.has(match[1].toLowerCase())) return;
    fields.set(match[1].toLowerCase(),match[2]);
    rest=rest.slice(match[0].length);
  }
  const encoded=fields.get('claims');
  if(fields.get('error')!=='insufficient_claims'||!encoded||encoded.length>11000||!/^[A-Za-z0-9+/_-]+={0,2}$/.test(encoded)) return;
  try {
    const bytes=Uint8Array.from(atob(encoded.replace(/-/g,'+').replace(/_/g,'/')),c=>c.charCodeAt(0));
    if(bytes.length>8192) return;
    const decoded=new TextDecoder('utf-8',{fatal:true}).decode(bytes);
    const parsed=JSON.parse(decoded,(key,value)=>{
      if(['__proto__','prototype','constructor'].includes(key)) throw new Error('INVALID_CLAIMS');
      return value;
    });
    if(!parsed||Array.isArray(parsed)||Object.keys(parsed).join(',')!=='access_token'||!parsed.access_token||typeof parsed.access_token!=='object'||Array.isArray(parsed.access_token)) return;
    return JSON.stringify(parsed);
  }catch {return;}
}

async function smallJson(response) {
  // Error pages/messages can contain private data. Bound the read and discard
  // everything except allowlisted references at the diagnostic boundary.
  const reader=response.body?.getReader?.();
  if(!reader) return;
  try {
    const chunks=[];let size=0;
    for(;;) {
      const {done,value}=await reader.read();
      if(done) break;
      size+=value.byteLength;
      if(size>16384) return;
      chunks.push(value);
    }
    const bytes=new Uint8Array(size);let offset=0;
    for(const chunk of chunks) {bytes.set(chunk,offset);offset+=chunk.byteLength;}
    return JSON.parse(new TextDecoder().decode(bytes));
  }catch {return;}
  finally {void reader.cancel().catch(()=>{});}
}

export async function graphFailure(response) {
  const error=new Error('REQUEST_FAILED_'+response.status);
  error.httpStatus=response.status;
  error.requestId=response.headers?.get('request-id');
  const header=response.headers?.get('www-authenticate');
  if(response.status===401&&header) {
    const claims=decodeChallenge(header);
    error.claimsChallenge=claims?'usable':'not usable';
    if(claims) challenges.set(error,claims);
  }
  const body=await smallJson(response);
  error.graphCode=body?.error?.code;
  error.requestId ||= body?.error?.innerError?.['request-id'];
  return copyClaims(error,signatureError(error.message,'profile',error));
}
