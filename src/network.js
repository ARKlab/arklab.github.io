import {signatureError} from './errors.js';

export async function jsonRequest(url, options={}, timeoutMs=8000) {
  let timer;
  const controller=typeof AbortController==='function'?new AbortController():undefined;
  const abort=()=>controller?.abort();
  if(options.signal?.aborted) abort();
  options.signal?.addEventListener('abort',abort,{once:true});
  try {
    return await Promise.race([
      (async()=>{
        let response;
        try {response=await fetch(url, {...options, signal:controller?.signal||options.signal,redirect:'error'});}
        catch {throw new Error('REQUEST_NETWORK_ERROR');}
        if (!response.ok) throw new Error('REQUEST_FAILED_'+response.status);
        try {return await response.json();}
        catch {throw new Error('REQUEST_INVALID_RESPONSE');}
      })(),
      new Promise((_,reject)=>{timer=setTimeout(()=>{reject(new Error('REQUEST_TIMEOUT'));abort();},timeoutMs);})
    ]);
  } finally {clearTimeout(timer);options.signal?.removeEventListener('abort',abort);}
}

export function fetchBundle(siteUrl) {
  const url=new URL('signature-bundle.json',siteUrl);
  url.searchParams.set('check',Date.now().toString());
  return jsonRequest(url.href,{cache:'no-store',credentials:'omit'}).catch(error=>{throw signatureError(error.message,'settings');});
}
