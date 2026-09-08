export async function jsonRequest(url, options={}, timeoutMs=8000) {
  let timer;
  try {
    return await Promise.race([
      (async()=>{
        const response=await fetch(url, {...options, redirect:'error'});
        if (!response.ok) throw new Error('REQUEST_FAILED_'+response.status);
        return response.json();
      })(),
      new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('REQUEST_TIMEOUT')),timeoutMs);})
    ]);
  } finally { clearTimeout(timer); }
}

export function fetchBundle(siteUrl) {
  const url=new URL('signature-bundle.json',siteUrl);
  url.searchParams.set('check',Date.now().toString());
  return jsonRequest(url.href,{cache:'no-store',credentials:'omit'});
}
