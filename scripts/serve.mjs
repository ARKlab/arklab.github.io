import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import {resolve,extname} from 'node:path';
import {fileURLToPath} from 'node:url';
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.json':'application/json','.css':'text/css','.xml':'application/xml','.png':'image/png'};
export function createPreviewServer(directory) {
  const root=resolve(directory);
  return createServer(async(req,res)=>{
    try{
      const path=decodeURIComponent(new URL(req.url,'http://localhost').pathname);
      let file=resolve(root,'.'+path);
      if(file!==root&&!file.startsWith(root+'/')){res.writeHead(403);res.end();return;}
      if(path.endsWith('/'))file=resolve(file,'index.html');
      // Read once, without a separate file check that could become stale.
      const bytes=await readFile(file);
      res.writeHead(200,{'Content-Type':types[extname(file)]||'application/octet-stream','Cache-Control':'no-store'});res.end(bytes);
    }catch{res.writeHead(404);res.end('Not found');}
  });
}
if(process.argv[1] && resolve(process.argv[1])===fileURLToPath(import.meta.url)) {
  createPreviewServer('dist').listen(8766,'127.0.0.1',()=>console.log('Preview http://127.0.0.1:8766 — browser review only; Outlook needs HTTPS.'));
}
