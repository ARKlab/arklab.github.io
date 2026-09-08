import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {once} from 'node:events';
import {createPreviewServer} from '../scripts/serve.mjs';

test('preview serves index pages and assets, and rejects missing or escaping paths',async t=>{
  const directory=await mkdtemp(join(tmpdir(),'ark-preview-'));
  const root=join(directory,'dist');
  await mkdir(join(root,'assets'),{recursive:true});
  await mkdir(join(root,'guide'));
  await writeFile(join(root,'index.html'),'<h1>Preview</h1>');
  await writeFile(join(root,'guide','index.html'),'<h1>Guide</h1>');
  await writeFile(join(root,'assets','sample.json'),'{}');
  await writeFile(join(directory,'outside.txt'),'Outside root');
  const server=createPreviewServer(root);
  t.after(async()=>{
    if(server.listening) {
      server.closeAllConnections();
      await new Promise((resolve,reject)=>server.close(error=>error?reject(error):resolve()));
    }
    await rm(directory,{recursive:true,force:true});
  });
  server.listen(0,'127.0.0.1');
  await once(server,'listening');
  const origin='http://127.0.0.1:'+server.address().port;
  for(const [path,body,type] of [
    ['/','<h1>Preview</h1>','text/html; charset=utf-8'],
    ['/guide/','<h1>Guide</h1>','text/html; charset=utf-8'],
    ['/assets/sample.json','{}','application/json']
  ]) {
    const response=await fetch(origin+path);
    assert.equal(response.status,200);
    assert.equal(response.headers.get('content-type'),type);
    assert.equal(await response.text(),body);
  }
  assert.equal((await fetch(origin+'/missing.html')).status,404);
  assert.equal((await fetch(origin+'/..%2Foutside.txt')).status,403);
  assert.equal((await fetch(origin+'/%ZZ')).status,404);
});
