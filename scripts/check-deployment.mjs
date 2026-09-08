import {readFile} from 'node:fs/promises';
const config=JSON.parse(await readFile('deployment.json','utf8'));
const origin=new URL(config.siteUrl).origin;
let failed=false;
for(const path of ['/','.well-known/microsoft-officeaddins-allowed.json','runtime.js','runtime.html','taskpane.html','signature-bundle.json','manifest.xml','assets/icon-64.png']) {
  const response=await fetch(new URL(path,origin+'/'),{redirect:'error'});
  if(!response.ok){failed=true;console.error(path,response.status);continue;}
  if(path.includes('well-known')) {
    const data=await response.json();
    if(!data.allowed?.includes(origin+'/runtime.js')){failed=true;console.error('Runtime missing from well-known file.');continue;}
  }
  console.log('OK',path);
}
if(!config.tenantId||!config.clientId) console.log('Microsoft 365 registration is still required.');
process.exitCode=failed?1:0;
