import {fetchBundle} from './network.js';
import {renderSignature} from './render.js';
const sample={name:'First name Last name',title:'Job title',email:'name@example.com',phones:['+353 00 000 0000'],office:'Office city, Country'};
fetchBundle(new URL('./',location.href).href).then(bundle=>{
  document.getElementById('full-preview').innerHTML=renderSignature(bundle,sample,{images:'preview'});
  document.getElementById('reply-preview').innerHTML=renderSignature(bundle,sample,{images:'preview',compact:true});
  const ready=bundle.deployment.clientId && bundle.deployment.tenantId;
  document.getElementById('deployment-status').textContent=bundle.enabled?'Automatic insertion enabled · '+bundle.revision:ready?'Automatic insertion paused · '+bundle.revision:'Prepared · Microsoft 365 setup pending';
}).catch(()=>{document.getElementById('deployment-status').textContent='Preview unavailable. Please reload.';});
