import {fetchBundle} from './network.js';
import {graphProfile,isConfigured} from './auth.js';
import {profileForSender,renderSignature} from './render.js';
import {applySignature,officeCall} from './office-flow.js';

const status=document.getElementById('status');
const connect=document.getElementById('connect');
const apply=document.getElementById('apply');
let bundle;
function say(message){status.textContent=message;}
function errorText(error) {
  const messages={ADMIN_SETUP_REQUIRED:'Your administrator still needs to connect Microsoft 365.',OUTLOOK_UPDATE_REQUIRED:'Update Outlook to a version that supports this add-in.',SIGN_IN_REQUIRED:'Choose Connect Microsoft 365 to sign in.',UNAPPROVED_SENDER:'This sending address is not configured for ARK signatures.'};
  return messages[error.message]||'Could not refresh your signature. Please try again or contact your administrator.';
}
Office.onReady(async()=>{
  try {
    bundle=await fetchBundle(__SITE_URL__);
    connect.disabled=!isConfigured(bundle.deployment);
    say(!isConfigured(bundle.deployment)?'Administrator setup is pending.':bundle.enabled?'Connect to preview your centrally managed signature.':'The pilot is paused. You can connect and preview your details.');
  }catch{say('Could not load the shared signature settings.');}
});
connect.addEventListener('click',async()=>{
  connect.disabled=true; apply.disabled=true;
  try {
    bundle=await fetchBundle(__SITE_URL__);
    const graph=await graphProfile(bundle.deployment,{interactive:true,loginHint:Office.context.mailbox.userProfile.emailAddress});
    const item=Office.context.mailbox.item;
    const sender=typeof item?.from?.getAsync==='function'?await officeCall(item.from,'getAsync'):{displayName:graph.displayName,emailAddress:Office.context.mailbox.userProfile.emailAddress};
    const profile=profileForSender(graph,sender,bundle.branding);
    document.getElementById('signature').innerHTML=renderSignature(bundle,profile,{images:'preview'});
    apply.disabled=!bundle.enabled || !item?.body?.setSignatureAsync;
    say(bundle.enabled?'Connected. Your signature uses the current company template.':'Connected. Automatic insertion is paused by your administrator.');
  }catch(error){say(errorText(error));}
  finally{connect.disabled=false;}
});
apply.addEventListener('click',async()=>{
  apply.disabled=true;
  try {
    bundle=await fetchBundle(__SITE_URL__);
    const result=await applySignature({item:Office.context.mailbox.item,bundle,getGraph:()=>graphProfile(bundle.deployment,{loginHint:Office.context.mailbox.userProfile.emailAddress})});
    say(result.status==='applied'?'Your signature has been updated in this message.':'Automatic insertion is paused by your administrator.');
  }catch(error){say(errorText(error));}
  finally{apply.disabled=!bundle?.enabled;}
});
