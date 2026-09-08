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
  const messages={ADMIN_SETUP_REQUIRED:'Your administrator still needs to connect Microsoft 365.',OUTLOOK_UPDATE_REQUIRED:'Update Outlook to a version that supports this add-in.',SIGN_IN_REQUIRED:'Microsoft needs your permission or a sign-in check before your profile can be loaded. Choose Continue with Microsoft 365.',UNAPPROVED_SENDER:'This sending address is not configured for ARK signatures.'};
  return messages[error.message]||'Could not refresh your signature. Please try again or contact your administrator.';
}
Office.onReady(async()=>{
  try {
    bundle=await fetchBundle(__SITE_URL__);
    if(!isConfigured(bundle.deployment)) {say('Administrator setup is pending.');return;}
    // Opening the panel may reuse Outlook's session, but must never launch a sign-in popup.
    await previewSignature(false);
  }catch{say('Could not load the shared signature settings.');}
});
async function previewSignature(interactive) {
  connect.disabled=true; apply.disabled=true;
  say('Loading your profile using your Outlook sign-in…');
  try {
    bundle=await fetchBundle(__SITE_URL__);
    const graph=await graphProfile(bundle.deployment,{interactive,loginHint:Office.context.mailbox.userProfile.emailAddress});
    const item=Office.context.mailbox.item;
    const sender=typeof item?.from?.getAsync==='function'?await officeCall(item.from,'getAsync'):{displayName:graph.displayName,emailAddress:Office.context.mailbox.userProfile.emailAddress};
    const profile=profileForSender(graph,sender,bundle.branding);
    document.getElementById('signature').innerHTML=renderSignature(bundle,profile,{images:'preview'});
    connect.textContent='Refresh preview';
    apply.disabled=!bundle.enabled || !item?.body?.setSignatureAsync;
    say(bundle.enabled?'Connected. Your signature uses the current company template.':'Connected. Automatic insertion is paused by your administrator.');
  }catch(error){connect.textContent='Continue with Microsoft 365';say(errorText(error));}
  finally{connect.disabled=false;}
}
connect.addEventListener('click',()=>previewSignature(true));
apply.addEventListener('click',async()=>{
  apply.disabled=true;
  try {
    bundle=await fetchBundle(__SITE_URL__);
    const result=await applySignature({item:Office.context.mailbox.item,bundle,getGraph:()=>graphProfile(bundle.deployment,{loginHint:Office.context.mailbox.userProfile.emailAddress})});
    say(result.status==='applied'?'Your signature has been updated in this message.':'Automatic insertion is paused by your administrator.');
  }catch(error){say(errorText(error));}
  finally{apply.disabled=!bundle?.enabled;}
});
