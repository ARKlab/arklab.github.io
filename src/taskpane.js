import {fetchBundle} from './network.js';
import {graphProfile,isConfigured} from './auth.js';
import {profileForSender,renderSignature} from './render.js';
import {applySignature,officeCall} from './office-flow.js';
import {errorText,supportDetails} from './errors.js';
import {isMobileOutlook} from './platform.js';

const status=document.getElementById('status');
const connect=document.getElementById('connect');
const apply=document.getElementById('apply');
let bundle;
let recover=false;
let signInAgain=false;
let recoveryError;
let recoveryLoginHint;
function say(message){
  status.textContent=message;
  // Outlook may briefly retain the previous HTML shell while loading new JS.
  const support=document.getElementById('support');
  if(support) support.hidden=true;
}
function showError(error) {
  recover=error.message.startsWith('SIGN_IN_')||error.stage==='profile';
  signInAgain=Boolean(error.reauthenticationRequired)||(signInAgain&&error.message.startsWith('SIGN_IN_'));
  recoveryError=error;
  recoveryLoginHint=Office.context.mailbox.userProfile.emailAddress;
  apply.disabled=true;
  document.getElementById('signature').innerHTML='';
  connect.textContent=signInAgain?'Sign in again':['SIGN_IN_REQUIRED','SIGN_IN_CANCELLED'].includes(error.message)?'Continue with Microsoft 365':'Retry connection';
  const message=errorText(error);
  say(signInAgain?message.replace('Choose Continue with Microsoft 365','Choose Sign in again').replace('Choose Retry connection','Choose Sign in again'):message);
  const diagnostics=document.getElementById('diagnostics');
  const support=document.getElementById('support');
  if(diagnostics) diagnostics.textContent=supportDetails(error);
  if(support) support.hidden=false;
}
// Opening the panel may reuse Outlook's session, but must never launch a sign-in popup.
Office.onReady(()=>{
  // Mobile exposes the pane while reading an email, not in compose mode.
  // Never offer to modify the received message that opened this pane.
  apply.hidden=isMobileOutlook();
  return previewSignature(false);
});
async function previewSignature(interactive) {
  connect.disabled=true; apply.disabled=true;
  say('Loading your profile using your Outlook sign-in…');
  try {
    const loginHint=Office.context.mailbox.userProfile.emailAddress;
    if(loginHint!==recoveryLoginHint) {signInAgain=false;recoveryError=undefined;}
    // Reuse the already-loaded settings on a recovery click so a network fetch
    // cannot consume the browser's user-activation window before the popup.
    if(!interactive||!recover||!bundle) bundle=await fetchBundle(__SITE_URL__);
    if(!isConfigured(bundle.deployment)) throw new Error('ADMIN_SETUP_REQUIRED');
    const graph=await graphProfile(bundle.deployment,{interactive,recover:interactive&&recover,reauthenticate:interactive&&signInAgain,recoveryError,loginHint});
    const item=Office.context.mailbox.item;
    const sender=typeof item?.from?.getAsync==='function'?await officeCall(item.from,'getAsync'):{displayName:graph.displayName,emailAddress:Office.context.mailbox.userProfile.emailAddress};
    const profile=profileForSender(graph,sender,bundle.branding);
    document.getElementById('signature').innerHTML=renderSignature(bundle,profile,{images:'preview'});
    connect.textContent='Refresh preview';
    recover=false;
    signInAgain=false;recoveryError=undefined;recoveryLoginHint=undefined;
    apply.disabled=isMobileOutlook() || !bundle.enabled || !item?.body?.setSignatureAsync;
    say(!bundle.enabled?'Connected. Automatic insertion is paused by your administrator.':isMobileOutlook()?'Connected. Close this panel and start a new message, reply or forward. Your signature is added automatically; expand a quick reply to see it.':'Connected. Your signature uses the current company template.');
  }catch(error){showError(error);}
  finally{connect.disabled=false;}
}
connect.addEventListener('click',()=>{if(!connect.disabled) return previewSignature(true);});
apply.addEventListener('click',async()=>{
  if (isMobileOutlook()) return;
  apply.disabled=true;connect.disabled=true;
  try {
    bundle=await fetchBundle(__SITE_URL__);
    const result=await applySignature({item:Office.context.mailbox.item,bundle,getGraph:()=>graphProfile(bundle.deployment,{loginHint:Office.context.mailbox.userProfile.emailAddress})});
    if(result.status==='applied') Office.context.mailbox.item.notificationMessages?.removeAsync('ark-signature-status',()=>{});
    apply.disabled=!bundle.enabled;
    say(result.status==='applied'?'Your signature has been updated in this message.':'Automatic insertion is paused by your administrator.');
  }catch(error){showError(error);}
  finally{connect.disabled=false;}
});
