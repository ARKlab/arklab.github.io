import {fetchBundle} from './network.js';
import {graphProfile} from './auth.js';
import {applySignature,completeEvent} from './office-flow.js';

const siteUrl=__SITE_URL__;
function notify(message) {
  const notifications=Office.context.mailbox.item?.notificationMessages;
  if (!notifications) return;
  notifications.replaceAsync('ark-signature-status',{
    type:'informationalMessage',message,icon:'Icon16x16',persistent:false
  },()=>{});
}

function messageFor(error) {
  if (error.message==='SIGN_IN_REQUIRED') return 'Open ARK signatures and choose Continue with Microsoft 365 to allow access to your profile.';
  if (error.message==='OUTLOOK_UPDATE_REQUIRED') return 'Your Outlook needs an update to use ARK signatures. Your existing signature has been kept.';
  if (error.message==='UNAPPROVED_SENDER') return 'ARK signatures is not configured for this sending address. Check the signature before sending.';
  if (error.message==='SIGN_IN_BROKER_REJECTED') return 'Microsoft rejected the Outlook sign-in (7000024). Open ARK signatures to retry; if it persists, quit and reopen Outlook.';
  if (error.message==='SIGN_IN_TIMEOUT') return 'The Outlook sign-in timed out. Your existing signature has been kept. Open ARK signatures to retry.';
  if (error.message==='SIGN_IN_UNAVAILABLE') return 'Outlook could not sign in to Microsoft 365. Your existing signature has been kept. Open ARK signatures to retry.';
  if (error.message==='REQUEST_TIMEOUT') return 'ARK signatures reached its time limit. Your existing signature has been kept. Open ARK signatures to retry.';
  return 'ARK signatures could not refresh. Your existing signature has been kept; use ARK signatures to retry.';
}

function onCompose(event) {
  return completeEvent(event,async context=>{
    const bundle=await fetchBundle(siteUrl);
    if (context.cancelled) return;
    const result=await applySignature({item:Office.context.mailbox.item,bundle,context,getGraph:()=>graphProfile(bundle.deployment,{loginHint:Office.context.mailbox.userProfile.emailAddress})});
    if (context.cancelled) return;
    if(result.status==='applied') {
      Office.context.mailbox.item.notificationMessages.removeAsync('ark-signature-status',()=>{});
      if (!result.directoryMatched) notify('This sending address uses name and email only. Ask your administrator to confirm its directory details.');
    }
  },{onError:error=>notify(messageFor(error))});
}
Office.onReady(()=>{});
Office.actions.associate('arkOnCompose',onCompose);
Office.actions.associate('arkOnFromChanged',onCompose);
