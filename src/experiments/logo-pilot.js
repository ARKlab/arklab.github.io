import {isLogoPilotAccount,makeLogoPilotBundle} from './logo-colours.js';
import {officeCall,applySignature} from '../office-flow.js';

export async function runLogoPilot({mailbox,fetchBase,fetchAssets,getGraph,restore=false}) {
  const email=mailbox.userProfile?.emailAddress;
  if(!await isLogoPilotAccount(email)) throw new Error('LOGO_PILOT_NOT_ASSIGNED');
  const item=mailbox.item;
  if(!item?.body?.setSignatureAsync) throw new Error('LOGO_PILOT_NEW_HTML_ONLY');
  const expected=email.trim().toLowerCase();
  async function checkSender() {
    if(mailbox.item!==item) throw new Error('SENDER_CHANGED');
    const sender=await officeCall(item.from,'getAsync');
    if(String(sender.emailAddress).toLowerCase()!==expected) throw new Error('LOGO_PILOT_OWN_SENDER_ONLY');
  }
  await checkSender();
  const compose=await officeCall(item,'getComposeTypeAsync');
  const type=await officeCall(item.body,'getTypeAsync');
  if(compose.composeType!=='newMail'||String(type).toLowerCase()!=='html') throw new Error('LOGO_PILOT_NEW_HTML_ONLY');
  const base=await fetchBase();
  if(!base.enabled) return {status:'paused'};
  const bundle=restore?base:makeLogoPilotBundle(base,await fetchAssets());
  await checkSender();
  return applySignature({item,bundle,getGraph:async()=>{
    const graph=await getGraph(base.deployment,{loginHint:email,interactive:false});
    const identities=[graph.mail,graph.userPrincipalName].filter(Boolean).map(value=>value.toLowerCase());
    if(!identities.includes(expected)) throw new Error('LOGO_PILOT_PROFILE_MISMATCH');
    await checkSender();
    return graph;
  }});
}

const messages={
  LOGO_PILOT_NOT_ASSIGNED:'This logo test is not assigned to the current account.',
  LOGO_PILOT_OWN_SENDER_ONLY:'Choose your own ARK address in From to try these logos.',
  LOGO_PILOT_PROFILE_MISMATCH:'The signed-in profile does not match this mailbox. Reconnect using your ARK account.',
  LOGO_PILOT_NEW_HTML_ONLY:'Use a new HTML email for the logo test. Replies and forwards keep the compact signature.',
  LOGO_PILOT_INVALID_ASSETS:'The test artwork could not be loaded. Your current signature has been kept.',
  LOGO_PILOT_TEMPLATE_CHANGED:'The company template has changed. The logo test needs an update before it can be used.',
  SENDER_CHANGED:'The sending account changed. Check From and try again.'
};

// Called after the normal preview has finished. Other accounts get no extra UI,
// artwork download, Graph request, automatic insertion or stored preference.
export async function installLogoPilot({document,mailbox,fetchBase,fetchAssets,getGraph}) {
  try {
    if(!await isLogoPilotAccount(mailbox.userProfile?.emailAddress)||!mailbox.item?.body?.setSignatureAsync) return false;
    if(document.getElementById('logo-pilot')) return true;
    const section=document.createElement('section');
    section.id='logo-pilot';
    section.setAttribute('aria-label','Logo colour test');
    section.style.cssText='border-top:1px solid #CED4CE;margin-top:20px;padding-top:18px;';
    const heading=document.createElement('h2');heading.textContent='Logo colour test';
    const note=document.createElement('p');note.className='small';note.textContent='After your normal signature appears, insert the test logos in this draft. New messages will keep using the standard design.';
    const insert=document.createElement('button');insert.id='insert-test-logos';insert.textContent='Insert test logos';
    const restore=document.createElement('button');restore.id='restore-standard-signature';restore.className='secondary';restore.textContent='Restore standard signature';
    const status=document.createElement('p');status.id='logo-pilot-status';status.className='small';status.setAttribute('role','status');
    section.append(heading,note,insert,restore,status);
    const preview=document.getElementById('signature');
    if(!preview?.parentNode) return false;
    preview.parentNode.insertBefore(section,preview);
    let busy=false;
    async function run(restoreStandard) {
      if(busy) return;
      const normalControls=['connect','apply'].map(id=>document.getElementById(id)).filter(Boolean);
      if(normalControls.some(control=>control.disabled)) {status.textContent='Wait for the profile to finish loading, or use Retry connection first.';return;}
      busy=true;insert.disabled=true;restore.disabled=true;
      const previous=normalControls.map(control=>control.disabled);
      normalControls.forEach(control=>{control.disabled=true;});
      status.textContent=restoreStandard?'Restoring the standard signature…':'Inserting test logos…';
      try {
        const result=await runLogoPilot({mailbox,fetchBase,fetchAssets,getGraph,restore:restoreStandard});
        status.textContent=result.status==='paused'?'Signature insertion is paused by your administrator.':restoreStandard?'Standard signature restored in this draft.':'Test logos inserted in this draft. Send a test email, then compare the received message in light and dark mode.';
      }catch(error){status.textContent=messages[error.message]||'The test could not complete. Use Refresh preview to check your connection, then try again.';}
      finally {
        normalControls.forEach((control,index)=>{control.disabled=previous[index];});
        busy=false;insert.disabled=false;restore.disabled=false;
      }
    }
    insert.addEventListener('click',()=>run(false));
    restore.addEventListener('click',()=>run(true));
    return true;
  }catch{return false;} // An optional experiment must not disrupt the released panel.
}
