import {profileForSender,renderSignature,plainSignature} from './render.js';
import {isMobileOutlook} from './platform.js';
import {signatureError} from './errors.js';

export function officeCall(object, method, ...args) {
  return new Promise((resolve,reject)=>{
    if (typeof object?.[method]!=='function') return reject(new Error('OUTLOOK_UPDATE_REQUIRED'));
    object[method](...args,result=>result.status==='succeeded'?resolve(result.value):reject(new Error('OUTLOOK_'+(result.error?.code||'FAILED'))));
  });
}

async function officeStep(stage,object,method,...args) {
  try {return await officeCall(object,method,...args);}
  catch(error) {throw signatureError(error.message,stage);}
}

async function mobileImageAttached(item,key) {
  try {return await officeStep('session-read',item.sessionData,'getAsync',key)==='added';}
  catch(error) {
    // A fresh compose session has no marker. Outlook's KeyNotFound (9050)
    // means this image has not been added; other failures must still stop us.
    if(error.message==='OUTLOOK_9050') return false;
    throw error;
  }
}

function checkActive(context) { if (context.cancelled) throw new Error('CANCELLED'); }

// Compose and From-change events can overlap in the same mobile runtime.
// Keep each item's upload/marker/write sequence ordered; other drafts proceed
// independently. This is not a lock across separate host runtimes.
const mobileInsertions=new WeakMap();
export async function applySignature({item,bundle,getGraph,context={cancelled:false},mobile=isMobileOutlook()}) {
  const insert=()=>insertSignature({item,bundle,getGraph,context,mobile});
  if (!mobile) return insert();
  const previous=mobileInsertions.get(item)||Promise.resolve();
  const pending=previous.catch(()=>{}).then(()=>{checkActive(context);return insert();});
  mobileInsertions.set(item,pending);
  try {return await pending;}
  finally {if(mobileInsertions.get(item)===pending) mobileInsertions.delete(item);}
}

async function insertSignature({item,bundle,getGraph,context,mobile}) {
  if (!bundle.enabled) return {status:'paused'};
  const sender=await officeStep('sender-read',item.from,'getAsync');
  const graph=await getGraph();
  checkActive(context);
  const profile=profileForSender(graph,sender,bundle.branding);
  const compose=await officeStep('compose-type',item,'getComposeTypeAsync');
  const compact=bundle.branding.compactReplies && ['reply','forward'].includes(compose.composeType);
  // Mobile supports HTML signatures, but not body.getTypeAsync.
  const bodyType=mobile?'html':await officeStep('body-type',item.body,'getTypeAsync');
  const isText=String(bodyType).toLowerCase()==='text';
  // Prepare/validate before adding any attachments.
  const signature=isText?plainSignature(bundle,profile,compact):renderSignature(bundle,profile,{compact,officeCss:true});
  checkActive(context);
  let existing=[];
  if (!isText && !compact) {
    if (mobile) {
      // Mobile cannot enumerate attachments. Per-item session markers survive
      // successive compose/From-change events without storing any user data.
      if (typeof item.sessionData?.getAsync!=='function' || typeof item.sessionData?.setAsync!=='function') throw new Error('OUTLOOK_UPDATE_REQUIRED');
    } else existing=await officeStep('attachment-list',item,'getAttachmentsAsync');
    for (const asset of Object.values(bundle.assets)) {
      if (!signature.includes('cid:'+asset.filename)) continue;
      checkActive(context);
      const key='ark-signature-image:'+asset.filename;
      const attached=mobile ? await mobileImageAttached(item,key) : existing.some(a=>a.name===asset.filename && a.isInline);
      checkActive(context);
      if (!attached) {
        await officeStep('attachment-upload',item,'addFileAttachmentFromBase64Async',asset.base64,asset.filename,{isInline:true});
        // Record an upload that already succeeded even if our deadline expired
        // while it was in flight. Never start another upload or signature write
        // after cancellation. A host shutdown can still interrupt this marker.
        if (mobile) await officeStep('session-write',item.sessionData,'setAsync',key,'added');
        checkActive(context);
      }
    }
  }
  checkActive(context);
  const latestSender=await officeStep('sender-check',item.from,'getAsync');
  if (latestSender.emailAddress.toLowerCase()!==sender.emailAddress.toLowerCase()) throw new Error('SENDER_CHANGED');
  checkActive(context);
  // Only replace the signature slot. Never replace the message body or send mail.
  await officeStep('signature-write',item.body,'setSignatureAsync',signature,{coercionType:isText?'text':'html'});
  return {status:'applied',revision:bundle.revision,compact,directoryMatched:profile.directoryMatched};
}

export async function completeEvent(event,work,{timeoutMs=60000,onError=()=>{}}={}) {
  let finished=false;
  const context={cancelled:false,deadlineAt:Date.now()+timeoutMs};
  const finish=()=>{if(!finished){finished=true;event.completed();}};
  const timer=setTimeout(()=>{context.cancelled=true;try{onError(new Error('REQUEST_TIMEOUT'));}finally{finish();}},timeoutMs);
  try { await work(context); }
  catch(error) { if(!context.cancelled) onError(error); }
  finally {clearTimeout(timer);finish();}
}
