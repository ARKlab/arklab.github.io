import {profileForSender,renderSignature,plainSignature} from './render.js';

export function officeCall(object, method, ...args) {
  return new Promise((resolve,reject)=>{
    if (typeof object?.[method]!=='function') return reject(new Error('OUTLOOK_UPDATE_REQUIRED'));
    object[method](...args,result=>result.status==='succeeded'?resolve(result.value):reject(new Error('OUTLOOK_'+(result.error?.code||'FAILED'))));
  });
}

function checkActive(context) { if (context.cancelled) throw new Error('CANCELLED'); }

export async function applySignature({item,bundle,getGraph,context={cancelled:false}}) {
  if (!bundle.enabled) return {status:'paused'};
  const sender=await officeCall(item.from,'getAsync');
  const graph=await getGraph();
  checkActive(context);
  const profile=profileForSender(graph,sender,bundle.branding);
  const compose=await officeCall(item,'getComposeTypeAsync');
  const compact=bundle.branding.compactReplies && ['reply','forward'].includes(compose.composeType);
  const bodyType=await officeCall(item.body,'getTypeAsync');
  const isText=String(bodyType).toLowerCase()==='text';
  // Prepare/validate before adding any attachments.
  const signature=isText?plainSignature(bundle,profile,compact):renderSignature(bundle,profile,{compact,officeCss:true});
  checkActive(context);
  let existing=[];
  if (!isText && !compact) {
    existing=await officeCall(item,'getAttachmentsAsync');
    for (const asset of Object.values(bundle.assets)) {
      if (!signature.includes('cid:'+asset.filename)) continue;
      checkActive(context);
      if (!existing.some(a=>a.name===asset.filename && a.isInline)) {
        await officeCall(item,'addFileAttachmentFromBase64Async',asset.base64,asset.filename,{isInline:true});
      }
    }
  }
  checkActive(context);
  const latestSender=await officeCall(item.from,'getAsync');
  if (latestSender.emailAddress.toLowerCase()!==sender.emailAddress.toLowerCase()) throw new Error('SENDER_CHANGED');
  checkActive(context);
  // Only replace the signature slot. Never replace the message body or send mail.
  await officeCall(item.body,'setSignatureAsync',signature,{coercionType:isText?'text':'html'});
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
