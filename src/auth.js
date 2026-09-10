import {createNestablePublicClientApplication,InteractionRequiredAuthError} from '@azure/msal-browser';
import {jsonRequest} from './network.js';
import {signatureError} from './errors.js';
import {remainingTime} from './deadline.js';
let appPromise;
let pendingAuth;

function timed(operation,timeoutMs) {
  let timer;
  // MSAL cannot cancel the Outlook broker. Ignore late results after this rejects.
  return Promise.race([operation,new Promise((_,reject)=>{
    timer=setTimeout(()=>reject(signatureError('SIGN_IN_TIMEOUT','sign-in')),timeoutMs);
  })]).finally(()=>clearTimeout(timer));
}

function authError(source,correlationId) {
  if(source.message==='REQUEST_TIMEOUT') return signatureError('REQUEST_TIMEOUT','sign-in');
  const error=signatureError('SIGN_IN_UNAVAILABLE','sign-in',source);
  if(!error.correlationId && correlationId) error.correlationId=correlationId;
  if(error.microsoftCode==='7000024'||/inconsistent broker application ids asserted by incoming credentials/i.test(String(source.errorMessage||source.message||''))) error.message='SIGN_IN_BROKER_REJECTED';
  else if(source instanceof InteractionRequiredAuthError) error.message='SIGN_IN_REQUIRED';
  else if(source.message==='SIGN_IN_TIMEOUT'||['timed_out','monitor_window_timeout','bridge_timeout'].includes(source.errorCode)) error.message='SIGN_IN_TIMEOUT';
  else if(source.errorCode==='user_cancelled') error.message='SIGN_IN_CANCELLED';
  else if(source.message==='SIGN_IN_PENDING') error.message='SIGN_IN_PENDING';
  return error;
}

function brokerCall(app,method,request,timeoutMs) {
  // A UI deadline cannot cancel native authentication. Do not pile new requests
  // onto a still-pending broker operation when the user chooses Retry.
  if(pendingAuth) throw signatureError('SIGN_IN_PENDING','sign-in');
  pendingAuth=Promise.resolve().then(()=>app[method](request)).finally(()=>{pendingAuth=undefined;});
  return timed(pendingAuth,timeoutMs);
}

export function isConfigured(config) {
  return /^[a-f0-9-]{36}$/i.test(config.clientId||'') && /^[a-f0-9-]{36}$/i.test(config.tenantId||'');
}

export async function graphProfile(config,{interactive=false,loginHint='',recover=false,context={}}={}) {
  if (!isConfigured(config)) throw new Error('ADMIN_SETUP_REQUIRED');
  if (!Office.context.requirements.isSetSupported('NestedAppAuth','1.1')) throw new Error('OUTLOOK_UPDATE_REQUIRED');
  remainingTime(context,10000,13000);
  if (!appPromise) appPromise=timed(createNestablePublicClientApplication({auth:{clientId:config.clientId,authority:'https://login.microsoftonline.com/'+config.tenantId}}),10000).catch(error=>{appPromise=undefined;throw authError(error);});
  const app=await timed(appPromise,remainingTime(context,10000,13000));
  const request={scopes:['User.Read'],...(loginHint?{loginHint}:{})};
  let attempts=0;
  async function acquire(forceRefresh) {
    const timeoutMs=remainingTime(context,15000,13000);
    const correlationId=typeof crypto!=='undefined'&&typeof crypto.randomUUID==='function'?crypto.randomUUID():undefined;
    const tokenRequest={...request,...(correlationId?{correlationId}:{}),...(forceRefresh?{forceRefresh:true}:{})};
    attempts++;
    try { return await brokerCall(app,'acquireTokenSilent',tokenRequest,timeoutMs); }
    catch(error) {
      // Never prompt during background insertion or retry arbitrary broker errors in a popup.
      if(interactive && error instanceof InteractionRequiredAuthError && authError(error).message==='SIGN_IN_REQUIRED') {
        try { return await brokerCall(app,'acquireTokenPopup',request,remainingTime(context,120000,13000)); }
        catch(popupError) {
          const failure=authError(popupError);
          failure.interactionAttempted=true;
          throw failure;
        }
      }
      throw authError(error,correlationId);
    }
  }
  for(let attempt=0;attempt<2;attempt++) {
    let token;
    try { token=await acquire(recover||attempt===1); }
    catch(error) {
      error.attempts=attempts;
      const retryable=error.message==='SIGN_IN_BROKER_REJECTED'||(error.message==='SIGN_IN_UNAVAILABLE'&&['temporarily_unavailable','server_error','bridge_connection_reset','no_network_connectivity'].includes(error.msalCode));
      // Only retry completed, identified failures. Never duplicate an outstanding
      // broker call, or silently loop over MFA, consent, policy or unknown errors.
      if(attempt!==0 || pendingAuth || error.interactionAttempted || !retryable) throw error;
      if(remainingTime(context,2000,13000)<2000) throw error;
      await new Promise(resolve=>setTimeout(resolve,1000));
      remainingTime(context,15000,13000);
      continue;
    }
    remainingTime(context,8000,5000);
    if (token.account?.tenantId && token.account.tenantId.toLowerCase()!==config.tenantId.toLowerCase()) throw new Error('TENANT_MISMATCH');
    try {
      // Only the signed-in person's profile; no directory-wide or mail permissions.
      const profile=await jsonRequest('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,officeLocation',{
        headers:{Authorization:'Bearer '+token.accessToken},cache:'no-store',credentials:'omit'
      },remainingTime(context,8000,5000));
      remainingTime(context,5000);
      return profile;
    }catch(error) {
      // A rejected cached access token gets exactly one fresh-token retry.
      if(error.message==='REQUEST_FAILED_401' && attempt===0) continue;
      const failure=signatureError(error.message,'profile');
      failure.attempts=attempts;
      throw failure;
    }
  }
}
