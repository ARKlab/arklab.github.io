import {createNestablePublicClientApplication,InteractionRequiredAuthError} from '@azure/msal-browser';
import {jsonRequest} from './network.js';
import {signatureError} from './errors.js';
let appPromise;
let pendingAuth;

function timed(operation,timeoutMs) {
  let timer;
  // MSAL cannot cancel the Outlook broker. Ignore late results after this rejects.
  return Promise.race([operation,new Promise((_,reject)=>{
    timer=setTimeout(()=>reject(signatureError('SIGN_IN_TIMEOUT','sign-in')),timeoutMs);
  })]).finally(()=>clearTimeout(timer));
}

function authError(source) {
  const error=signatureError('SIGN_IN_UNAVAILABLE','sign-in',source);
  if(error.microsoftCode==='7000024') error.message='SIGN_IN_BROKER_REJECTED';
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

export async function graphProfile(config,{interactive=false,loginHint='',recover=false}={}) {
  if (!isConfigured(config)) throw new Error('ADMIN_SETUP_REQUIRED');
  if (!Office.context.requirements.isSetSupported('NestedAppAuth','1.1')) throw new Error('OUTLOOK_UPDATE_REQUIRED');
  if (!appPromise) appPromise=timed(createNestablePublicClientApplication({auth:{clientId:config.clientId,authority:'https://login.microsoftonline.com/'+config.tenantId}}),10000).catch(error=>{appPromise=undefined;throw authError(error);});
  const app=await appPromise;
  const request={scopes:['User.Read'],...(loginHint?{loginHint}:{})};
  async function acquire(forceRefresh) {
    try { return await brokerCall(app,'acquireTokenSilent',{...request,...(forceRefresh?{forceRefresh:true}:{})},15000); }
    catch(error) {
      // Never prompt during background insertion or retry arbitrary broker errors in a popup.
      if(interactive && error instanceof InteractionRequiredAuthError && authError(error).message==='SIGN_IN_REQUIRED') {
        try { return await brokerCall(app,'acquireTokenPopup',request,120000); }
        catch(popupError) { throw authError(popupError); }
      }
      throw authError(error);
    }
  }
  for(let attempt=0;attempt<2;attempt++) {
    const token=await acquire(recover||attempt===1);
    if (token.account?.tenantId && token.account.tenantId.toLowerCase()!==config.tenantId.toLowerCase()) throw new Error('TENANT_MISMATCH');
    try {
      // Only the signed-in person's profile; no directory-wide or mail permissions.
      return await jsonRequest('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,officeLocation',{
        headers:{Authorization:'Bearer '+token.accessToken},cache:'no-store',credentials:'omit'
      });
    }catch(error) {
      // A rejected cached access token gets exactly one fresh-token retry.
      if(error.message==='REQUEST_FAILED_401' && attempt===0) continue;
      throw signatureError(error.message,'profile');
    }
  }
}
