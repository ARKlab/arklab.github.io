import {createNestablePublicClientApplication,InteractionRequiredAuthError} from '@azure/msal-browser';
import {jsonRequest} from './network.js';
import {signatureError} from './errors.js';
import {remainingTime} from './deadline.js';
import {claimsFor,copyClaims} from './graph-error.js';
let appPromise;
let pendingAuth;
const recoveryContexts=new WeakMap();

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

export async function graphProfile(config,{interactive=false,loginHint='',recover=false,reauthenticate=false,recoveryError,context={}}={}) {
  if (!isConfigured(config)) throw new Error('ADMIN_SETUP_REQUIRED');
  if (!Office.context.requirements.isSetSupported('NestedAppAuth','1.1')) throw new Error('OUTLOOK_UPDATE_REQUIRED');
  remainingTime(context,10000,13000);
  if (!appPromise) appPromise=timed(createNestablePublicClientApplication({auth:{clientId:config.clientId,authority:'https://login.microsoftonline.com/'+config.tenantId}}),10000).catch(error=>{appPromise=undefined;throw authError(error);});
  const app=await timed(appPromise,remainingTime(context,10000,13000));
  const request={scopes:['User.Read'],...(loginHint?{loginHint}:{})};
  // A challenge belongs to the mailbox/app that received it, not a later account.
  const recoveryKey=[config.clientId,config.tenantId,loginHint.toLowerCase()].join('|');
  let challengeError=recoveryError&&recoveryContexts.get(recoveryError)===recoveryKey?recoveryError:undefined;
  let attempts=0;
  let profileAttempts=0;
  let interactionAttempted=false;
  const explicitSignIn=interactive&&reauthenticate;
  function failed(error) {
    error.attempts=attempts;
    error.profileAttempts=profileAttempts;
    error.authMode=interactionAttempted?'interactive':'silent';
    if(challengeError) {
      copyClaims(challengeError,error);
      error.claimsChallenge=challengeError.claimsChallenge;
    }
    recoveryContexts.set(error,recoveryKey);
    return error;
  }
  async function popup(correlationId) {
    interactionAttempted=true;
    const claims=claimsFor(challengeError);
    try {
      return await brokerCall(app,'acquireTokenPopup',{...request,correlationId,...(claims?{claims}:{})},remainingTime(context,120000,13000));
    }catch(error) {throw authError(error,correlationId);}
  }
  async function acquire(forceRefresh) {
    const timeoutMs=remainingTime(context,15000,13000);
    const correlationId=typeof crypto!=='undefined'&&typeof crypto.randomUUID==='function'?crypto.randomUUID():undefined;
    const claims=claimsFor(challengeError);
    const tokenRequest={...request,...(correlationId?{correlationId}:{}),...(forceRefresh?{forceRefresh:true}:{}),...(claims?{claims}:{})};
    attempts++;
    if(explicitSignIn) return popup(correlationId);
    try { return await brokerCall(app,'acquireTokenSilent',tokenRequest,timeoutMs); }
    catch(error) {
      // Never prompt during background insertion or retry arbitrary broker errors in a popup.
      if(interactive && !interactionAttempted && error instanceof InteractionRequiredAuthError && authError(error).message==='SIGN_IN_REQUIRED') {
        return popup(correlationId);
      }
      throw authError(error,correlationId);
    }
  }
  for(let attempt=0;attempt<2;attempt++) {
    let token;
    try { token=await acquire(recover||attempt===1); }
    catch(error) {
      failed(error);
      const retryable=error.message==='SIGN_IN_BROKER_REJECTED'||(error.message==='SIGN_IN_UNAVAILABLE'&&['temporarily_unavailable','server_error','bridge_connection_reset','no_network_connectivity'].includes(error.msalCode));
      // Only retry completed, identified failures. Never duplicate an outstanding
      // broker call, or silently loop over MFA, consent, policy or unknown errors.
      if(attempt!==0 || pendingAuth || interactionAttempted || !retryable) throw error;
      if(remainingTime(context,2000,13000)<2000) throw error;
      await new Promise(resolve=>setTimeout(resolve,1000));
      remainingTime(context,15000,13000);
      continue;
    }
    remainingTime(context,8000,5000);
    if (token.account?.tenantId && token.account.tenantId.toLowerCase()!==config.tenantId.toLowerCase()) throw new Error('TENANT_MISMATCH');
    const clientRequestId=typeof crypto!=='undefined'&&typeof crypto.randomUUID==='function'?crypto.randomUUID():undefined;
    try {
      // Only the signed-in person's profile; no directory-wide or mail permissions.
      const profile=await jsonRequest('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,proxyAddresses,jobTitle,businessPhones,mobilePhone,officeLocation',{
        headers:{Authorization:'Bearer '+token.accessToken,...(clientRequestId?{'client-request-id':clientRequestId,'return-client-request-id':'true'}:{})},cache:'no-store',credentials:'omit'
      },remainingTime(context,8000,5000));
      remainingTime(context,5000);
      return profile;
    }catch(error) {
      profileAttempts++;
      // Keep the latest valid challenge through silent/interactive recovery.
      if(claimsFor(error)) challengeError=error;
      // A rejected cached access token gets one fresh-token retry. An explicit
      // sign-in gets one profile request, never a second popup or retry loop.
      if(error.message==='REQUEST_FAILED_401' && attempt===0 && !interactionAttempted) continue;
      const failure=copyClaims(error,signatureError(error.message,'profile',{...error,clientRequestId}));
      if(error.message==='REQUEST_FAILED_401') failure.reauthenticationRequired=true;
      throw failed(failure);
    }
  }
}
