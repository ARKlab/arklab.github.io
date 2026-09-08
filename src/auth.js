import {createNestablePublicClientApplication,InteractionRequiredAuthError} from '@azure/msal-browser';
import {jsonRequest} from './network.js';
let appPromise;

export function isConfigured(config) {
  return /^[a-f0-9-]{36}$/i.test(config.clientId||'') && /^[a-f0-9-]{36}$/i.test(config.tenantId||'');
}

export async function graphProfile(config,{interactive=false,loginHint=''}={}) {
  if (!isConfigured(config)) throw new Error('ADMIN_SETUP_REQUIRED');
  if (!Office.context.requirements.isSetSupported('NestedAppAuth','1.1')) throw new Error('OUTLOOK_UPDATE_REQUIRED');
  if (!appPromise) appPromise=createNestablePublicClientApplication({auth:{clientId:config.clientId,authority:'https://login.microsoftonline.com/'+config.tenantId}}).catch(error=>{appPromise=undefined;throw error;});
  const app=await appPromise;
  const request={scopes:['User.Read'],...(loginHint?{loginHint}:{})};
  let token;
  try {token=await app.acquireTokenSilent(request);}
  catch(error) {
    if (interactive && error instanceof InteractionRequiredAuthError) token=await app.acquireTokenPopup(request);
    else throw new Error(error instanceof InteractionRequiredAuthError?'SIGN_IN_REQUIRED':'SIGN_IN_UNAVAILABLE');
  }
  if (token.account?.tenantId && token.account.tenantId.toLowerCase()!==config.tenantId.toLowerCase()) throw new Error('TENANT_MISMATCH');
  // Only the signed-in person's profile; no directory-wide or mail permissions.
  return jsonRequest('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,officeLocation',{
    headers:{Authorization:'Bearer '+token.accessToken},cache:'no-store',credentials:'omit'
  });
}
