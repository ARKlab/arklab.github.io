const msalCodes=new Set(['temporarily_unavailable','server_error','bridge_connection_reset','bridge_timeout','bridge_handshake_failed','bridge_response_invalid','no_network_connectivity','no_account_error','no_account_found','nested_app_auth_bridge_disabled','unknown_error','invalid_grant','interaction_required','login_required','consent_required','user_cancelled','timed_out','monitor_window_timeout']);
const subCodes=new Set(['basic_action','additional_action','message_only','consent_required','user_password_expired','bad_token','token_expired','protection_policy_required']);
// Retain only support references, never raw provider messages, tokens or profiles.
export function signatureError(code,stage,source={}) {
  const error=new Error(code);
  error.stage=stage;
  error.at=new Date().toISOString();
  const aadsts=String(source.errorMessage||source.message||'').match(/\bAADSTS(\d{4,10})\b/);
  const numeric=Array.isArray(source.errorCodes)?source.errorCodes.find(c=>/^\d{4,10}$/.test(String(c))):undefined;
  const singular=String(source.errorCode||'').match(/^(?:AADSTS)?(\d{4,10})$/);
  const providerCode=aadsts?.[1] || numeric || source.microsoftCode || singular?.[1];
  if (/^\d{4,10}$/.test(String(providerCode))) error.microsoftCode=String(providerCode);
  if (/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i.test(source.correlationId||'')) error.correlationId=source.correlationId;
  if(msalCodes.has(source.errorCode||source.msalCode)) error.msalCode=source.errorCode||source.msalCode;
  if(subCodes.has(source.subError)) error.subError=source.subError;
  return error;
}

export function errorText(error) {
  const messages={
    ADMIN_SETUP_REQUIRED:'Your administrator still needs to connect Microsoft 365.',
    OUTLOOK_UPDATE_REQUIRED:'Update Outlook to a version that supports this add-in.',
    SIGN_IN_REQUIRED:'Microsoft needs your permission or a sign-in check before your profile can be loaded. Choose Continue with Microsoft 365.',
    SIGN_IN_UNAVAILABLE:'Outlook could not complete your Microsoft 365 sign-in. Choose Retry connection. If it fails again, quit and reopen Outlook.',
    SIGN_IN_BROKER_REJECTED:'Microsoft rejected the sign-in supplied by Outlook. Choose Retry connection once. If it fails again, fully quit and reopen Outlook, then open ARK signatures.',
    SIGN_IN_TIMEOUT:'Outlook did not finish signing in within the time limit. Choose Retry connection. If it stalls again, fully quit and reopen Outlook.',
    SIGN_IN_PENDING:'The earlier Outlook sign-in is still pending. Wait briefly and retry, or fully quit and reopen Outlook if it remains stuck.',
    SIGN_IN_CANCELLED:'The Microsoft sign-in was cancelled. Choose Continue with Microsoft 365 when you are ready.',
    TENANT_MISMATCH:'The signed-in Microsoft account does not belong to this organisation. Check the account selected in Outlook.',
    UNAPPROVED_SENDER:'This sending address is not configured for ARK signatures.',
    SENDER_CHANGED:'The sending address changed. Refresh the preview before trying again.',
    REQUEST_TIMEOUT:'The request took too long. Check your connection and try again.',
  };
  if(messages[error.message]) return messages[error.message];
  if(error.stage==='settings') return 'Could not load the company signature settings. Check your connection and choose Retry connection.';
  if(error.stage==='profile') return 'Microsoft 365 could not return your profile. Choose Retry connection. If it fails again, share the support details with IT.';
  if(/^OUTLOOK_/.test(error.message)) return 'Outlook could not update this message. Try again in a new draft.';
  return 'Could not refresh your signature. Please try again or contact your administrator.';
}

export function supportDetails(error) {
  // Error.message from arbitrary libraries is deliberately excluded.
  const lines=['ARK signatures '+(typeof __APP_VERSION__==='string'?__APP_VERSION__:'development'),'Time (UTC): '+(error.at||new Date().toISOString())];
  if(['sign-in','settings','profile'].includes(error.stage)) lines.push('Step: '+error.stage);
  if(/^\d{4,10}$/.test(error.microsoftCode||'')) lines.push('Microsoft code: AADSTS'+error.microsoftCode);
  if(/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i.test(error.correlationId||'')) lines.push('Correlation ID: '+error.correlationId);
  if(msalCodes.has(error.msalCode)) lines.push('Sign-in code: '+error.msalCode);
  if(subCodes.has(error.subError)) lines.push('Sign-in detail: '+error.subError);
  if(error.attempts===1||error.attempts===2) lines.push('Token attempts: '+error.attempts);
  if(/^(SIGN_IN_[A-Z_]+|REQUEST_(TIMEOUT|NETWORK_ERROR|INVALID_RESPONSE|FAILED_\d{3}))$/.test(error.message)) lines.push('Result: '+error.message);
  return lines.join('\n');
}
