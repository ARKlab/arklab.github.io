// This is rollout targeting, not access control. The source and logo files are public.
// Keep the pilot account's address out of the static configuration.
const targetDigest='60e2cda6273e570f6f94f568358d16afd50331ca18f5baa29298e5d261661ac0';
export async function isLogoPilotAccount(email) {
  if(typeof email!=='string'||typeof crypto==='undefined'||!crypto.subtle) return false;
  const input=new TextEncoder().encode('ark-signature-logo-trial:'+email.trim().toLowerCase());
  const digest=await crypto.subtle.digest('SHA-256',input);
  return [...new Uint8Array(digest)].map(v=>v.toString(16).padStart(2,'0')).join('')===targetDigest;
}

const dimensions={ark:[92,32],artesian:[64,20]};
function imagePair(assets,key) {
  const [width,height]=dimensions[key];
  const alt=key==='ark'?'ARK':'Artesian';
  return `<span class="arklogotrial-light"><img src="cid:${assets[key+'Fallback'].filename}" alt="${alt}" width="${width}" height="${height}" border="0" style="display:inline-block;width:${width}px;height:${height}px;vertical-align:middle;border:0;"></span><!--[if !mso]><!--><span class="arklogotrial-dark-${key}"><img src="cid:${assets[key+'White'].filename}" alt="${alt}" width="0" height="0" border="0" style="display:none;width:0;height:0;max-height:0;overflow:hidden;border:0;font-size:0;line-height:0;"></span><!--<![endif]-->`;
}

function darkRules(prefix='') {
  return `${prefix}.arklogotrial-light{display:none!important;}`+Object.entries(dimensions).map(([key,[w,h]])=>`${prefix}.arklogotrial-dark-${key} img{display:inline-block!important;width:${w}px!important;height:${h}px!important;max-height:${h}px!important;vertical-align:middle!important;}`).join('');
}

// These selectors affect only the pilot logos, including copies in quoted replies.
// Default hiding also lives in inline styles and zero HTML image dimensions.
export const logoSwitchCss=`<style type="text/css">@media (prefers-color-scheme:dark){${darkRules()}\n}\n${darkRules('[data-ogsc] ')}${darkRules('[data-ogsb] ')}</style>`;

export function makeLogoPilotBundle(base,assets) {
  if(!assets||Object.keys(assets).length!==4) throw new Error('LOGO_PILOT_INVALID_ASSETS');
  for(const key of ['arkFallback','arkWhite','artesianFallback','artesianWhite']) {
    const asset=assets?.[key];
    if(!asset||!/^ark-signatures-trial-[a-z0-9-]+\.png$/.test(asset.filename)||typeof asset.base64!=='string'||asset.base64.length>134000||!/^iVBORw0KGgo[A-Za-z0-9+/]*={0,2}$/.test(asset.base64)) throw new Error('LOGO_PILOT_INVALID_ASSETS');
  }
  const template=base.templates.full;
  const arkImage=/<img\s+src="{{ARK_LOGO}}"[^>]*>/g;
  if((template.match(arkImage)||[]).length!==1||template.split('{{ARTESIAN_LINK}}').length!==2||!base.branding.artesianWordmark||!base.branding.artesianWebsite) throw new Error('LOGO_PILOT_TEMPLATE_CHANGED');
  const full=logoSwitchCss+template.replace(arkImage,imagePair(assets,'ark'))
    .replace('{{ARTESIAN_LINK}}',`<a href="{{ARTESIAN_URL}}" style="text-decoration:none;">${imagePair(assets,'artesian')}</a>`)
    .replace('width="90" valign="middle" style="width:90px;','width="92" valign="middle" style="width:92px;');
  return {...base,revision:base.revision+'-logo-pilot',assets:{...base.assets,...assets},templates:{...base.templates,full}};
}
