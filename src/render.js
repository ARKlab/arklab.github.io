export function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

export function emailAddress(value) {
  const email = String(value ?? '').trim();
  if (email.length > 254 || !/^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9](?:[A-Z0-9.-]*[A-Z0-9])?\.[A-Z]{2,}$/i.test(email)) throw new Error('INVALID_SENDER');
  return email;
}

export function httpsUrl(value, optional = false) {
  if (!value && optional) return '';
  const url = new URL(value);
  if (url.protocol !== 'https:' || url.username || url.password) throw new Error('INVALID_HTTPS_URL');
  return url.href;
}

function text(value, max = 180) { return String(value ?? '').replace(/[\u0000-\u001F\u007F]/g, ' ').trim().slice(0, max); }

export function profileForSender(graph, sender, branding) {
  const email = emailAddress(sender.emailAddress);
  const domain = email.split('@')[1].toLowerCase();
  if (!branding.approvedSenderDomains.map(d => d.toLowerCase()).includes(domain)) throw new Error('UNAPPROVED_SENDER');
  const identityMatch = [graph.mail, graph.userPrincipalName].filter(Boolean).some(e => e.toLowerCase() === email.toLowerCase());
  // Never put the signed-in person's title or numbers on another sending identity.
  if (!identityMatch) return {name:text(sender.displayName) || email, email, title:'', office:'', phones:[], directoryMatched:false};
  const phoneValues = [...(Array.isArray(graph.businessPhones) ? graph.businessPhones : [])];
  if (branding.includeMobilePhone && graph.mobilePhone) phoneValues.push(graph.mobilePhone);
  const seen = new Set();
  const phones = phoneValues.map(v => text(v, 60)).filter(v => {
    const key=v.replace(/[^0-9+]/g,'');
    if (!key || seen.has(key)) return false;
    seen.add(key); return true;
  }).slice(0, branding.maxPhoneNumbers);
  return {name:text(graph.displayName)||text(sender.displayName)||email, email, title:text(graph.jobTitle), office:branding.showOfficeLocation ? text(graph.officeLocation) : '', phones, directoryMatched:true};
}

function phoneHtml(phone) {
  const label=escapeHtml(phone).replace(/ /g,'&#160;');
  const compact=phone.replace(/[\s().-]/g,'');
  return /^\+[1-9]\d{6,14}$/.test(compact)
    ? `<a href="tel:${compact}" style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#102326;text-decoration:none;">${label}</a>`
    : label;
}

export function fillTemplate(template, values) {
  let rendered=template.replace(/{{#([A-Z_]+)}}([\s\S]*?){{\/\1}}/g, (_,key,body)=>values[key] ? body : '');
  rendered=rendered.replace(/{{([A-Z_]+)}}/g, (_,key)=>{
    if (!Object.hasOwn(values,key)) throw new Error(`UNKNOWN_TEMPLATE_FIELD:${key}`);
    return values[key];
  });
  if (/{{|}}/.test(rendered)) throw new Error('UNRESOLVED_TEMPLATE_FIELD');
  return rendered;
}

export function renderSignature(bundle, profile, {compact=false, images='cid', officeCss=false}={}) {
  const b=bundle.branding;
  const artesianUrl=httpsUrl(b.artesianWebsite,true);
  const logo = key => images==='preview' ? `data:image/png;base64,${bundle.assets[key].base64}` : `cid:${bundle.assets[key].filename}`;
  const artesianName=!compact && b.artesianWordmark
    ? `<img src="${logo('artesian')}" alt="Artesian" width="62" height="18" border="0" style="display:inline-block;width:62px;height:18px;border:0;vertical-align:middle;">`
    : 'Artesian';
  const values={
    NAME:escapeHtml(profile.name), TITLE:escapeHtml(profile.title), EMAIL:escapeHtml(profile.email),
    EMAIL_HREF:escapeHtml('mailto:'+encodeURIComponent(emailAddress(profile.email)).replace('%40','@')),
    PHONES:profile.phones.map(phoneHtml).join(' <span style="color:#A5ADB1;">&nbsp;&middot;&nbsp;</span> '),
    OFFICE:escapeHtml(b.locationLine || (b.showOfficeLocation?profile.office:'')), DESCRIPTOR:escapeHtml(b.descriptor), RELATIONSHIP:escapeHtml(b.relationship),
    ARK_URL:escapeHtml(httpsUrl(b.arkWebsite)), ARK_LABEL:escapeHtml(b.arkWebsiteLabel),
    ARTESIAN_URL:escapeHtml(httpsUrl(b.artesianWebsite,true)), ARTESIAN_LABEL:escapeHtml(b.artesianWebsiteLabel),
    ARK_LOGO:logo('ark'),
    ARTESIAN_LINK:artesianUrl?`<a href="${escapeHtml(artesianUrl)}" style="font-family:Arial,Helvetica,sans-serif;font-size:12px;font-weight:600;color:#006D7B;text-decoration:none;">${artesianName}</a>`:artesianName
  };
  const html=fillTemplate(compact?bundle.templates.reply:bundle.templates.full,values);
  const result=officeCss ? toOfficeCss(html,bundle.revision) : html;
  if (images!=='preview' && result.length>30000) throw new Error('SIGNATURE_TOO_LARGE');
  return result;
}

// setSignatureAsync currently documents internal CSS support, not inline CSS.
// Keep portable inline templates as source and give Outlook a scoped style block.
export function toOfficeCss(html, revision) {
  const styles=new Map(); const prefix='arksig_'+String(revision).replace(/[^a-z0-9]/gi,'').slice(0,16)+'_';
  const converted=html.replace(/\sstyle="([^"]*)"/g,(_,style)=>{
    if (!styles.has(style)) styles.set(style,prefix+styles.size);
    return ` class="${styles.get(style)}"`;
  });
  const css=[...styles].map(([style,name])=>'.'+name+'{'+style+'}').join('\n');
  return `<style type="text/css">${css}</style>${converted}`;
}

export function plainSignature(bundle, p, compact=false) {
  const lines=[p.name,p.title,p.email,...p.phones].filter(Boolean);
  if (compact) lines.push(`ARK · ${bundle.branding.relationship} Artesian`);
  else lines.push('',`ARK · ${bundle.branding.relationship} Artesian`,bundle.branding.descriptor,bundle.branding.arkWebsite,...(bundle.branding.artesianWebsite?[bundle.branding.artesianWebsite]:[]));
  const location=bundle.branding.locationLine || (bundle.branding.showOfficeLocation?p.office:'');
  if (location && !compact) lines.push(location);
  return lines.join('\n');
}
