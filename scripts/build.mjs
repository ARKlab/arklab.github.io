import {readFile,writeFile,mkdir,cp,rm} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {build} from 'esbuild';
import {httpsUrl,renderSignature} from '../src/render.js';

const {version}=JSON.parse(await readFile('package.json','utf8'));
const deployment=JSON.parse(await readFile('deployment.json','utf8'));
const branding=JSON.parse(await readFile('branding.json','utf8'));
const uuid=/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const allowedDeployment=['siteUrl','addinId','manifestVersion','tenantId','clientId','enabled'];
if(Object.keys(deployment).some(k=>!allowedDeployment.includes(k))) throw new Error('Unexpected deployment field. No secrets belong in this static project.');
const site=new URL(httpsUrl(deployment.siteUrl));
if(site.pathname!=='/' || site.search || site.hash) throw new Error('The add-in must be hosted at an origin root for the Outlook well-known file.');
if(!uuid.test(deployment.addinId)||!/^\d+\.\d+\.\d+\.\d+$/.test(deployment.manifestVersion)) throw new Error('Invalid manifest identity/version.');
for(const key of ['clientId','tenantId']) if(deployment[key]&&!uuid.test(deployment[key])) throw new Error('Invalid '+key);
if(typeof deployment.enabled!=='boolean'||(deployment.enabled&&(!deployment.tenantId||!deployment.clientId))) throw new Error('Configure Microsoft 365 before enabling the pilot.');
if(branding.schemaVersion!==1||!Array.isArray(branding.approvedSenderDomains)||!branding.approvedSenderDomains.length) throw new Error('Invalid branding configuration.');
if(!Number.isInteger(branding.maxPhoneNumbers)||branding.maxPhoneNumbers<0||branding.maxPhoneNumbers>2) throw new Error('Choose zero, one or two phone numbers.');
for(const key of ['includeMobilePhone','showOfficeLocation','compactReplies','artesianWordmark']) if(typeof branding[key]!=='boolean') throw new Error('Invalid '+key);
for(const key of ['descriptor','relationship','arkWebsiteLabel','artesianWebsiteLabel','locationLine']) if(typeof branding[key]!=='string'||branding[key].length>150) throw new Error('Invalid '+key);
if(branding.approvedSenderDomains.some(d=>typeof d!=='string'||!/^([a-z0-9-]+\.)+[a-z]{2,}$/i.test(d))) throw new Error('Invalid sending domain.');
httpsUrl(branding.arkWebsite);httpsUrl(branding.artesianWebsite,true);
const templates={full:await readFile('templates/full.html','utf8'),reply:await readFile('templates/reply.html','utf8')};
for(const template of Object.values(templates)) if(/<(script|iframe|object|embed|form|input)\b|\son[a-z]+\s*=|javascript:/i.test(template)) throw new Error('Templates must contain passive email HTML only.');
const assets={};
for(const [key,path] of Object.entries({ark:'public/assets/ark-logo.png',artesian:'public/assets/artesian-wordmark-color.png'})) {
  const bytes=await readFile(path);
  if(bytes.length>100000||bytes.subarray(0,8).toString('hex')!=='89504e470d0a1a0a') throw new Error('Use small PNG logo assets.');
  const hash=createHash('sha256').update(bytes).digest('hex').slice(0,12);
  assets[key]={filename:`ark-signatures-${key}-${hash}.png`,base64:bytes.toString('base64')};
}
const contents={branding,templates,assets,deployment,enabled:deployment.enabled};
const revision=createHash('sha256').update(JSON.stringify(contents)).digest('hex').slice(0,12);
const bundle={...contents,revision};
const sample={name:'First name Last name',title:'Job title',email:'name@example.com',office:'Office city, Country',phones:['+353 00 000 0000']};
renderSignature(bundle,sample,{officeCss:true});renderSignature(bundle,sample,{compact:true,officeCss:true});
await rm('dist',{recursive:true,force:true});await mkdir('dist',{recursive:true});
await cp('public','dist',{recursive:true});
await writeFile('dist/signature-bundle.json',JSON.stringify(bundle));
await mkdir('dist/.well-known',{recursive:true});
await writeFile('dist/.well-known/microsoft-officeaddins-allowed.json',JSON.stringify({allowed:[new URL('runtime.js',site).href]},null,2));
await writeFile('dist/.nojekyll','');
await build({entryPoints:['src/runtime.js','src/taskpane.js','src/preview.js'],outdir:'dist',bundle:true,format:'iife',platform:'browser',target:'es2020',minify:true,legalComments:'eof',define:{__SITE_URL__:JSON.stringify(site.href),__APP_VERSION__:JSON.stringify(version)}});

const xml=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const url=file=>xml(new URL(file,site).href);
function command(surface,id){return `<ExtensionPoint xsi:type="${surface}"><OfficeTab id="TabDefault"><Group id="${id}Group"><Label resid="GroupLabel"/><Control xsi:type="Button" id="${id}"><Label resid="PaneLabel"/><Supertip><Title resid="PaneLabel"/><Description resid="PaneDescription"/></Supertip><Icon><bt:Image size="16" resid="Icon16x16"/><bt:Image size="32" resid="Icon32x32"/><bt:Image size="80" resid="Icon80x80"/></Icon><Action xsi:type="ShowTaskpane"><SourceLocation resid="TaskpaneUrl"/></Action></Control></Group></OfficeTab></ExtensionPoint>`;}
const manifest=`<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="MailApp">
<Id>${xml(deployment.addinId)}</Id><Version>${xml(deployment.manifestVersion)}</Version><ProviderName>ARK</ProviderName><DefaultLocale>en-US</DefaultLocale><DisplayName DefaultValue="ARK signatures"/><Description DefaultValue="ARK, home of Artesian. Centrally managed employee signatures."/><IconUrl DefaultValue="${url('assets/icon-64.png')}"/><HighResolutionIconUrl DefaultValue="${url('assets/icon-128.png')}"/><SupportUrl DefaultValue="${url('')}"/>
<AppDomains><AppDomain>${xml(site.origin)}</AppDomain><AppDomain>https://login.microsoftonline.com</AppDomain><AppDomain>https://graph.microsoft.com</AppDomain></AppDomains><Hosts><Host Name="Mailbox"/></Hosts><Requirements><Sets DefaultMinVersion="1.1"><Set Name="Mailbox"/></Sets></Requirements>
<FormSettings><Form xsi:type="ItemRead"><DesktopSettings><SourceLocation DefaultValue="${url('taskpane.html')}"/><RequestedHeight>300</RequestedHeight></DesktopSettings></Form></FormSettings><Permissions>ReadWriteItem</Permissions><Rule xsi:type="RuleCollection" Mode="Or"><Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/></Rule><DisableEntityHighlighting>true</DisableEntityHighlighting>
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0"><VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1"><Description resid="AppDescription"/><Requirements><bt:Sets DefaultMinVersion="1.13"><bt:Set Name="Mailbox"/></bt:Sets></Requirements><Hosts><Host xsi:type="MailHost"><Runtimes><Runtime resid="RuntimeUrl"><Override type="javascript" resid="RuntimeJs"/></Runtime></Runtimes><DesktopFormFactor>
${command('MessageReadCommandSurface','ArkReadPane')}${command('MessageComposeCommandSurface','ArkComposePane')}
<ExtensionPoint xsi:type="LaunchEvent"><LaunchEvents><LaunchEvent Type="OnNewMessageCompose" FunctionName="arkOnCompose"/><LaunchEvent Type="OnMessageFromChanged" FunctionName="arkOnFromChanged"/></LaunchEvents><SourceLocation resid="RuntimeUrl"/></ExtensionPoint></DesktopFormFactor></Host></Hosts>
<Resources><bt:Images><bt:Image id="Icon16x16" DefaultValue="${url('assets/icon-16.png')}"/><bt:Image id="Icon32x32" DefaultValue="${url('assets/icon-32.png')}"/><bt:Image id="Icon80x80" DefaultValue="${url('assets/icon-80.png')}"/></bt:Images><bt:Urls><bt:Url id="TaskpaneUrl" DefaultValue="${url('taskpane.html')}"/><bt:Url id="RuntimeUrl" DefaultValue="${url('runtime.html')}"/><bt:Url id="RuntimeJs" DefaultValue="${url('runtime.js')}"/></bt:Urls><bt:ShortStrings><bt:String id="GroupLabel" DefaultValue="ARK"/><bt:String id="PaneLabel" DefaultValue="ARK signatures"/></bt:ShortStrings><bt:LongStrings><bt:String id="AppDescription" DefaultValue="ARK, home of Artesian. Centrally managed employee signatures."/><bt:String id="PaneDescription" DefaultValue="Connect Microsoft 365 and preview or refresh your company signature."/></bt:LongStrings></Resources></VersionOverrides></VersionOverrides></OfficeApp>`;
await writeFile('dist/manifest.xml',manifest);
await writeFile('dist/employee-template.html','<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Signature preview</title><body>'+renderSignature(bundle,sample,{images:'preview'})+'</body></html>');
console.log(`Built revision ${revision}; automatic insertion ${deployment.enabled?'enabled':'paused'}.`);
