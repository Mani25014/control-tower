/* NovintiX Delivery Management Control Tower — Microsoft 365 edition
   Sign-in: Microsoft Entra ID (MSAL), NovintiX accounts only.
   Data: SharePoint lists on the configured site (or built-in sample data in demo mode). */
(()=>{
"use strict";
const CFG0=window.CT_CONFIG||{};
const ASSETS=window.CT_ASSETS||{};
const DEMO=!!CFG0.demo||!CFG0.clientId;

/* ---------- screens ---------- */
const SCREENS=[
 ['dashboard','Executive dashboard','Overview','--amber','<path d="M3 13h8V3H3zm10 8h8V11h-8zM3 21h8v-6H3zm10-18v6h8V3z"/>'],
 ['team','Employee database','Masters','--c-team','<path d="M16 11a3 3 0 1 0-3-3 3 3 0 0 0 3 3zm-8 0a3 3 0 1 0-3-3 3 3 0 0 0 3 3zm0 2c-2.7 0-8 1.3-8 4v3h16v-3c0-2.7-5.3-4-8-4zm8 0c-.3 0-.7 0-1.1.1A4.2 4.2 0 0 1 17 17v3h7v-3c0-2.7-5.3-4-8-4z"/>'],
 ['projects','Projects & initiatives','Masters','--c-proj','<path d="M10 4H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-8z"/>'],
 ['allocation','Billability & allocation','Delivery','--c-util','<path d="M5 9h3v11H5zm5.5-5h3v16h-3zM16 13h3v7h-3z"/>'],
 ['activities','Activities & responsibilities','Delivery','--c-act','<path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"/>'],
 ['productivity','Productivity tracker','Delivery','--c-prod','<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm1 11h-6v-2h4V6h2z"/>'],
 ['leave','Leave, WFH & permission','People','--c-leave','<path d="M19 4h-1V2h-2v2H8V2H6v2H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2zm0 16H5V9h14z"/>'],
 ['userids','Customer user IDs','People','--c-uid','<path d="M12 1 3 5v6c0 5.5 3.8 10.7 9 12 5.2-1.3 9-6.5 9-12V5zm0 6a3 3 0 1 1-3 3 3 3 0 0 1 3-3zm0 13.2a7.6 7.6 0 0 1-5-3.6c.1-1.7 3.3-2.6 5-2.6s4.9.9 5 2.6a7.6 7.6 0 0 1-5 3.6z"/>'],
 ['access','Team access','Admin','--c-asg','<path d="M12 12a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm0 2c-2.7 0-8 1.3-8 4v2h9.5a6 6 0 0 1 4.6-5.8A13 13 0 0 0 12 14zm7 1-1.1 2.2-2.4.4 1.8 1.7-.4 2.4 2.1-1.1 2.1 1.1-.4-2.4 1.8-1.7-2.4-.4z"/>'],
 ['settings','Settings','Admin','--faint','<path d="M19.4 13a7.5 7.5 0 0 0 0-2l2.1-1.6-2-3.5-2.5 1a7.4 7.4 0 0 0-1.7-1L15 3h-4l-.4 2.7a7.4 7.4 0 0 0-1.7 1l-2.5-1-2 3.5L6.6 11a7.5 7.5 0 0 0 0 2l-2.1 1.6 2 3.5 2.5-1a7.4 7.4 0 0 0 1.7 1L11 21h4l.4-2.7a7.4 7.4 0 0 0 1.7-1l2.5 1 2-3.5zM13 15.5A3.5 3.5 0 1 1 16.5 12 3.5 3.5 0 0 1 13 15.5z"/>']
];
const SC=Object.fromEntries(SCREENS.map(s=>[s[0],s]));
const SECTIONS=['Overview','Masters','Delivery','People','Admin'];
const LEVELS=[['none','No access'],['view','View'],['own','Own entries'],['edit','Edit']];

/* ---------- SharePoint list schema ---------- */
const N='n',T='t',M='m'; // number, text, multi-line text
const SCHEMA={
 employees:{list:'CT_Employees',label:'Employee database',f:[['name','Employee Name',T],['designation','Designation',T],['email','Email',T],['managerId','Manager ID',T],['managerName','Manager Name',T],['orgRole','Org Role',T],['projectRole','Project Role',T],['doj','Date Of Joining',T],['yoePrior','YoE Before NovintiX',N],['dob','Birthday',T],['bu','Business Unit',T],['billable','Billable',T],['billableType','Billable Type',T],['location','Location',T],['country','Country',T],['empType','Employment Type',T],['workMode','Work Mode',T],['shift','Shift',T],['status','Employment Status',T],['seat','Seat No',T],['remarks','Remarks',M]]},
 projects:{list:'CT_Projects',label:'Projects & initiatives',f:[['portfolio','Portfolio',T],['program','Program',T],['project','Project',T],['goal','Goal',M],['status','Status',T],['start','Start Date',T],['end','End Date',T],['contract','Contract Type',T],['revenue','Revenue Type',T],['sowStart','SOW Start',T],['sowEnd','SOW End',T]]},
 alloc:{list:'CT_Allocation',label:'Billability & allocation',f:[['projectId','Project ID',T],['empId','Employee ID',T],['role','Role',T],['alloc','Allocation %',N],['type','Allocation Type',T],['remarks','Remarks',M]]},
 activities:{list:'CT_Activities',label:'Activities & responsibilities',f:[['projectId','Project ID',T],['actId','Activity ID',T],['ownerId','Owner ID',T],['desc','Activity Description',M],['frequency','Frequency',T],['priority','Priority',T],['status','Status',T],['start','Start Date',T],['target','Target Date',T],['actualEnd','Actual End Date',T],['backup1','Backup 1 ID',T],['backup2','Backup 2 ID',T],['remarks','Remarks',M]]},
 productivity:{list:'CT_Productivity',label:'Productivity tracker',f:[['empId','Employee ID',T],['actId','Activity ID',T],['actDesc','Activity Description',M],['projectId','Project ID',T],['sub','Sub Activity',M],['date','Effort Date',T],['hours','Effort Hours',N],['remarks','Remarks',M],['by','Logged By',T]]},
 leave:{list:'CT_Leave',label:'Leave, WFH & permission',f:[['empId','Employee ID',T],['date','Date',T],['kind','Type',T],['clientInformed','Client Informed',T],['clients','Clients',T],['remarks','Remarks',M]]},
 userids:{list:'CT_CustomerIDs',label:'Customer user IDs',f:[['empId','Employee ID',T],['projectId','Project ID',T],['program','Program',T],['project','Project Name',T],['sponsor','Project Sponsor',T],['userName','User Name',T],['expiry','Expiry Date',T],['status','Status',T]]},
 access:{list:'CT_Access',label:'Team access',f:[['name','Name',T],['empId','Employee ID',T],['role','App Role',T],['active','Active',T],['addedBy','Added By',T]]},
 config:{list:'CT_Config',label:'App settings',f:[['value','Value',M]]}
};
const TABLES=Object.keys(SCHEMA);
const iname=k=>'ct'+k[0].toUpperCase()+k.slice(1);
const CF_TABLES={employees:'Employee database',projects:'Projects & initiatives',alloc:'Billability & allocation',activities:'Activities & responsibilities',userids:'Customer user IDs',leave:'Leave, WFH & permission'};

/* ---------- defaults for customizable settings ---------- */
const LIST_DEF={
 orgRole:['Org role',['Team Lead','Core Team','PMO','Team Member']],
 projectRole:['Project role',['Program Lead','Project Lead','Project Executor']],
 billableType:['Billable type',['Single Role','Dual Role','Partial Role','Indirect','Irrelevant','Yet To Map']],
 empType:['Employment type',['Full Time','Intern','Contract']],
 workMode:['Work mode',['Office','Remote','Hybrid']],
 shift:['Shift',['General','Day','Night']],
 empStatus:['Employment status',['Active','Notice Period','Inactive','Exited']],
 portfolio:['Portfolio',['Product Engineering','Operations']],
 prjStatus:['Project / initiative status',['Active','Planned','In-Progress','Accelerate','On Hold','Closed']],
 contract:['Contract type',['T&M','Fixed Price','MSM','-']],
 revenue:['Revenue type',['Billable','Non Billable','-']],
 allocRole:['Allocation role',['Client Lead','Project Manager','Delivery Lead','SME','Executor','Support']],
 actFrequency:['Activity frequency',['Once','Daily','Twice a week','Weekly','Bi-Weekly','Monthly','Quarterly','Half Yearly','As Needed']],
 actPriority:['Activity priority',['High','Medium','Low']],
 actStatus:['Activity status',['Open','In-Progress','Hold','Closed']],
 uidStatus:['Customer account status',['Active','Requested','In-Active','Revoked']]
};
const LOCKED={revenue:'Billable',actStatus:'Closed',empStatus:'Active'};
const ALL=(lvl)=>Object.fromEntries(SCREENS.map(s=>[s[0],lvl]));
const ROLE_DEF={
 'Admin':ALL('edit'),
 'Leadership':{...ALL('edit'),access:'view',settings:'none'},
 'Core Team':{...ALL('edit'),access:'view',settings:'none'},
 'Project Lead':{dashboard:'view',team:'view',projects:'view',allocation:'edit',activities:'edit',productivity:'edit',leave:'edit',userids:'view',access:'none',settings:'none'},
 'Team Member':{dashboard:'none',team:'view',projects:'view',allocation:'view',activities:'view',productivity:'own',leave:'view',userids:'none',access:'none',settings:'none'},
 'Viewer':{dashboard:'view',team:'view',projects:'view',allocation:'view',activities:'view',productivity:'view',leave:'view',userids:'view',access:'none',settings:'none'}
};

/* ---------- state ---------- */
const S={t:{},ready:false,missing:[],setup:null};
for(const t of TABLES)S.t[t]={};
const UI={page:'dashboard',f:{},sort:{},tab:{},week:null,day:null,cal:null,per:'month'};
let ME={email:'',name:'',role:null,empId:null};
let STATE='boot'; // boot | signin | denied | pending | setup | app
let AUTH_ERR='';

/* ---------- utils ---------- */
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const $=s=>document.querySelector(s);
const pad=n=>String(n).padStart(2,'0');
const isoD=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const TODAY=()=>isoD(new Date());
const parse=s=>{const[y,m,d]=String(s).split('-').map(Number);return new Date(y,(m||1)-1,d||1)};
const addD=(s,n)=>{const d=parse(s);d.setDate(d.getDate()+n);return isoD(d)};
const MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const fd=s=>{if(!s)return'—';const d=parse(s);return`${d.getDate()} ${MON[d.getMonth()]} ${d.getFullYear()}`};
const fds=s=>{if(!s)return'—';const d=parse(s);return`${d.getDate()} ${MON[d.getMonth()]}`};
const daysBetween=(a,b)=>Math.round((parse(b)-parse(a))/864e5);
const isWE=s=>{const g=parse(s).getDay();return g===0||g===6};
const monday=s=>{const d=parse(s);const g=(d.getDay()+6)%7;d.setDate(d.getDate()-g);return isoD(d)};
const nid=p=>p+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const n1=x=>(Math.round((+x||0)*10)/10).toLocaleString();
const pct=(a,b)=>b?Math.round(a/b*100):0;
const ttxt=s=>String(s??'').toLowerCase();
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const opt=arr=>[...new Set(arr.filter(Boolean))].sort();
const rows=t=>S.t[t]||{};
const E=()=>rows('employees'),P=()=>rows('projects'),AL=()=>rows('alloc'),AC=()=>rows('activities'),UID=()=>rows('userids'),LV=()=>rows('leave'),TS=()=>rows('productivity'),ACC=()=>rows('access');
const empName=id=>E()[id]?.name||id||'—';
const first=id=>empName(id).split(' ')[0];
const prj=id=>P()[id];
const prjShort=id=>{const p=prj(id);if(!p)return id||'—';const dupe=Object.values(P()).filter(q=>q.program===p.program&&q.project===p.project).length>1;if(!dupe)return p.project;const g=(p.goal||'').split(/ - | – |\(/)[0].trim();return`${p.project} (${g.length>28?g.slice(0,26)+'…':g})`};
const prjName=id=>{const p=prj(id);return p?`${p.program} · ${prjShort(id)}`:(id||'—')};
const isClient=id=>/^PE_/.test(id||'');
const ACTIVE_ST=['Active','In-Progress','Accelerate'];
function stPill(s,label){const m={Active:'ok',Closed:'',Planned:'info','In-Progress':'warn',Accelerate:'ok','On Hold':'bad',Hold:'bad',Open:'info',Billable:'ok',Shadow:'',Yes:'ok',No:'','In-Active':'bad',Admin:'bad',Leadership:'info','Core Team':'info','Project Lead':'warn','Team Member':'',Viewer:''};return`<span class="pill ${m[s]??''}">${esc(label||s||'Not set')}</span>`}
function toast(msg,err){const t=document.createElement('div');t.className='toast'+(err?' err':'');t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),err?6000:2600)}
function errMsg(e){if(e&&e.status===403)return"SharePoint refused this change. Ask an admin to check your permission on the Control Tower site.";if(e&&e.status===404)return'That record no longer exists. Refresh and try again.';if(e&&e.msg)return e.msg;return'Could not save. Check your connection and try again.'}

/* ---------- settings (stored in CT_Config) ---------- */
function cfgVal(k){const r=rows('config')[k];if(!r||!r.value)return null;try{return JSON.parse(r.value)}catch(e){return null}}
function CFG(){const lists={},sl=cfgVal('lists')||{};for(const k in LIST_DEF)lists[k]=Array.isArray(sl[k])&&sl[k].length?sl[k]:LIST_DEF[k][1];
  const roles=cfgVal('roles')||ROLE_DEF;roles.Admin=ALL('edit');
  return{lists,roles,pages:cfgVal('pages')||{},custom:cfgVal('custom')||{},hours:{day:8,half:4,perm:2,...(cfgVal('hours')||{})}}}
const L=k=>CFG().lists[k];
const H=()=>CFG().hours;
const TITLE=k=>CFG().pages[k]?.title||SC[k][1];
const cfFields=t=>CFG().custom[t]||[];
const cfCols=t=>cfFields(t).map(f=>({k:'x_'+f.key,t:f.label,v:x=>f.type==='date'?fd(x['x_'+f.key]):(x['x_'+f.key]??'')}));
const cfExport=(t,r)=>Object.fromEntries(cfFields(t).map(f=>[f.label,r['x_'+f.key]??'']));
const fieldsOf=t=>[...SCHEMA[t].f,...(CF_TABLES[t]?cfFields(t).map(f=>['x_'+f.key,f.label,f.type==='number'?N:f.type==='long'?M:T]):[])];

/* ---------- permissions ---------- */
function perm(screen){if(!ME.role)return'none';if(ME.role==='Admin')return'edit';const r=CFG().roles[ME.role];const v=r?r[screen]:'none';if(CFG().pages[screen]?.show===false&&ME.role!=='Admin')return'none';return v||'none'}
const canSee=s=>perm(s)!=='none';
const canEdit=s=>perm(s)==='edit';

/* =========================================================
   DATA BACKENDS
   ========================================================= */
const Store={
  async loadAll(){},async upsert(){},async remove(){},async ensureField(){}
};

/* ---- demo backend: sample data in memory ---- */
function demoInit(){
  const seed=window.CT_SEED||{};
  for(const t of TABLES){S.t[t]={};for(const r of seed[t]||[])S.t[t][r.id]={...r}}
  const emps=Object.values(S.t.employees);
  for(const e of emps){const em=(e.email||'').trim().toLowerCase();if(!em)continue;S.t.access[em]={id:em,name:e.name,empId:e.id,role:defaultRole(e),active:'Yes',addedBy:'setup'}}
  Store.loadAll=async()=>{};
  Store.upsert=async(t,r)=>{S.t[t][r.id]={...r}};
  Store.remove=async(t,id)=>{delete S.t[t][id]};
  Store.ensureField=async()=>{};
}
function defaultRole(e){const b=(CFG0.bootstrapAdmins||[]).map(x=>x.toLowerCase());if(b.includes((e.email||'').trim().toLowerCase()))return'Admin';
  if(e.orgRole==='Team Lead')return'Leadership';if(e.orgRole==='Core Team'||e.orgRole==='PMO')return'Core Team';if(/Lead/.test(e.projectRole||''))return'Project Lead';return'Team Member'}

/* ---- SharePoint backend ---- */
let MSAL=null,ACCOUNT=null;
const SITE=(CFG0.siteUrl||'').replace(/\/$/,'');
const SP_ORIGIN=SITE?new URL(SITE).origin:'';
const SP_SCOPE=SP_ORIGIN+'/AllSites.Manage';
const ET={}; // list item entity types
const SPID={}; // table -> {key: spItemId}
async function getToken(force){
  try{const r=await MSAL.acquireTokenSilent({scopes:[SP_SCOPE],account:ACCOUNT,forceRefresh:!!force});return r.accessToken}
  catch(e){await MSAL.acquireTokenRedirect({scopes:[SP_SCOPE],account:ACCOUNT});throw e}
}
async function sp(path,{method='GET',body,headers={}}={}){
  let tok=await getToken();
  for(let attempt=0;attempt<6;attempt++){
    let r;
    try{r=await fetch(SITE+path,{method,headers:{Authorization:'Bearer '+tok,Accept:'application/json;odata=verbose','Content-Type':'application/json;odata=verbose',...headers},body:body?JSON.stringify(body):undefined})}
    catch(e){if(attempt<2){await sleep(800);continue}throw{status:0,msg:'Could not reach SharePoint. Check your connection.'}}
    if(r.status===429||r.status===503){await sleep(((+r.headers.get('Retry-After'))||2)*1000);continue}
    if(r.status===401&&attempt===0){tok=await getToken(true);continue}
    if(!r.ok){let m='';try{const j=JSON.parse(await r.text());m=j.error?.message?.value||''}catch(e){}throw{status:r.status,msg:m}}
    if(r.status===204)return null;const txt=await r.text();return txt?JSON.parse(txt).d:null;
  }
  throw{status:429,msg:'SharePoint is busy. Wait a minute and try again.'};
}
const lt=t=>`/_api/web/lists/getbytitle('${SCHEMA[t].list}')`;
function itemToRow(t,it){const r={id:it.Title};for(const[k,,ty]of fieldsOf(t)){let v=it[iname(k)];if(v==null)v=ty===N?0:'';r[k]=v}return r}
function rowToItem(t,r){const o={__metadata:{type:ET[t]},Title:String(r.id)};for(const[k,,ty]of fieldsOf(t)){let v=r[k];if(ty===N)v=v===''||v==null?null:+v;else v=v==null?'':String(v);o[iname(k)]=v}return o}
async function spLoadTable(t){
  const sel=['Id','Title',...fieldsOf(t).map(f=>iname(f[0]))].join(',');
  let url=`${lt(t)}/items?$select=${sel}&$top=5000`;const out={},ids={};
  while(url){const d=await sp(url);for(const it of d.results){const r=itemToRow(t,it);out[r.id]=r;ids[r.id]=it.Id}url=d.__next?d.__next.replace(SITE,''):null}
  S.t[t]=out;SPID[t]=ids;
}
async function spInit(){
  S.missing=[];
  await Promise.all(TABLES.map(async t=>{try{const d=await sp(`${lt(t)}?$select=ListItemEntityTypeFullName`);ET[t]=d.ListItemEntityTypeFullName}catch(e){if(e.status===404)S.missing.push(t);else throw e}}));
  if(S.missing.length)return;
  // load config + access first, then the rest
  await Promise.all(['config','access'].map(spLoadTable));
  await Promise.all(TABLES.filter(t=>!['config','access'].includes(t)).map(spLoadTable));
  Store.loadAll=async()=>{await Promise.all(TABLES.map(spLoadTable))};
  Store.upsert=async(t,r)=>{const body=rowToItem(t,r);const id=SPID[t][r.id];
    if(id)await sp(`${lt(t)}/items(${id})`,{method:'POST',body,headers:{'X-HTTP-Method':'MERGE','IF-MATCH':'*'}});
    else{const d=await sp(`${lt(t)}/items`,{method:'POST',body});SPID[t][r.id]=d.Id}
    S.t[t][r.id]={...r}};
  Store.remove=async(t,key)=>{const id=SPID[t][key];if(id)await sp(`${lt(t)}/items(${id})`,{method:'POST',headers:{'X-HTTP-Method':'DELETE','IF-MATCH':'*'}});delete S.t[t][key];delete SPID[t][key]};
  Store.ensureField=async(t,key,label,ty)=>{await spCreateField(t,key,label,ty)};
}
function fieldXml(k,label,ty){const n=iname(k);const lab=esc(label);
  if(ty===N)return`<Field Type="Number" Name="${n}" StaticName="${n}" DisplayName="${lab}" Decimals="2" />`;
  if(ty===M)return`<Field Type="Note" Name="${n}" StaticName="${n}" DisplayName="${lab}" NumLines="4" RichText="FALSE" />`;
  return`<Field Type="Text" Name="${n}" StaticName="${n}" DisplayName="${lab}" MaxLength="255" />`}
async function spCreateField(t,k,label,ty){
  try{await sp(`${lt(t)}/fields/getbyinternalnameortitle('${iname(k)}')?$select=Id`);return}catch(e){if(e.status!==404&&e.status!==400)throw e}
  await sp(`${lt(t)}/fields/createfieldasxml`,{method:'POST',body:{parameters:{__metadata:{type:'SP.XmlSchemaFieldCreationInformation'},SchemaXml:fieldXml(k,label,ty),Options:25}}});
}
async function spCreateList(t){
  try{const d=await sp(`${lt(t)}?$select=ListItemEntityTypeFullName`);ET[t]=d.ListItemEntityTypeFullName}
  catch(e){if(e.status!==404)throw e;
    await sp('/_api/web/lists',{method:'POST',body:{__metadata:{type:'SP.List'},AllowContentTypes:true,BaseTemplate:100,ContentTypesEnabled:false,Description:'NovintiX Control Tower — '+SCHEMA[t].label,Title:SCHEMA[t].list}});
    const d=await sp(`${lt(t)}?$select=ListItemEntityTypeFullName`);ET[t]=d.ListItemEntityTypeFullName}
  for(const[k,label,ty]of SCHEMA[t].f)await spCreateField(t,k,label,ty);
}
async function spAddToSiteGroup(email,editRole){
  const grp=await sp(`/_api/web/${editRole?'AssociatedMemberGroup':'AssociatedVisitorGroup'}?$select=Id`);
  await sp(`/_api/web/sitegroups(${grp.Id})/users`,{method:'POST',body:{__metadata:{type:'SP.User'},LoginName:'i:0#.f|membership|'+email}});
}

/* =========================================================
   AUTH + BOOT
   ========================================================= */
const domOK=email=>(CFG0.allowedDomains||['novintix.com']).some(d=>email.endsWith('@'+d.toLowerCase()));
async function boot(){
  window.addEventListener('hashchange',route);
  if(DEMO){demoInit();STATE='signin';paint();return}
  if(!window.msal){AUTH_ERR='The Microsoft sign-in library did not load. Check that msal-browser.min.js is next to index.html.';STATE='signin';paint();return}
  MSAL=new msal.PublicClientApplication({auth:{clientId:CFG0.clientId,authority:'https://login.microsoftonline.com/'+(CFG0.tenant||'organizations'),redirectUri:location.origin+location.pathname,navigateToLoginRequestUrl:false},cache:{cacheLocation:'localStorage'}});
  try{if(MSAL.initialize)await MSAL.initialize();const res=await MSAL.handleRedirectPromise();ACCOUNT=res?.account||MSAL.getAllAccounts()[0]||null}
  catch(e){AUTH_ERR='Sign-in did not complete: '+(e.errorMessage||e.message||e);ACCOUNT=null}
  if(!ACCOUNT){STATE='signin';paint();return}
  await afterSignIn(ACCOUNT.username,ACCOUNT.name);
}
async function afterSignIn(email,name){
  ME.email=String(email||'').toLowerCase();ME.name=name||ME.email;
  if(!domOK(ME.email)){STATE='denied';paint();return}
  STATE='boot';paint();
  if(!DEMO){try{await spInit()}catch(e){AUTH_ERR=e.status===403||e.status===401?"You signed in, but you don't have access to the Control Tower SharePoint site yet. Ask an admin to add you in Team access.":('Could not open the SharePoint site: '+(e.msg||e.status));STATE='denied';paint();return}}
  const boot=(CFG0.bootstrapAdmins||[]).map(x=>x.toLowerCase()).includes(ME.email);
  if(S.missing.length){if(boot){ME.role='Admin';STATE='setup'}else STATE='pending';paint();return}
  resolveMe(boot);
  STATE=ME.role?'app':'pending';route();
  if(!DEMO)setInterval(()=>{if(document.visibilityState==='visible'&&!FCTX)Store.loadAll().then(()=>{resolveMe(boot);paint()}).catch(()=>{})},120000);
}
function resolveMe(boot){const a=ACC()[ME.email];
  if(a&&a.active!=='No'){ME.role=a.role;ME.empId=a.empId||null;if(a.name)ME.name=a.name}
  else if(boot){ME.role='Admin';ME.empId=Object.values(E()).find(e=>(e.email||'').trim().toLowerCase()===ME.email)?.id||null}
  else{ME.role=null}
  if(!ME.empId){const e=Object.values(E()).find(x=>(x.email||'').trim().toLowerCase()===ME.email);if(e)ME.empId=e.id}
}
function signIn(){AUTH_ERR='';MSAL.loginRedirect({scopes:[SP_SCOPE],prompt:'select_account',domainHint:(CFG0.allowedDomains||['novintix.com'])[0]})}
function signOut(){if(DEMO){ME={email:'',name:'',role:null,empId:null};STATE='signin';paint();return}MSAL.logoutRedirect({account:ACCOUNT,postLogoutRedirectUri:location.origin+location.pathname})}

/* =========================================================
   SHELL
   ========================================================= */
function route(){const p=(location.hash||'').slice(1);UI.page=SC[p]?p:'dashboard';if(STATE==='app'&&!canSee(UI.page)){UI.page=SCREENS.map(s=>s[0]).find(canSee)||'dashboard'}paint();window.scrollTo(0,0)}
let rq=0;function render(){if(rq)return;rq=requestAnimationFrame(()=>{rq=0;paint()})}
const MS_LOGO='<svg viewBox="0 0 21 21" aria-hidden="true"><rect x="1" y="1" width="9" height="9" fill="#F25022"/><rect x="11" y="1" width="9" height="9" fill="#7FBA00"/><rect x="1" y="11" width="9" height="9" fill="#00A4EF"/><rect x="11" y="11" width="9" height="9" fill="#FFB900"/></svg>';
function paint(){
  const root=$('#root');
  if(STATE!=='app'){root.innerHTML=gate();return}
  if(!root.querySelector('.app'))root.innerHTML=`<div class="app"><aside class="rail"><div class="brand">${ASSETS.logo?`<img class="logo-img" src="${ASSETS.logo}" alt="NovintiX">`:'<div class="logo">novinti<b>X</b></div>'}<div class="tag">Accelerate transformation</div><div class="brand-bar"></div><h1>Delivery Management<br><span>Control Tower</span></h1></div><nav id="nav"></nav><div class="me" id="me"></div></aside><div style="min-width:0"><div class="mobnav" id="mobnav"></div><main id="view"></main></div></div>`;
  renderNav();renderMe();
  const v=$('#view');const ae=document.activeElement,aid=ae&&ae.id,sel=ae&&ae.selectionStart;
  v.style.setProperty('--mod',`var(${SC[UI.page][3]})`);
  v.innerHTML=(DEMO?`<div class="banner">Demo mode: sample data, nothing is saved. Deploy with Microsoft sign-in to store changes in SharePoint.</div>`:'')+(VIEWS[UI.page]||VIEWS.dashboard)();
  if(aid){const el=document.getElementById(aid);if(el){el.focus();try{if(sel!=null)el.setSelectionRange(sel,sel)}catch(e){}}}
}
function gate(){
  const card=(inner)=>`<div class="signin"><div class="box">${ASSETS.title?`<img src="${ASSETS.title}" alt="Delivery Management Control Tower" style="width:100%;height:auto;border-radius:12px">`:''}${inner}</div></div>`;
  const logo=ASSETS.logo?`<img class="logo-img" src="${ASSETS.logo}" alt="NovintiX">`:'';
  if(STATE==='boot')return card(`${logo}<p>Opening the Control Tower…</p><div class="progress"><span style="width:40%"></span></div>`);
  if(STATE==='signin'){
    if(DEMO){const accs=Object.values(ACC()).sort((a,b)=>(a.role==='Admin'?-1:0)-(b.role==='Admin'?-1:0)||a.name.localeCompare(b.name));
      return card(`${logo}<h1>Sign in to the <span>Control Tower</span></h1><p>Use your NovintiX Microsoft 365 account (<b>@novintix.com</b>).</p><label class="fine" for="demo-who">Demo: choose who to sign in as, to preview what each role can see</label><select id="demo-who">${accs.map(a=>`<option value="${esc(a.id)}">${esc(a.name)} · ${esc(a.role)}</option>`).join('')}</select><button class="ms-btn" data-demosignin="1">${MS_LOGO}Sign in with Microsoft</button><p class="fine">This preview uses sample data. The deployed app signs people in with Microsoft and only lets @novintix.com accounts in.</p>`)}
    return card(`${logo}<h1>Sign in to the <span>Control Tower</span></h1><p>Use your NovintiX Microsoft 365 account.</p>${AUTH_ERR?`<div class="err">${esc(AUTH_ERR)}</div>`:''}<button class="ms-btn" data-signin="1">${MS_LOGO}Sign in with Microsoft</button><p class="fine">Only @${esc((CFG0.allowedDomains||['novintix.com']).join(', @'))} accounts can sign in. Access is given by your Control Tower admin.</p>`)}
  if(STATE==='denied')return card(`${logo}<h1>You can't open the <span>Control Tower</span> with this account</h1>${AUTH_ERR?`<div class="err">${esc(AUTH_ERR)}</div>`:`<div class="err">You're signed in as <b>${esc(ME.email)}</b>. Only @${esc((CFG0.allowedDomains||['novintix.com']).join(', @'))} accounts can use this app.</div>`}<button class="ms-btn" data-signout="1">Sign out and use another account</button>`);
  if(STATE==='pending')return card(`${logo}<h1>Access <span>pending</span></h1><p>You're signed in as <b>${esc(ME.email)}</b>, but you haven't been given access yet${S.missing.length?' (the app is still being set up)':''}. Ask your Control Tower admin to add you in <b>Team access</b>.</p><button class="ms-btn" data-signout="1">Sign out</button>`);
  if(STATE==='setup')return setupScreen();
  return'';
}
function renderNav(){let h='',m='';
  for(const sec of SECTIONS){const items=SCREENS.filter(s=>s[2]===sec&&canSee(s[0]));if(!items.length)continue;h+=`<div class="nav-sec">${esc(sec)}</div>`;
    for(const[k,,,c,ic]of items){const t=TITLE(k);h+=`<a href="#${k}" data-p="${k}" class="${k===UI.page?'on':''}"><i style="background:color-mix(in srgb,var(${c}) 22%,transparent);color:var(${c})"><svg viewBox="0 0 24 24" fill="currentColor">${ic}</svg></i>${esc(t)}</a>`;m+=`<a href="#${k}" class="${k===UI.page?'on':''}">${esc(t)}</a>`}}
  $('#nav').innerHTML=h;$('#mobnav').innerHTML=m}
function renderMe(){const e=ME.empId&&E()[ME.empId];
  $('#me').innerHTML=`<div class="who"><div style="min-width:0"><strong>${esc(e?e.name:ME.name)}</strong><small style="display:block;word-break:break-all;color:var(--amber-2)">${esc(ME.email)}</small><small>${esc(e?e.designation+' · ':'')}${esc(ME.role||'')}</small></div></div><div style="display:flex;gap:6px;margin-top:10px">${DEMO?'':'<button class="btn sm" data-refresh="1" style="flex:1">Refresh</button>'}<button class="btn sm" data-signout="1" style="flex:1">Sign out</button></div>`}
function head(title,sub,acts=''){const sec=SC[UI.page]?.[2]||'';return`<div class="top"><div><div class="eyebrow">${esc(sec)}</div><h2>${esc(title)}</h2>${sub?`<p>${sub}</p>`:''}</div><div class="acts">${acts}</div></div>`}
const addBtn=(screen,what,label)=>canEdit(screen)?`<button class="btn pri" data-add="${what}">+ ${esc(label)}</button>`:'';
const expBtn=what=>`<button class="btn" data-export="${what}">Export CSV</button>`;
const roNote=s=>canEdit(s)?'':`<div class="note" style="margin-bottom:12px"><b>View only.</b> Your role can view this page. Ask a lead or admin if something here needs changing.</div>`;

/* ---------- table engine ---------- */
function table(key,cols,list,opts={}){
  const f=UI.f[key]||{};const st=UI.sort[key]||opts.sort||{};let r=list;
  if(f.q){const q=ttxt(f.q);r=r.filter(x=>ttxt(opts.hay?opts.hay(x):JSON.stringify(x)).includes(q))}
  for(const fl of opts.filters||[]){const v=f[fl.k];if(v)r=r.filter(x=>fl.test(x,v))}
  if(st.k){const c=cols.find(c=>c.k===st.k);if(c){const g=c.sv||c.v||(x=>x[c.k]);r=[...r].sort((a,b)=>{const A=g(a),B=g(b);return(A>B?1:A<B?-1:0)*(st.d==='desc'?-1:1)})}}
  UI.last=UI.last||{};UI.last[key]=r;
  const tools=`<div class="tools"><input type="search" id="q-${key}" data-f="${key}" data-fk="q" placeholder="Search…" value="${esc(f.q||'')}" aria-label="Search">${(opts.filters||[]).map(fl=>`<select id="f-${key}-${fl.k}" data-f="${key}" data-fk="${fl.k}" aria-label="${esc(fl.label)}"><option value="">${esc(fl.label)}: all</option>${fl.opts.map(o=>{const[v,l]=Array.isArray(o)?o:[o,o];return`<option value="${esc(v)}"${f[fl.k]===v?' selected':''}>${esc(l)}</option>`}).join('')}</select>`).join('')}<span class="count">${r.length} of ${list.length}</span></div>`;
  const th=cols.map(c=>`<th class="${c.num?'num':''}${c.nos?' nos':''}" ${c.nos?'':`data-sort="${key}" data-k="${c.k}"`}>${esc(c.t)}${st.k===c.k?`<span class="ar">${st.d==='desc'?'↓':'↑'}</span>`:''}</th>`).join('');
  const body=r.length?r.map(x=>`<tr ${opts.edit?`data-row="${esc(opts.edit)}" data-id="${esc(x.id)}" tabindex="0"`:''}>${cols.map(c=>`<td class="${c.num?'num ':''}${c.cls||''}">${c.h?c.h(x):esc(c.v?c.v(x):x[c.k])}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${cols.length}"><div class="empty">${opts.empty||'Nothing matches these filters.'}</div></td></tr>`;
  return tools+`<div class="tw"><table><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>`;
}
const empOpts=()=>Object.values(E()).sort((a,b)=>a.name.localeCompare(b.name)).map(e=>[e.id,e.name]);
const prjOpts=()=>Object.values(P()).sort((a,b)=>a.id.localeCompare(b.id)).map(p=>[p.id,`${p.id} — ${p.program} · ${p.project}${p.goal?' · '+p.goal:''}`]);

/* =========================================================
   DERIVED METRICS
   ========================================================= */
function allocBy(){const byE={},byP={};for(const a of Object.values(AL())){(byE[a.empId]??=[]).push(a);(byP[a.projectId]??=[]).push(a)}return{byE,byP}}
function billingOf(empId,byE){return(byE[empId]||[]).filter(a=>a.type==='Billable'&&isClient(a.projectId)&&ACTIVE_ST.includes(prj(a.projectId)?.status))}
function leaveHours(empId,from,to){let h=0;for(const l of Object.values(LV())){if(l.empId!==empId||l.date<from||l.date>to||isWE(l.date))continue;h+=l.kind==='Leave'?H().day:l.kind==='Half-day leave'?H().half:l.kind==='Permission'?H().perm:0}return h}
function workdays(from,to){let n=0;for(let d=from;d<=to;d=addD(d,1))if(!isWE(d))n++;return n}
function totalYoE(e){if(!e.doj)return +e.yoePrior||0;return(+e.yoePrior||0)+Math.max(0,daysBetween(e.doj,TODAY())/365.25)}
function overdue(a){return a.status!=='Closed'&&a.target&&a.target<TODAY()}
function period(k){const t=TODAY();const d=parse(t);
  if(k==='week')return{from:monday(t),to:t,label:'This week'};
  if(k==='last'){const a=new Date(d.getFullYear(),d.getMonth()-1,1),b=new Date(d.getFullYear(),d.getMonth(),0);return{from:isoD(a),to:isoD(b),label:MON[a.getMonth()]+' '+a.getFullYear()}}
  if(k==='quarter'){const q=Math.floor(d.getMonth()/3)*3;return{from:isoD(new Date(d.getFullYear(),q,1)),to:t,label:'Quarter to date'}}
  if(k==='30')return{from:addD(t,-29),to:t,label:'Last 30 days'};
  return{from:isoD(new Date(d.getFullYear(),d.getMonth(),1)),to:t,label:MON[d.getMonth()]+' '+d.getFullYear()+' to date'}}
/* hours model: possible billable vs actual billable vs non-billable split */
function hoursModel(from,to){
  const {byE}=allocBy();const emps=Object.values(E()).filter(e=>e.status!=='Exited');const wd=workdays(from,to);
  const billKey=new Set(Object.values(AL()).filter(a=>a.type==='Billable').map(a=>a.empId+'|'+a.projectId));
  const per={};for(const e of emps){const bl=billingOf(e.id,byE);const share=Math.min(100,bl.reduce((s,a)=>s+(+a.alloc||0),0))/100;const cap=Math.max(0,wd*H().day-leaveHours(e.id,from,to));
    per[e.id]={e,billable:bl.length>0,cap,possible:bl.length?cap*share:0,actual:0,indirect:0,initiative:0,internal:0,total:0,sup:{}}}
  for(const x of Object.values(TS())){if(x.date<from||x.date>to)continue;const p=per[x.empId];if(!p)continue;const h=+x.hours||0;p.total+=h;
    if(billKey.has(x.empId+'|'+x.projectId))p.actual+=h;
    else if(isClient(x.projectId)){p.indirect+=h;p.sup[x.projectId]=(p.sup[x.projectId]||0)+h}
    else if(/^NVTX_INTL/.test(x.projectId))p.initiative+=h;else p.internal+=h}
  const sum=k=>Object.values(per).reduce((s,p)=>s+p[k],0);
  return{per,possible:sum('possible'),actual:sum('actual'),indirect:sum('indirect'),initiative:sum('initiative'),internal:sum('internal'),total:sum('total'),wd}
}

/* =========================================================
   VIEWS
   ========================================================= */
const VIEWS={};

VIEWS.dashboard=()=>{
  const emps=Object.values(E()).filter(e=>e.status!=='Exited');const {byE,byP}=allocBy();const t=TODAY();
  const pr=period(UI.per);const hm=hoursModel(pr.from,pr.to);
  const billable=emps.filter(e=>billingOf(e.id,byE).length);const nonBill=emps.filter(e=>!billingOf(e.id,byE).length);
  const cnt=(f)=>emps.filter(f).length;
  const comp=[['Team Leads',cnt(e=>e.orgRole==='Team Lead')],['Core Team',cnt(e=>e.orgRole==='Core Team')],['PMO',cnt(e=>e.orgRole==='PMO')],['Project Leads',cnt(e=>/Lead/.test(e.projectRole||''))],['Team Members',cnt(e=>e.orgRole==='Team Member')]];
  const acts=Object.values(AC());const open=acts.filter(a=>a.status!=='Closed');const od=open.filter(overdue);
  const activeClient=Object.values(P()).filter(p=>isClient(p.id)&&p.status==='Active');
  // support map
  const sup=activeClient.map(p=>{const bill=(byP[p.id]||[]).filter(a=>a.type==='Billable').map(a=>a.empId);const s={};
    for(const a of byP[p.id]||[])if(a.type!=='Billable'&&!bill.includes(a.empId))(s[a.empId]??={h:0,acts:0,role:a.role});
    for(const a of open)if(a.projectId===p.id)for(const who of[a.ownerId,a.backup1,a.backup2])if(who&&!bill.includes(who))(s[who]??={h:0,acts:0,role:''}).acts++;
    for(const id in hm.per){const h=hm.per[id].sup[p.id];if(h&&!bill.includes(id))(s[id]??={h:0,acts:0,role:''}).h+=h}
    return{p,bill,s}}).filter(x=>x.bill.length||Object.keys(x.s).length);
  const nbSupport=nonBill.filter(e=>sup.some(x=>x.s[e.id]));
  const nbInd=nonBill.reduce((s,e)=>s+(hm.per[e.id]?.indirect||0),0);
  const nonBillH=hm.indirect+hm.initiative+hm.internal;
  const mx=Math.max(1,hm.possible,hm.actual+0,nonBillH);
  // leave for period
  const lv=Object.values(LV()).filter(l=>l.date>=pr.from&&l.date<=pr.to);
  const lk=k=>lv.filter(l=>l.kind===k).length;
  const lper={};for(const l of lv){const o=(lper[l.empId]??={L:0,H:0,W:0,P:0,n:0});o[l.kind==='Leave'?'L':l.kind==='Half-day leave'?'H':l.kind==='WFH'?'W':'P']++;if(l.clientInformed==='No')o.n++}
  const outToday=Object.values(LV()).filter(l=>l.date===t);
  // alerts
  const al=[];
  for(const p of activeClient){if(p.sowEnd&&p.sowEnd<t)al.push(['bad',`<b>${esc(p.program)} · ${esc(prjShort(p.id))}</b> is Active but its SOW ended ${fd(p.sowEnd)}.`,'projects']);else if(p.sowEnd&&daysBetween(t,p.sowEnd)<=45)al.push(['warn',`<b>${esc(p.program)} · ${esc(prjShort(p.id))}</b> SOW ends ${fd(p.sowEnd)}.`,'projects'])}
  for(const u of Object.values(UID())){if(!u.expiry)continue;const d=daysBetween(t,u.expiry);if(d<0&&u.status==='Active')al.push(['bad',`${esc(empName(u.empId))}'s <b>${esc(u.program)}</b> login expired ${fd(u.expiry)} but is marked Active.`,'userids']);else if(d>=0&&d<=30)al.push(['warn',`${esc(empName(u.empId))}'s <b>${esc(u.program)}</b> login expires ${fd(u.expiry)}.`,'userids'])}
  const noAlloc=emps.filter(e=>!(byE[e.id]||[]).length);if(noAlloc.length)al.push(['warn',`No allocation: ${noAlloc.map(e=>esc(e.name)).join(', ')}.`,'allocation']);
  if(od.length)al.push(['info',`${od.length} action items are past their target date.`,'activities']);
  const go=s=>canSee(s)?s:'';
  const per=`<div class="seg" role="group" aria-label="Period">${[['week','Week'],['month','Month'],['last','Last month'],['quarter','Quarter'],['30','30 days']].map(([k,l])=>`<button data-per="${k}" class="${UI.per===k?'on':''}">${l}</button>`).join('')}</div>`;
  const hb=(lbl,v,col,sub)=>`<div class="hrow"><span>${lbl}${sub?`<div class="sub">${sub}</div>`:''}</span><span class="trk"><span style="width:${v/mx*100}%;background:${col}"></span></span><span class="v">${n1(v)} h</span></div>`;
  return `${ASSETS.title?`<div class="titlecard"><img src="${ASSETS.title}" alt="Delivery Management Control Tower — People, Projects, Initiatives, Utilization, Productivity"></div>`:''}
  <div class="kstrip">
    <div class="kt hl"><b>${emps.length}</b><span>People in the team</span></div>
    ${comp.map(([k,v])=>`<div class="kt"><b>${v}</b><span>${k}</span></div>`).join('')}
    <div class="kt"><b style="color:var(--ok)">${billable.length}</b><span>Billable</span></div>
    <div class="kt"><b style="color:var(--warn)">${nonBill.length}</b><span>Non-billable · ${nbSupport.length} supporting clients</span></div>
  </div>
  <div class="grid g3">
   <div class="card span2"><div class="hd"><h3>Billing hours</h3>${per}</div><div class="bd">
     <p class="sub" style="margin:0 0 12px">${esc(pr.label)} · ${hm.wd} working days · ${H().day} h/day, less leave and permission</p>
     <div class="hbars">
      ${hb('Possible billable hours',hm.possible,'var(--c-team)',`${billable.length} billable people × capacity × billable %`)}
      ${hb('Actual billable hours',hm.actual,'var(--c-proj)',`${pct(hm.actual,hm.possible)}% of possible`)}
      ${hb('Non-billable hours',nonBillH,'var(--amber)',`${n1(hm.total)} h logged in total`)}
     </div>
     <div style="margin-top:16px;font-size:12px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.5px">Where non-billable hours went</div>
     <div class="trk" style="display:flex;height:14px;border-radius:7px;overflow:hidden;background:var(--card-2);margin:8px 0">${nonBillH?[[hm.indirect,'var(--c-leave)'],[hm.initiative,'var(--c-act)'],[hm.internal,'var(--faint)']].map(([v,c])=>`<span style="width:${v/nonBillH*100}%;background:${c}"></span>`).join(''):''}</div>
     <div class="legend" style="font-size:12.5px"><span><i style="background:var(--c-leave)"></i>Indirect support to customers <b>${n1(hm.indirect)} h</b></span><span><i style="background:var(--c-act)"></i>Org initiatives <b>${n1(hm.initiative)} h</b></span><span><i style="background:var(--faint)"></i>Internal &amp; common <b>${n1(hm.internal)} h</b></span></div>
     ${hm.total?'':'<p class="sub" style="margin-top:10px">No hours logged in this period yet. Figures fill in as people use the Productivity tracker.</p>'}
   </div></div>
   <div class="card"><div class="hd"><h3>Needs attention</h3><small>${al.length}</small></div><div class="bd"><div class="alerts">${al.length?al.slice(0,10).map(([s,m,g])=>`<div class="alert"><span class="dot" style="background:var(--${s==='bad'?'bad':s==='warn'?'warn':'info'})"></span><div>${m}</div>${go(g)?`<a data-go="${g}">Open</a>`:'<span></span>'}</div>`).join(''):'<div class="empty">All clear.</div>'}</div></div></div>

   <div class="card" style="grid-column:1/-1"><div class="hd"><h3>Who supports the billable work</h3><small>${nbSupport.length} of ${nonBill.length} non-billable people support customer projects · ${n1(nbInd)} h indirect in ${esc(pr.label.toLowerCase())}</small></div><div class="bd">
     <div class="support">${sup.map(({p,bill,s})=>`<div class="sp"><h4>${esc(p.program)} · ${esc(prjShort(p.id))}</h4><div class="sub">${esc(p.goal||'')}</div>
       <div class="lane"><span>Billable</span><span>${bill.length?bill.map(id=>`<span class="chip g">${esc(empName(id))}</span>`).join(''):'<span class="sub">—</span>'}</span></div>
       <div class="lane"><span>Supported by</span><span>${Object.keys(s).length?Object.entries(s).sort((a,b)=>b[1].h-a[1].h).map(([id,o])=>{const nb=!billingOf(id,byE).length;return`<span class="chip ${nb?'o':''}" title="${nb?'Non-billable':'Billable elsewhere'}${o.role?' · '+esc(o.role):''}${o.acts?' · '+o.acts+' open actions':''}">${esc(empName(id))}${o.h?` · ${n1(o.h)} h`:''}</span>`}).join(''):'<span class="sub">No support recorded</span>'}</span></div></div>`).join('')}</div>
     <div class="legend"><span><i style="background:var(--ok)"></i>Billable to the customer</span><span><i style="background:var(--warn)"></i>Non-billable, supporting indirectly</span><span><i style="background:var(--line)"></i>Billable elsewhere, also supporting</span></div>
   </div></div>

   <div class="card span2"><div class="hd"><h3>Org structure</h3><small>Click a lead to drill into their team</small></div><div class="bd">${orgTree(byE)}</div></div>
   <div class="card"><div class="hd"><h3>Leave &amp; permission</h3><small>${esc(pr.label)}</small></div><div class="bd">
     <div class="kstrip" style="grid-template-columns:repeat(2,minmax(0,1fr));margin-bottom:10px">${[['Leave days',lk('Leave'),'bad'],['Half days',lk('Half-day leave'),'warn'],['Permissions',lk('Permission'),''],['WFH days',lk('WFH'),'info']].map(([k,v,c])=>`<div class="kt"><b style="${c?`color:var(--${c})`:''}">${v}</b><span>${k}</span></div>`).join('')}</div>
     <div class="sub" style="margin-bottom:6px">Out today: ${outToday.length?outToday.map(l=>`<span class="chip">${esc(first(l.empId))} · ${esc(l.kind==='Half-day leave'?'½ day':l.kind)}</span>`).join(''):'everyone in'}</div>
     <div class="tw" style="max-height:260px;overflow:auto"><table><thead><tr><th class="nos">Person</th><th class="nos num">L</th><th class="nos num">½</th><th class="nos num">P</th><th class="nos num">WFH</th></tr></thead><tbody>${Object.entries(lper).sort((a,b)=>(b[1].L+b[1].H/2)-(a[1].L+a[1].H/2)).map(([id,o])=>`<tr><td>${esc(empName(id))}${o.n?` <span class="pill warn" title="Client not informed">${o.n} not informed</span>`:''}</td><td class="num">${o.L}</td><td class="num">${o.H}</td><td class="num">${o.P}</td><td class="num">${o.W}</td></tr>`).join('')||'<tr><td colspan="5"><div class="empty">No leave in this period.</div></td></tr>'}</tbody></table></div>
   </div></div>

   <div class="card" style="grid-column:1/-1"><div class="hd"><h3>Hours by person</h3><small>${esc(pr.label)}</small></div><div class="bd" style="padding-top:8px">${table('hrs',[
      {k:'n',t:'Team member',sv:x=>x.e.name,h:x=>`<b>${esc(x.e.name)}</b><div class="sub">${esc(x.e.designation)}</div>`},
      {k:'b',t:'Billing',sv:x=>x.billable?1:0,h:x=>x.billable?stPill('Billable'):stPill('Shadow','Non-billable')},
      {k:'possible',t:'Possible billable',num:1,v:x=>x.billable?n1(x.possible):'—'},
      {k:'actual',t:'Actual billable',num:1,v:x=>n1(x.actual)},
      {k:'indirect',t:'Indirect support',num:1,v:x=>n1(x.indirect)},
      {k:'initiative',t:'Initiatives',num:1,v:x=>n1(x.initiative)},
      {k:'internal',t:'Internal',num:1,v:x=>n1(x.internal)},
      {k:'total',t:'Total logged',num:1,v:x=>n1(x.total)}],Object.values(hm.per).map(p=>({...p,id:p.e.id})),{sort:{k:'b',d:'desc'},hay:x=>x.e.name,filters:[{k:'b',label:'Billing',opts:[['y','Billable'],['n','Non-billable']],test:(x,v)=>x.billable===(v==='y')}]})}</div></div>
  </div>`;
};
function orgTree(byE){
  const emps=Object.values(E()).filter(e=>e.status!=='Exited');const ids=new Set(emps.map(e=>e.id));const kids={};
  for(const e of emps){const m=e.managerId&&ids.has(e.managerId)&&e.managerId!==e.id?e.managerId:'_root';(kids[m]??=[]).push(e)}
  const size=id=>(kids[id]||[]).reduce((s,k)=>s+1+size(k.id),0);
  const line=e=>{const b=billingOf(e.id,byE).length;return`<span class="nm">${esc(e.name)}</span><span class="sub">${esc(e.designation)}</span>${e.orgRole&&e.orgRole!=='Team Member'?`<span class="chip b">${esc(e.orgRole)}</span>`:''}${/Lead/.test(e.projectRole||'')?`<span class="chip">${esc(e.projectRole)}</span>`:''}<span class="pill ${b?'ok':'warn'}">${b?'Billable':'Non-billable'}</span>`};
  const node=(e,depth)=>{const ch=(kids[e.id]||[]).sort((a,b)=>size(b.id)-size(a.id)||a.name.localeCompare(b.name));
    if(!ch.length)return`<div class="leaf">${line(e)}</div>`;
    return`<details ${depth<1?'open':''}><summary><span class="car">▸</span>${line(e)}<span class="cnt">${ch.length} direct · ${size(e.id)} total</span></summary>${ch.map(c=>node(c,depth+1)).join('')}</details>`};
  const roots=(kids._root||[]).sort((a,b)=>size(b.id)-size(a.id));
  return`<div class="tree">${roots.map(r=>node(r,0)).join('')}</div>`;
}

VIEWS.team=()=>{
  const {byE}=allocBy();
  const list=Object.values(E()).map(e=>({...e,_y:totalYoE(e),_bill:billingOf(e.id,byE).length?'Billable':'Non-billable',_acc:ACC()[(e.email||'').trim().toLowerCase()]}));
  const cols=[
    {k:'id',t:'ID',cls:'mono'},
    {k:'name',t:'Name',h:x=>`<b>${esc(x.name)}</b><div class="sub">${esc(x.designation)}</div>`},
    {k:'orgRole',t:'Org role'},{k:'projectRole',t:'Project role'},
    {k:'managerId',t:'Reports to',v:x=>x.managerId?empName(x.managerId):x.managerName},
    {k:'doj',t:'Joined',v:x=>fd(x.doj),sv:x=>x.doj},
    {k:'_y',t:'Total YoE',num:1,v:x=>n1(x._y)},
    {k:'_bill',t:'Billing',h:x=>(x._bill==='Billable'?stPill('Billable'):stPill('Shadow','Non-billable'))+(x.billableType?`<div class="sub">${esc(x.billableType)}</div>`:'')},
    {k:'status',t:'Status',h:x=>stPill(x.status)},
    ...cfCols('employees'),
    {k:'_acc',t:'App access',sv:x=>x._acc?.role||'',h:x=>x._acc&&x._acc.active!=='No'?stPill(x._acc.role):'<span class="pill">No access</span>'}
  ];
  return head(TITLE('team'),'Everyone in Product Engineering with their Org Role and Project Role. Update this whenever someone joins, moves or leaves.',addBtn('team','emp','Add team member')+expBtn('emp'))+roNote('team')+
  table('emp',cols,list,{edit:canEdit('team')?'emp':'',sort:{k:'name'},filters:[
    {k:'orgRole',label:'Org role',opts:opt(list.map(x=>x.orgRole)),test:(x,v)=>x.orgRole===v},
    {k:'mgr',label:'Manager',opts:opt(list.map(x=>x.managerId)).map(id=>[id,empName(id)]),test:(x,v)=>x.managerId===v},
    {k:'bill',label:'Billing',opts:['Billable','Non-billable'],test:(x,v)=>x._bill===v}]});
};

VIEWS.projects=()=>{
  const {byP}=allocBy();const acts=Object.values(AC());const t=TODAY();
  const list=Object.values(P()).map(p=>({...p,_n:(byP[p.id]||[]).length,_open:acts.filter(a=>a.projectId===p.id&&a.status!=='Closed').length}));
  const sow=p=>{if(!isClient(p.id))return'';if(!p.sowEnd)return'<span class="pill">No SOW dates</span>';const d=daysBetween(t,p.sowEnd);if(d<0)return`<span class="pill ${p.status==='Active'?'bad':''}">Ended ${fds(p.sowEnd)}</span>`;return`<span class="pill ${d<=45?'warn':'ok'}">${fd(p.sowEnd)}</span>`};
  const cols=[{k:'id',t:'ID',cls:'mono'},{k:'program',t:'Customer / program',h:x=>`<b>${esc(x.program)}</b><div class="sub">${esc(x.project)}</div>`},{k:'goal',t:'Goal',cls:'clip'},{k:'status',t:'Status',h:x=>stPill(x.status)},{k:'contract',t:'Contract',h:x=>`${esc(x.contract)}<div class="sub">${esc(x.revenue)}</div>`},{k:'start',t:'Start',v:x=>fd(x.start),sv:x=>x.start},{k:'sowEnd',t:'SOW end',h:sow,sv:x=>x.sowEnd},{k:'_n',t:'People',num:1,h:x=>x._n?String(x._n):'<span class="sub">All</span>'},{k:'_open',t:'Open actions',num:1},...cfCols('projects')];
  const tab=UI.tab.prj||'client';const shown=list.filter(x=>tab==='all'||(tab==='client'?isClient(x.id):!isClient(x.id)));
  return head(TITLE('projects'),'Every customer project handled or running, and every organisation initiative planned or under way.',addBtn('projects','prj','Add project / initiative')+expBtn('prj'))+roNote('projects')+
  `<div class="tabs">${[['client','Customer projects'],['init','Org initiatives'],['all','All']].map(([k,l])=>`<button data-tab="prj" data-v="${k}" class="${tab===k?'on':''}">${l} <span class="sub">${list.filter(x=>k==='all'||(k==='client'?isClient(x.id):!isClient(x.id))).length}</span></button>`).join('')}</div>`+
  table('prj',cols,shown,{edit:canEdit('projects')?'prj':'',sort:{k:'id'},filters:[{k:'program',label:'Program',opts:opt(shown.map(x=>x.program)),test:(x,v)=>x.program===v},{k:'status',label:'Status',opts:opt(shown.map(x=>x.status)),test:(x,v)=>x.status===v}]});
};

VIEWS.allocation=()=>{
  const list=Object.values(AL()).map(a=>({...a,_p:prjName(a.projectId),_e:empName(a.empId)}));const {byP}=allocBy();const common=Object.values(P()).filter(p=>!(byP[p.id]||[]).length);
  const cols=[{k:'projectId',t:'Project / initiative',h:x=>`<span class="mono">${esc(x.projectId)}</span><div><b>${esc(x._p)}</b></div><div class="sub">${esc(prj(x.projectId)?.goal||'')}</div>`},{k:'_e',t:'Resource',h:x=>`<b>${esc(x._e)}</b><div class="sub mono">${esc(x.empId)}</div>`},{k:'role',t:'Role'},{k:'alloc',t:'Allocation',num:1,v:x=>(+x.alloc||0)+'%'},{k:'type',t:'Type',h:x=>stPill(x.type)},{k:'remarks',t:'Remarks',cls:'clip'},...cfCols('alloc')];
  const load={};for(const a of list)if(a.type==='Billable')load[a.empId]=(load[a.empId]||0)+(+a.alloc||0);const over=Object.entries(load).filter(([,v])=>v>100);
  return head(TITLE('allocation'),'Who is responsible for each project and initiative, in what role and on what billing basis. A project with nobody named is common to the whole team.',addBtn('allocation','alloc','Allocate resource')+expBtn('alloc'))+roNote('allocation')+
  (over.length?`<div class="note" style="margin-bottom:12px"><b>Billable allocation above 100%:</b> ${over.map(([e,v])=>`${esc(empName(e))} (${v}%)`).join(', ')}. Possible billable hours are capped at 100% per person.</div>`:'')+
  `<div class="card" style="margin-bottom:14px"><div class="hd"><h3>Common for all team members</h3><small>${common.length} items anyone can log time against</small></div><div class="bd">${common.map(p=>`<span class="chip" title="${esc(p.id)}">${esc(p.goal||p.project)}</span>`).join('')}</div></div>`+
  table('alloc',cols,list,{edit:canEdit('allocation')?'alloc':'',sort:{k:'projectId'},hay:x=>x.projectId+x._p+x._e+x.role,filters:[{k:'p',label:'Project',opts:opt(list.map(x=>x.projectId)).map(id=>[id,id+' — '+prjName(id)]),test:(x,v)=>x.projectId===v},{k:'e',label:'Resource',opts:opt(list.map(x=>x.empId)).map(id=>[id,empName(id)]),test:(x,v)=>x.empId===v},{k:'type',label:'Type',opts:['Billable','Shadow'],test:(x,v)=>x.type===v},{k:'kind',label:'Kind',opts:[['c','Customer projects'],['i','Org initiatives']],test:(x,v)=>(v==='c')===isClient(x.projectId)}]});
};

VIEWS.activities=()=>{
  const list=Object.values(AC()).map(a=>({...a,_p:prjName(a.projectId),_o:a.ownerId?empName(a.ownerId):'Common for all'}));const me=ME.empId;
  const tgt=x=>{if(!x.target)return'—';const bad=x.start&&x.target<x.start;return`<span class="${overdue(x)?'pill bad':''}" ${bad?'title="Target is before start"':''}>${fd(x.target)}${bad?' ⚠':''}</span>`};
  const cols=[{k:'actId',t:'Activity',cls:'mono'},{k:'desc',t:'Action item',h:x=>`<b style="font-weight:500">${esc(x.desc)}</b><div class="sub">${esc(x._p)} · ${esc(x.frequency||'')}</div>`},{k:'_o',t:'Owner',h:x=>`${esc(x._o)}${x.backup1?`<div class="sub">Backup: ${esc(empName(x.backup1))}${x.backup2?', '+esc(empName(x.backup2)):''}</div>`:''}`},{k:'priority',t:'Priority',h:x=>`<span class="pill ${x.priority==='High'?'bad':x.priority==='Medium'?'warn':''}">${esc(x.priority||'—')}</span>`},{k:'status',t:'Status',h:x=>stPill(x.status)},{k:'target',t:'Target',h:tgt,sv:x=>x.target},{k:'remarks',t:'Remarks',cls:'clip sub'},...cfCols('activities')];
  const c=f=>list.filter(f).length;
  return head(TITLE('activities'),'Action items for each project and initiative, with owners and backups. Everyone logs productivity against these.',addBtn('activities','act','Add action item')+expBtn('act'))+roNote('activities')+
  `<div class="kstrip">${[['Open',c(x=>x.status==='Open')],['In progress',c(x=>x.status==='In-Progress')],['On hold',c(x=>x.status==='Hold')],['Overdue',c(overdue)],['Closed',c(x=>x.status==='Closed')]].map(([k,v])=>`<div class="kt"><b>${v}</b><span>${k}</span></div>`).join('')}</div>`+
  table('act',cols,list,{edit:canEdit('activities')?'act':'',sort:{k:'target'},hay:x=>x.actId+x.desc+x._p+x._o+x.remarks,filters:[{k:'mine',label:'Show',opts:[...(me?[['me','Assigned to me']]:[]),['od','Overdue only'],['open','Not closed']],test:(x,v)=>v==='me'?(x.ownerId===me||x.backup1===me||x.backup2===me):v==='od'?overdue(x):x.status!=='Closed'},{k:'p',label:'Project',opts:opt(list.map(x=>x.projectId)).map(id=>[id,id+' — '+prjName(id)]),test:(x,v)=>x.projectId===v},{k:'o',label:'Owner',opts:opt(list.map(x=>x.ownerId)).map(id=>[id,empName(id)]),test:(x,v)=>x.ownerId===v||x.backup1===v||x.backup2===v},{k:'status',label:'Status',opts:L('actStatus'),test:(x,v)=>x.status===v}]});
};

VIEWS.productivity=()=>{
  const lvl=perm('productivity');const me=ME.empId;const t=TODAY();
  const tabsAvail=[['mine','My tracker'],...(lvl==='edit'||lvl==='view'?[['team','Team log']]:[]),...(lvl==='edit'?[['util','Utilization']]:[])];
  let tab=UI.tab.prod||'mine';if(!tabsAvail.some(x=>x[0]===tab))tab='mine';
  const tabs=`<div class="tabs">${tabsAvail.map(([k,l])=>`<button data-tab="prod" data-v="${k}" class="${tab===k?'on':''}">${l}</button>`).join('')}</div>`;
  const acts=lvl==='edit'&&tab!=='mine'?`<button class="btn pri" data-add="tsadmin">+ Log for a team member</button>${expBtn('ts')}`:(lvl==='view'?expBtn('ts'):'');
  let body='';
  if(tab==='mine')body=me?myTracker(me,t):`<div class="note">Your sign-in (${esc(ME.email)}) isn't linked to a record in the Employee database. Ask an admin to set your Employee ID in <b>Team access</b>, or to correct your email in the Employee database.</div>`;
  else if(tab==='team'){const list=Object.values(TS()).map(x=>({...x,_e:empName(x.empId)}));
    const cols=[{k:'date',t:'Date',v:x=>fd(x.date),sv:x=>x.date},{k:'_e',t:'Team member'},{k:'actId',t:'Activity',h:x=>`<span class="mono">${esc(x.actId)}</span><div>${esc(x.actDesc)}</div><div class="sub">${esc(prjName(x.projectId))}</div>`},{k:'sub',t:'Sub activity',cls:'clip'},{k:'hours',t:'Hours',num:1},{k:'remarks',t:'Remarks',cls:'clip sub'}];
    body=table('ts',cols,list,{edit:lvl==='edit'?'tsadmin':'',sort:{k:'date',d:'desc'},hay:x=>x._e+x.actId+x.actDesc+x.sub+x.remarks,filters:[{k:'e',label:'Team member',opts:opt(list.map(x=>x.empId)).map(id=>[id,empName(id)]),test:(x,v)=>x.empId===v},{k:'r',label:'Period',opts:[['7','Last 7 days'],['30','Last 30 days'],['m','This month']],test:(x,v)=>v==='m'?x.date.slice(0,7)===t.slice(0,7):x.date>=addD(t,-+v)}],empty:'No hours logged yet.'})}
  else{const wk=UI.week||monday(t);const end=addD(wk,4);const to=end<t?end:t;const hm=hoursModel(wk,to<wk?wk:to);
    const list=Object.values(hm.per).map(p=>({id:p.e.id,name:p.e.name,exp:p.cap,h:p.total,cl:p.actual+p.indirect,pct:p.cap?p.total/p.cap*100:0}));
    const cols=[{k:'name',t:'Team member',h:x=>`<b>${esc(x.name)}</b>`},{k:'exp',t:'Expected',num:1,v:x=>n1(x.exp)},{k:'h',t:'Logged',num:1,v:x=>n1(x.h)},{k:'cl',t:'On client work',num:1,v:x=>n1(x.cl)},{k:'pct',t:'Logged vs expected',h:x=>`<div style="display:flex;gap:8px;align-items:center;min-width:140px"><span style="flex:1;height:8px;border-radius:5px;background:var(--card-2);overflow:hidden;display:block"><span style="display:block;height:100%;width:${Math.min(100,x.pct)}%;background:${x.pct>=90?'var(--ok)':x.pct>=50?'var(--amber)':'var(--bad)'}"></span></span><span class="sub" style="width:38px;text-align:right">${Math.round(x.pct)}%</span></div>`}];
    body=`<div class="tools"><button class="btn sm" data-wk="-7">← Previous week</button><b>${fd(wk)} – ${fd(end)}</b><button class="btn sm" data-wk="7">Next week →</button></div>`+table('util',cols,list,{sort:{k:'pct'}})}
  return head(TITLE('productivity'),'Log the hours you spend on the action items assigned to you, plus the common items open to the whole team.',acts)+tabs+body;
};
function myTracker(me,t){
  const wk=UI.week||monday(t);const day=UI.day||t;const mine=Object.values(TS()).filter(x=>x.empId===me);
  const days=[];for(let i=0;i<7;i++)days.push(addD(wk,i));
  const acts=Object.values(AC());const assigned=acts.filter(a=>a.status!=='Closed'&&(a.ownerId===me||a.backup1===me||a.backup2===me));const common=acts.filter(a=>!a.ownerId);
  const {byE}=allocBy();const myP=new Set((byE[me]||[]).map(a=>a.projectId));const other=acts.filter(a=>a.status!=='Closed'&&myP.has(a.projectId)&&!assigned.includes(a)&&a.ownerId);
  const og=(l,arr)=>arr.length?`<optgroup label="${esc(l)}">${arr.map(a=>`<option value="${esc(a.id)}">${esc(a.actId)} · ${esc((a.desc||'').slice(0,90))} (${esc(prjShort(a.projectId))})</option>`).join('')}</optgroup>`:'';
  const dayE=mine.filter(x=>x.date===day);const dayH=d=>mine.filter(x=>x.date===d).reduce((s,x)=>s+(+x.hours||0),0);const wkH=days.reduce((s,d)=>s+dayH(d),0);
  return`<div class="tools"><button class="btn sm" data-wk="-7">← Previous</button><b>Week of ${fd(wk)}</b><button class="btn sm" data-wk="7">Next →</button><button class="btn sm ghost" data-wk="0">This week</button><span class="count">${n1(wkH)} h this week</span></div>
  <div class="week" style="margin-bottom:14px">${days.map((d,i)=>{const h=dayH(d);const lv=Object.values(LV()).find(l=>l.empId===me&&l.date===d);return`<button class="day${d===day?' on':''}" data-day="${d}"><small>${['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][i]} ${parse(d).getDate()} ${MON[parse(d).getMonth()]}</small><b>${n1(h)}</b> <span class="sub">h</span>${lv?`<small style="color:var(--info)">${esc(lv.kind)}</small>`:''}<div class="meter"><span style="width:${Math.min(100,h/H().day*100)}%"></span></div></button>`}).join('')}</div>
  <div class="grid g2"><div class="card"><div class="hd"><h3>Log effort for ${fd(day)}</h3></div><div class="bd"><form class="form" id="logform">
   <label class="full">Activity<select class="field" id="lg-act" required><option value="">Choose an action item…</option>${og('Assigned to me',assigned)}${og('Common for all',common)}${og('Other items on my projects',other)}</select><span class="hint">${assigned.length} assigned to you.</span></label>
   <label class="full">Sub activity / what you did<input class="field" id="lg-sub" required maxlength="300"></label>
   <label>Date<input class="field" type="date" id="lg-date" value="${day}" required></label><label>Hours<input class="field" type="number" id="lg-h" min="0.25" max="16" step="0.25" value="1" required></label>
   <label class="full">Remarks<input class="field" id="lg-rem" maxlength="300" placeholder="Optional"></label><div class="full"><button class="btn pri" type="submit">Add entry</button></div></form></div></div>
  <div class="card"><div class="hd"><h3>Entries on ${fd(day)}</h3><small>${n1(dayE.reduce((s,x)=>s+(+x.hours||0),0))} h</small></div><div class="bd">${dayE.length?dayE.map(x=>`<div class="alert" style="grid-template-columns:minmax(0,1fr) auto auto"><div><span class="mono">${esc(x.actId)}</span> ${esc(x.actDesc)}<div class="sub">${esc(x.sub)}${x.remarks?' · '+esc(x.remarks):''}</div></div><b>${n1(x.hours)} h</b><button class="btn sm danger" data-delts="${esc(x.id)}">Delete</button></div>`).join(''):'<div class="empty">Nothing logged for this day yet.</div>'}</div></div></div>`;
}

VIEWS.leave=()=>{
  const t=TODAY();const ym=UI.cal||t.slice(0,7);const[y,m]=ym.split('-').map(Number);const nd=new Date(y,m,0).getDate();
  const lv=Object.values(LV());const monthL=lv.filter(l=>l.date.slice(0,7)===ym);const emps=Object.values(E()).sort((a,b)=>a.name.localeCompare(b.name));
  const idx={};for(const l of monthL)(idx[l.empId+l.date]??=[]).push(l);const code=k=>k==='Leave'?'L':k==='Half-day leave'?'H':k==='WFH'?'W':'P';
  const ds=[];for(let d=1;d<=nd;d++)ds.push(`${ym}-${pad(d)}`);const sum=k=>monthL.filter(l=>l.kind===k).length;const tab=UI.tab.leave||'cal';
  const cal=`<div class="tools"><button class="btn sm" data-cal="-1">← Previous</button><b>${MON[m-1]} ${y}</b><button class="btn sm" data-cal="1">Next →</button><span class="count">${sum('Leave')} leave · ${sum('Half-day leave')} half-day · ${sum('WFH')} WFH · ${sum('Permission')} permission</span></div>
   <div class="tw"><table class="cal"><thead><tr><th>Team member</th>${ds.map(d=>`<th class="${isWE(d)?'we':''}${d===t?' today':''}">${parse(d).getDate()}<div class="sub" style="font-size:9.5px">${'SMTWTFS'[parse(d).getDay()]}</div></th>`).join('')}</tr></thead><tbody>${emps.map(e=>`<tr><td>${esc(e.name)}</td>${ds.map(d=>{const x=idx[e.id+d];return`<td class="${isWE(d)?'we':''}${d===t?' today':''}">${x?x.map(l=>`<span class="mk ${code(l.kind)}" title="${esc(l.kind)}${l.clientInformed==='No'?' · client not informed':''}">${code(l.kind)}</span>`).join(''):''}</td>`}).join('')}</tr>`).join('')}</tbody></table></div>
   <div class="legend"><span><span class="mk L">L</span> Leave</span><span><span class="mk H">H</span> Half-day</span><span><span class="mk W">W</span> WFH</span><span><span class="mk P">P</span> Permission</span></div>`;
  const list=lv.map(l=>({...l,_e:empName(l.empId)}));
  const cols=[{k:'date',t:'Date',v:x=>fd(x.date),sv:x=>x.date},{k:'_e',t:'Team member',h:x=>`<b>${esc(x._e)}</b>`},{k:'kind',t:'Type',h:x=>`<span class="mk ${code(x.kind)}">${code(x.kind)}</span> ${esc(x.kind)}`},{k:'clientInformed',t:'Client informed',h:x=>!x.clientInformed?'<span class="sub">n/a</span>':x.clientInformed==='Yes'?'<span class="pill ok">Yes</span>':'<span class="pill warn">No</span>'},{k:'clients',t:'Clients'},{k:'remarks',t:'Remarks',cls:'clip sub'},...cfCols('leave')];
  return head(TITLE('leave'),'Who is on leave, working from home or taking permission, and whether their clients were told.',addBtn('leave','leave','Record leave / WFH')+expBtn('leave'))+roNote('leave')+
  `<div class="tabs">${[['cal','Month view'],['list','All records']].map(([k,l])=>`<button data-tab="leave" data-v="${k}" class="${tab===k?'on':''}">${l}</button>`).join('')}</div>`+
  (tab==='cal'?cal:table('leave',cols,list,{edit:canEdit('leave')?'leave':'',sort:{k:'date',d:'desc'},hay:x=>x._e+x.kind+x.date+(x.remarks||''),filters:[{k:'e',label:'Team member',opts:opt(list.map(x=>x.empId)).map(id=>[id,empName(id)]),test:(x,v)=>x.empId===v},{k:'k',label:'Type',opts:['Leave','Half-day leave','WFH','Permission'],test:(x,v)=>x.kind===v},{k:'mo',label:'Month',opts:opt(list.map(x=>x.date.slice(0,7))).reverse().map(v=>[v,MON[+v.slice(5)-1]+' '+v.slice(0,4)]),test:(x,v)=>x.date.slice(0,7)===v}]}));
};

VIEWS.userids=()=>{
  const t=TODAY();const list=Object.values(UID()).map(u=>({...u,_e:empName(u.empId),_d:u.expiry?daysBetween(t,u.expiry):null}));
  const exp=x=>{if(!x.expiry)return'<span class="sub">No expiry</span>';if(x._d<0)return`<span class="pill bad">Expired ${fds(x.expiry)}</span>`;return`<span class="pill ${x._d<=30?'warn':'ok'}">${fd(x.expiry)}</span>`};
  const cols=[{k:'_e',t:'Team member',h:x=>`<b>${esc(x._e)}</b><div class="sub mono">${esc(x.empId)}</div>`},{k:'program',t:'Customer',h:x=>`<b>${esc(x.program)}</b><div class="sub">${esc(x.project)}</div>`},{k:'projectId',t:'Project ID',cls:'mono'},{k:'sponsor',t:'Sponsor'},{k:'userName',t:'Account name',cls:'mono'},{k:'expiry',t:'Expiry',h:exp,sv:x=>x.expiry||'9999'},{k:'status',t:'Status',h:x=>stPill(x.status)},...cfCols('userids')];
  return head(TITLE('userids'),'Customer system accounts with sponsor, account name and expiry, so access is renewed before it lapses.',addBtn('userids','uid','Add account')+expBtn('uid'))+roNote('userids')+
  table('uid',cols,list,{edit:canEdit('userids')?'uid':'',sort:{k:'expiry'},filters:[{k:'p',label:'Customer',opts:opt(list.map(x=>x.program)),test:(x,v)=>x.program===v},{k:'s',label:'Expiry',opts:[['x','Expired'],['30','Within 30 days'],['ok','Valid']],test:(x,v)=>x._d==null?v==='ok':v==='x'?x._d<0:v==='30'?x._d>=0&&x._d<=30:x._d>30}]});
};

VIEWS.access=()=>{
  const list=Object.values(ACC()).map(a=>({...a,_e:a.empId?E()[a.empId]:null}));const roles=Object.keys(CFG().roles);
  const noAccess=Object.values(E()).filter(e=>e.status!=='Exited'&&!ACC()[(e.email||'').trim().toLowerCase()]);
  const cols=[{k:'name',t:'Name',h:x=>`<b>${esc(x.name||x._e?.name||'')}</b><div class="sub">${esc(x._e?.designation||'')}</div>`},{k:'id',t:'Email (sign-in)',cls:'mono'},{k:'empId',t:'Employee ID',cls:'mono'},{k:'role',t:'Role',h:x=>stPill(x.role)},{k:'active',t:'Access',h:x=>x.active==='No'?'<span class="pill bad">Removed</span>':'<span class="pill ok">Active</span>'}];
  return head(TITLE('access'),'Who can sign in and what each role can see. People sign in with their @novintix.com Microsoft account; add them here when they join and remove them when they leave.',canEdit('access')?`<button class="btn pri" data-add="acc">+ Give access</button>`:'')+
  (noAccess.length&&canEdit('access')?`<div class="note" style="margin-bottom:12px"><b>${noAccess.length} people in the Employee database have no app access:</b> ${noAccess.map(e=>esc(e.name)).join(', ')}. <button class="btn sm" data-grantall="1">Give them access with suggested roles</button></div>`:'')+
  `<div class="card" style="margin-bottom:14px"><div class="hd"><h3>What each role can do</h3><small>${canSee('settings')?'Change this in Settings → Roles':'Set by the admin'}</small></div><div class="bd" style="padding-top:8px">${roleMatrix(false)}</div></div>`+
  table('acc',cols,list,{edit:canEdit('access')?'acc':'',sort:{k:'name'},hay:x=>x.name+x.id+x.role,filters:[{k:'r',label:'Role',opts:roles,test:(x,v)=>x.role===v},{k:'a',label:'Access',opts:[['Yes','Active'],['No','Removed']],test:(x,v)=>(x.active==='No'?'No':'Yes')===v}]});
};
function roleMatrix(editable){const c=CFG();const roles=Object.keys(c.roles);const sc=SCREENS;
  return`<div class="tw matrix" style="border:0"><table><thead><tr><th class="nos">Screen</th>${roles.map(r=>`<th class="nos">${esc(r)}</th>`).join('')}</tr></thead><tbody>${sc.map(([k])=>`<tr><td><b>${esc(TITLE(k))}</b></td>${roles.map(r=>{const v=c.roles[r][k]||'none';if(!editable||r==='Admin')return`<td>${esc((LEVELS.find(l=>l[0]===v)||[])[1]||v)}</td>`;return`<td><select data-rm="${esc(r)}|${k}" aria-label="${esc(r)} ${esc(TITLE(k))}">${LEVELS.filter(l=>l[0]!=='own'||k==='productivity').map(([lv,lb])=>`<option value="${lv}"${v===lv?' selected':''}>${lb}</option>`).join('')}</select></td>`}).join('')}</tr>`).join('')}</tbody></table></div>`}

VIEWS.settings=()=>{
  const c=CFG();const tab=UI.tab.set||'roles';
  const tabs=`<div class="tabs">${[['roles','Roles'],['pages','Pages & names'],['lists','Dropdown lists'],['fields','Custom fields'],['hours','Working hours']].map(([k,l])=>`<button data-tab="set" data-v="${k}" class="${tab===k?'on':''}">${l}</button>`).join('')}</div>`;let b='';
  if(tab==='roles')b=`<div class="card"><div class="hd"><h3>Role permissions</h3><small>Admin always has full access</small></div><div class="bd" style="padding-top:8px">${roleMatrix(true)}<div class="tools" style="margin-top:12px"><input class="field" id="newrole" placeholder="New role name, e.g. Delivery Manager" maxlength="30"><button class="btn" data-addrole="1">Add role</button><span class="sub">Roles in use can't be deleted.</span>${Object.keys(c.roles).filter(r=>!['Admin'].includes(r)&&!Object.values(ACC()).some(a=>a.role===r)).map(r=>`<button class="btn sm danger" data-delrole="${esc(r)}">Delete ${esc(r)}</button>`).join('')}</div><button class="btn pri" data-setsave="roles">Save role permissions</button></div></div>`;
  else if(tab==='pages')b=`<div class="card" style="max-width:720px"><div class="hd"><h3>Rename or hide pages</h3><small>Hidden pages stay visible to Admins</small></div><div class="bd"><div class="form" style="grid-template-columns:minmax(0,1fr) auto">${SCREENS.filter(s=>!['settings','access'].includes(s[0])).map(([k,t])=>`<label>${esc(t)}<input class="field" id="pg-${k}" value="${esc(c.pages[k]?.title||'')}" placeholder="${esc(t)}" maxlength="40"></label><label style="flex-direction:row;align-items:center;gap:6px;padding-top:18px">${k==='dashboard'?'<span class="sub">Always on</span>':`<input type="checkbox" id="pgs-${k}" ${c.pages[k]?.show!==false?'checked':''}> Show`}</label>`).join('')}</div><div style="margin-top:14px"><button class="btn pri" data-setsave="pages">Save pages</button></div></div></div>`;
  else if(tab==='lists')b=`<div class="note" style="margin-bottom:12px">One option per line. Keep <b>Billable</b> in Revenue type, <b>Closed</b> in Activity status and <b>Active</b> in Employment status.</div><div class="grid g3">${Object.entries(LIST_DEF).map(([k,[lbl]])=>`<div class="card"><div class="hd"><h3>${esc(lbl)}</h3></div><div class="bd"><textarea class="field" style="width:100%;min-height:130px" id="ls-${k}" aria-label="${esc(lbl)}">${esc(c.lists[k].join('\n'))}</textarea></div></div>`).join('')}</div><div style="margin-top:14px"><button class="btn pri" data-setsave="lists">Save lists</button></div>`;
  else if(tab==='fields'){const rowsF=Object.entries(CF_TABLES).flatMap(([t,tl])=>(c.custom[t]||[]).map((f,i)=>`<tr><td>${esc(tl)}</td><td><b>${esc(f.label)}</b></td><td>${esc({text:'Short text',long:'Long text',number:'Number',date:'Date',select:'Dropdown'}[f.type])}</td><td class="sub">${esc((f.options||[]).join(', '))}</td><td><button class="btn sm danger" data-cfdel="${t}:${i}">Remove</button></td></tr>`));
    b=`<div class="grid g2"><div class="card"><div class="hd"><h3>Add a column</h3><small>Also adds the column to the SharePoint list</small></div><div class="bd"><form class="form" id="cfform"><label class="full">Page<select class="field" id="cf-t">${Object.entries(CF_TABLES).map(([k,l])=>`<option value="${k}">${esc(l)}</option>`).join('')}</select></label><label>Column name<input class="field" id="cf-l" required maxlength="40"></label><label>Type<select class="field" id="cf-ty"><option value="text">Short text</option><option value="long">Long text</option><option value="number">Number</option><option value="date">Date</option><option value="select">Dropdown</option></select></label><label class="full">Dropdown options<input class="field" id="cf-o" placeholder="Comma separated, only for Dropdown"></label><div class="full"><button class="btn pri" type="submit">Add column</button></div></form></div></div><div class="card"><div class="hd"><h3>Custom columns</h3></div><div class="bd" style="padding-top:8px"><div class="tw" style="border:0"><table><thead><tr><th class="nos">Page</th><th class="nos">Column</th><th class="nos">Type</th><th class="nos">Options</th><th class="nos"></th></tr></thead><tbody>${rowsF.join('')||'<tr><td colspan="5"><div class="empty">No custom columns yet.</div></td></tr>'}</tbody></table></div></div></div></div>`}
  else b=`<div class="card" style="max-width:560px"><div class="hd"><h3>Working hours</h3><small>Used for possible billable hours and utilization</small></div><div class="bd"><form class="form"><label>Hours per working day<input class="field" type="number" step="0.5" min="1" max="12" id="h-day" value="${c.hours.day}"></label><label>Hours off for half-day leave<input class="field" type="number" step="0.5" min="0" max="12" id="h-half" value="${c.hours.half}"></label><label>Hours off for permission<input class="field" type="number" step="0.5" min="0" max="12" id="h-perm" value="${c.hours.perm}"></label></form><button class="btn pri" data-setsave="hours" style="margin-top:12px">Save working hours</button></div></div>`;
  return head('Settings','Shape the app to how your team works: who can see which screen, page names, dropdown choices, extra columns and working hours.')+tabs+b;
};

/* ---------- first-time setup (SharePoint) ---------- */
function setupScreen(){const st=S.setup||{};const logo=ASSETS.logo?`<img class="logo-img" src="${ASSETS.logo}" alt="NovintiX">`:'';
  return`<div class="signin"><div class="box" style="width:min(720px,100%)">${logo}<h1>First-time <span>setup</span></h1><p>Signed in as <b>${esc(ME.email)}</b>. The Control Tower needs ${S.missing.length} SharePoint lists on <b>${esc(SITE)}</b>. This creates them, loads your starting data and gives the team access with suggested roles.</p>
  <div class="steps" style="color:var(--fg)">${[['Create the SharePoint lists',st.lists],['Load the starting data (26 people, 44 projects, 271 action items, leave and IDs)',st.data],['Give the team access with suggested roles',st.access]].map(([l,d],i)=>`<div class="step${d?' done':''}"><span class="n">${d?'✓':i+1}</span><span>${l}</span><span class="sub">${d?'Done':''}</span></div>`).join('')}</div>
  ${window.CT_SEED||st.data?'':`<label class="fine" for="seedfile">Starting data file (<b>starting-data.json</b> from the package — keep it off the web host)</label><input type="file" id="seedfile" accept=".json,application/json" style="color:#fff">`}
  ${st.msg?`<p>${esc(st.msg)}</p><div class="progress"><span style="width:${st.pct||0}%"></span></div>`:''}${st.err?`<div class="err">${esc(st.err)}</div>`:''}
  <button class="ms-btn" data-runsetup="1" ${st.running?'disabled':''}>${st.running?'Setting up…':'Run setup'}</button><button class="btn" data-signout="1" style="align-self:flex-start">Sign out</button></div></div>`}
async function runSetup(){const fileEl=document.getElementById('seedfile');const keepFile=fileEl&&fileEl.files[0];S.setup={running:true,msg:'Creating lists…',pct:2};paint();if(keepFile){try{window.CT_SEED=JSON.parse(await keepFile.text())}catch(er){S.setup={err:'That starting data file could not be read. Choose starting-data.json from the package.'};paint();return}}
  try{let i=0;for(const t of TABLES){S.setup.msg=`Creating ${SCHEMA[t].list}…`;S.setup.pct=Math.round(++i/TABLES.length*30);paint();await spCreateList(t)}S.setup.lists=true;
    await Promise.all(TABLES.map(t=>spLoadTable(t)));
    const seed=window.CT_SEED||{};const jobs=[];for(const t of Object.keys(seed))if(!Object.keys(S.t[t]||{}).length)for(const r of seed[t])jobs.push([t,r]);
    let done=0;const total=jobs.length;S.setup.msg=`Loading starting data: 0 of ${total}`;paint();
    Store.upsert=async(t,r)=>{const d=await sp(`${lt(t)}/items`,{method:'POST',body:rowToItem(t,r)});SPID[t][r.id]=d.Id;S.t[t][r.id]={...r}};
    const worker=async()=>{while(jobs.length){const[t,r]=jobs.shift();await Store.upsert(t,r);done++;if(done%10===0||done===total){S.setup.msg=`Loading starting data: ${done} of ${total}`;S.setup.pct=30+Math.round(done/Math.max(1,total)*60);paint()}}};
    await Promise.all([worker(),worker(),worker(),worker()]);S.setup.data=true;
    S.setup.msg='Giving the team access…';paint();
    for(const e of Object.values(S.t.employees)){const em=(e.email||'').trim().toLowerCase();if(!em||!domOK(em)||S.t.access[em])continue;await Store.upsert('access',{id:em,name:e.name,empId:e.id,role:em===ME.email?'Admin':defaultRole(e),active:'Yes',addedBy:ME.email})}
    if(!S.t.access[ME.email])await Store.upsert('access',{id:ME.email,name:ME.name,empId:'',role:'Admin',active:'Yes',addedBy:'setup'});
    S.setup.access=true;S.setup.msg='Done. Opening the Control Tower…';S.setup.pct=100;paint();
    S.missing=[];await spInit();resolveMe(true);STATE='app';setTimeout(route,800);
  }catch(e){S.setup.running=false;S.setup.err='Setup stopped: '+(e.msg||e.status||e)+'. You can run it again; finished steps are skipped.';paint()}}

/* =========================================================
   FORMS
   ========================================================= */
const SEL=lk=>({type:'select',lk});
const FORMS={
 emp:{screen:'team',table:'employees',title:'team member',fields:[['id','Employee ID',{req:1,lockEdit:1}],['name','Full name',{req:1}],['designation','Designation',{req:1}],['email','Email (@novintix.com)',{type:'email'}],['orgRole','Org role',SEL('orgRole')],['projectRole','Project role',SEL('projectRole')],['managerId','Reporting manager',{type:'emp'}],['doj','Date of joining',{type:'date'}],['yoePrior','YoE before NovintiX',{type:'number',step:'0.1'}],['dob','Birthday (DD-Mon)',{}],['billable','Billable',{opts:['Yes','No']}],['billableType','Billable type',SEL('billableType')],['empType','Employment type',SEL('empType')],['workMode','Work mode',SEL('workMode')],['shift','Shift',SEL('shift')],['seat','Seat no.',{}],['location','Location',{}],['country','Country',{}],['bu','Business unit',{}],['status','Employment status',SEL('empStatus')],['remarks','Remarks',{type:'textarea'}]],
  defaults:()=>({bu:'Product Engineering',location:'Coimbatore',country:'India',empType:'Full Time',workMode:'Office',shift:'General',status:'Active',orgRole:'Team Member',projectRole:'Project Executor',billable:'Yes'})},
 prj:{screen:'projects',table:'projects',title:'project / initiative',fields:[['id','ID',{req:1,lockEdit:1}],['portfolio','Portfolio',SEL('portfolio')],['program','Customer / program',{req:1,list:'programs'}],['project','Project',{req:1}],['goal','Goal',{type:'textarea',req:1}],['status','Status',SEL('prjStatus')],['start','Start date',{type:'date'}],['end','End date',{type:'date'}],['contract','Contract type',SEL('contract')],['revenue','Revenue type',SEL('revenue')],['sowStart','SOW start',{type:'date'}],['sowEnd','SOW end',{type:'date'}]],
  defaults:()=>({portfolio:'Product Engineering',status:'Active',contract:'T&M',revenue:'Billable',id:nextId('PE_DEL_')}),hint:()=>`Next free IDs: <b>${nextId('PE_DEL_')}</b> (customer), <b>${nextId('PE_SUP_')}</b> (support), <b>${nextId('NVTX_INTL_')}</b> (initiative).`},
 alloc:{screen:'allocation',table:'alloc',title:'allocation',fields:[['projectId','Project / initiative',{type:'prj',req:1,full:1}],['empId','Resource',{type:'emp',req:1}],['role','Role',SEL('allocRole')],['alloc','Allocation %',{type:'number',step:'5'}],['type','Allocation type',{opts:['Billable','Shadow']}],['remarks','Remarks',{type:'textarea'}]],defaults:()=>({role:'Support',alloc:0,type:'Shadow'})},
 act:{screen:'activities',table:'activities',title:'action item',fields:[['projectId','Project / initiative',{type:'prj',req:1,full:1}],['actId','Activity ID',{req:1}],['ownerId','Owner',{type:'emp',hint:'Leave empty for an item common to everyone'}],['desc','Action item',{type:'textarea',req:1}],['frequency','Frequency',SEL('actFrequency')],['priority','Priority',SEL('actPriority')],['status','Status',SEL('actStatus')],['start','Start date',{type:'date'}],['target','Target date',{type:'date'}],['actualEnd','Actual end date',{type:'date'}],['backup1','Backup 1',{type:'emp'}],['backup2','Backup 2',{type:'emp'}],['remarks','Remarks',{type:'textarea'}]],defaults:()=>({actId:nextAct(),status:'Open',priority:'Medium',frequency:'Once',start:TODAY()})},
 uid:{screen:'userids',table:'userids',title:'customer account',fields:[['empId','Team member',{type:'emp',req:1}],['projectId','Project',{type:'prj',req:1}],['program','Customer',{req:1,list:'programs'}],['project','Project name',{}],['sponsor','Project sponsor',{}],['userName','Account / user name',{req:1}],['expiry','Expiry date',{type:'date'}],['status','Status',SEL('uidStatus')]],defaults:()=>({status:'Active'})},
 leave:{screen:'leave',table:'leave',title:'leave / WFH',fields:[['empId','Team member',{type:'emp',req:1}],['kind','Type',{opts:['Leave','Half-day leave','WFH','Permission']}],['date','Date',{type:'date',req:1}],['to','Until (optional)',{type:'date',newOnly:1,hint:'Adds one record per working day'}],['clientInformed','Client informed',{opts:[['','Not applicable'],['Yes','Yes'],['No','No']]}],['remarks','Remarks',{type:'textarea'}]],defaults:()=>({kind:'Leave',date:TODAY(),clientInformed:'Yes'})},
 tsadmin:{screen:'productivity',table:'productivity',title:'productivity entry',fields:[['empId','Team member',{type:'emp',req:1}],['actRow','Activity',{type:'act',req:1,full:1}],['sub','Sub activity',{req:1,full:1}],['date','Effort date',{type:'date',req:1}],['hours','Hours',{type:'number',step:'0.25',req:1}],['remarks','Remarks',{full:1}]],defaults:()=>({date:TODAY(),hours:1})},
 acc:{screen:'access',table:'access',title:'app access',fields:[['id','Sign-in email',{type:'email',req:1,lockEdit:1,hint:'Their @novintix.com Microsoft 365 address'}],['empId','Employee record',{type:'emp'}],['name','Display name',{}],['role','Role',{type:'role',req:1}],['active','Access',{opts:[['Yes','Active'],['No','Removed']]}]],defaults:()=>({role:'Team Member',active:'Yes'})}
};
function nextId(pre){let mx=0;for(const id in P())if(id.startsWith(pre))mx=Math.max(mx,+id.slice(pre.length)||0);return pre+String(mx+1).padStart(4,'0')}
function nextAct(){let mx=0;for(const a of Object.values(AC())){const n=+(a.actId||'').replace(/\D/g,'');if(n>mx)mx=n}return'ACT '+String(mx+1).padStart(5,'0')}
let FCTX=null;
function openForm(kind,id){
  const F=FORMS[kind];const editing=!!id;let val;
  if(kind==='tsadmin'){if(id){const x=TS()[id];const a=Object.values(AC()).find(a=>a.actId===x.actId&&a.projectId===x.projectId);val={...x,actRow:a?a.id:''}}else val=F.defaults()}
  else val=id?{...rows(F.table)[id]}:F.defaults();
  if(kind==='acc'&&!id&&UI.prefillAcc){Object.assign(val,UI.prefillAcc);UI.prefillAcc=null}
  FCTX={kind,id,editing};
  const cfl=CF_TABLES[F.table]?cfFields(F.table).map(f=>['x_'+f.key,f.label,f.type==='select'?{opts:['',...(f.options||[])]}:{type:f.type==='long'?'textarea':f.type==='text'?'text':f.type}]):[];
  const emps=empOpts(),prjs=prjOpts(),progs=opt(Object.values(P()).map(p=>p.program));
  const fld=([k,l,o])=>{if(o.newOnly&&editing)return'';const v=val[k]??'';const req=o.req?' required':'';const ro=o.lockEdit&&editing?' readonly':'';const idA=`id="ff-${k}" name="${k}"`;let inp;
    if(o.lk||o.opts){let os=o.lk?[...L(o.lk)]:o.opts;if(o.lk&&v&&!os.includes(String(v)))os=[String(v),...os];inp=`<select class="field" ${idA}${req}>${os.map(x=>{const[a,b]=Array.isArray(x)?x:[x,x];return`<option value="${esc(a)}"${String(v)===a?' selected':''}>${esc(b)}</option>`}).join('')}</select>`}
    else if(o.type==='emp')inp=`<select class="field" ${idA}${req}><option value="">—</option>${emps.map(([a,b])=>`<option value="${esc(a)}"${v===a?' selected':''}>${esc(b)}</option>`).join('')}</select>`;
    else if(o.type==='prj')inp=`<select class="field" ${idA}${req}><option value="">Choose…</option>${prjs.map(([a,b])=>`<option value="${esc(a)}"${v===a?' selected':''}>${esc(b)}</option>`).join('')}</select>`;
    else if(o.type==='role')inp=`<select class="field" ${idA}${req}>${Object.keys(CFG().roles).map(r=>`<option${v===r?' selected':''}>${esc(r)}</option>`).join('')}</select>`;
    else if(o.type==='act')inp=`<select class="field" ${idA}${req}><option value="">Choose…</option>${Object.values(AC()).sort((a,b)=>(a.actId||'').localeCompare(b.actId||'')).map(x=>`<option value="${esc(x.id)}"${v===x.id?' selected':''}>${esc(x.actId)} · ${esc((x.desc||'').slice(0,70))} · ${esc(x.ownerId?empName(x.ownerId):'Common')}</option>`).join('')}</select>`;
    else if(o.type==='textarea')inp=`<textarea class="field" ${idA}${req}>${esc(v)}</textarea>`;
    else inp=`<input class="field" type="${o.type||'text'}" ${o.step?`step="${o.step}"`:''} ${o.list?`list="dl-${o.list}"`:''} ${idA} value="${esc(v)}"${req}${ro}>`;
    return`<label class="${o.type==='textarea'||o.full?'full':''}">${esc(l)}${o.req?' *':''}${inp}${o.hint?`<span class="hint">${esc(o.hint)}</span>`:''}</label>`};
  const accExtra=kind==='emp'&&!editing&&canEdit('access')?`<label class="full" style="flex-direction:row;align-items:center;gap:8px;color:var(--fg)"><input type="checkbox" id="ff-giveaccess" checked> Also give app access with a suggested role</label>`:'';
  $('#layer').innerHTML=`<div class="scrim" data-close="1"><div class="drawer" role="dialog" aria-modal="true" aria-label="${editing?'Edit':'Add'} ${esc(F.title)}" style="--mod:var(${SC[F.screen][3]})"><div><div class="eyebrow">${editing?'Edit':'New'} ${esc(F.title)}</div><h3>${esc(editing?(val.name||val.goal||val.desc||val.userName||val.id||''):'Add '+F.title)}</h3></div>${F.hint?`<div class="note">${F.hint()}</div>`:''}
   <form class="form" id="drawerform">${[...F.fields,...cfl].map(fld).join('')}${accExtra}<datalist id="dl-programs">${progs.map(p=>`<option value="${esc(p)}">`).join('')}</datalist></form><div id="confirm"></div>
   <div class="foot"><div>${editing&&kind!=='acc'?`<button class="btn danger" data-del="1">Delete</button>`:''}</div><div style="display:flex;gap:8px"><button class="btn" data-close="1">Cancel</button><button class="btn pri" data-save="1">${editing?'Save changes':'Add'}</button></div></div></div></div>`;
  const f1=$('#drawerform .field:not([readonly])');f1&&f1.focus();
  if(kind==='uid')$('#ff-projectId').addEventListener('change',e=>{const p=prj(e.target.value);if(p){$('#ff-program').value=p.program;$('#ff-project').value=p.project}});
  if(kind==='acc')$('#ff-empId').addEventListener('change',e=>{const x=E()[e.target.value];if(x){if(!$('#ff-name').value)$('#ff-name').value=x.name;if(!$('#ff-id').value&&x.email)$('#ff-id').value=x.email.trim().toLowerCase();$('#ff-role').value=defaultRole(x)}});
}
function closeForm(){$('#layer').innerHTML='';FCTX=null}
function formVals(){const o={};for(const el of $('#drawerform').elements){if(!el.name)continue;o[el.name]=el.type==='number'?(el.value===''?0:+el.value):el.value.trim()}return o}
async function saveForm(){
  const form=$('#drawerform');if(!form.reportValidity())return;const{kind,id,editing}=FCTX;const F=FORMS[kind];const v=formVals();
  const btn=document.querySelector('[data-save]');btn.disabled=true;btn.textContent='Saving…';
  try{
    if(kind==='emp'){if(!editing&&E()[v.id])throw{msg:'That Employee ID already exists.'};v.managerName=v.managerId?empName(v.managerId):'';const give=$('#ff-giveaccess')?.checked;await Store.upsert('employees',v);
      const em=(v.email||'').toLowerCase();if(give&&em){if(!domOK(em))toast('Saved. App access needs an @novintix.com email.',true);else if(!ACC()[em])await grantAccess({id:em,name:v.name,empId:v.id,role:defaultRole(v),active:'Yes'})}}
    else if(kind==='prj'){if(!editing&&P()[v.id])throw{msg:'That ID already exists.'};await Store.upsert('projects',v)}
    else if(kind==='leave'){const {byE}=allocBy();const clients=opt((byE[v.empId]||[]).filter(a=>isClient(a.projectId)&&a.type==='Billable').map(a=>prj(a.projectId)?.program)).join('; ');
      const xs=Object.fromEntries(Object.entries(v).filter(([k])=>k.startsWith('x_')));
      const mk=(d,rid)=>({id:rid,empId:v.empId,date:d,kind:v.kind,clientInformed:['WFH','Permission'].includes(v.kind)?'':v.clientInformed,clients,remarks:v.remarks,...xs});
      if(editing)await Store.upsert('leave',mk(v.date,id));
      else{let ds=[v.date];if(v.to&&v.to>v.date){ds=[];for(let d=v.date;d<=v.to;d=addD(d,1))if(!isWE(d))ds.push(d)}if(ds.length>31)throw{msg:'Pick a range of 31 working days or fewer.'};for(const d of ds)await Store.upsert('leave',mk(d,nid('l')))}}
    else if(kind==='tsadmin'){const a=AC()[v.actRow];if(!a)throw{msg:'Choose an activity.'};await Store.upsert('productivity',{id:id||nid('x'),empId:v.empId,actId:a.actId,actDesc:a.desc,projectId:a.projectId,sub:v.sub,date:v.date,hours:+v.hours,remarks:v.remarks,by:ME.email})}
    else if(kind==='acc'){v.id=v.id.toLowerCase();if(!domOK(v.id))throw{msg:'Use their @novintix.com email address.'};if(!editing&&ACC()[v.id])throw{msg:'This person already has access. Edit their row instead.'};if(v.id===ME.email&&v.role!=='Admin'&&ME.role==='Admin'&&!Object.values(ACC()).some(a=>a.id!==ME.email&&a.role==='Admin'&&a.active!=='No'))throw{msg:"You're the only Admin. Make someone else Admin first."};
      if(!v.name&&v.empId)v.name=empName(v.empId);await grantAccess({...rows('access')[v.id]||{},...v},editing)}
    else{const r={...v,id:id||v.id||nid(kind[0])};await Store.upsert(F.table,r)}
    closeForm();toast(editing?'Saved':'Added');render();
  }catch(e){btn.disabled=false;btn.textContent=editing?'Save changes':'Add';toast(errMsg(e),true)}
}
async function grantAccess(a,editing){a.addedBy=a.addedBy||ME.email;await Store.upsert('access',a);
  if(!DEMO&&a.active!=='No'&&!editing){const r=CFG().roles[a.role]||{};const edits=a.role==='Admin'||Object.values(r).some(x=>x==='edit'||x==='own');
    try{await spAddToSiteGroup(a.id,edits)}catch(e){toast(`Access saved. Also add ${a.id} to the SharePoint site's ${edits?'Members':'Visitors'} group (Site settings → Site permissions); the app couldn't do it automatically.`,true)}}}
async function grantAll(){const todo=Object.values(E()).filter(e=>e.status!=='Exited'&&e.email&&!ACC()[(e.email||'').trim().toLowerCase()]&&domOK(e.email.trim().toLowerCase()));
  for(const e of todo)await grantAccess({id:e.email.trim().toLowerCase(),name:e.name,empId:e.id,role:defaultRole(e),active:'Yes'});toast(`Gave access to ${todo.length} people`);render()}
function askDelete(){$('#confirm').innerHTML=`<div class="confirm">Delete this record for everyone? This can't be undone. <button class="btn sm danger" data-delyes="1">Delete</button><button class="btn sm" data-delno="1">Keep</button></div>`}
async function doDelete(){const{kind,id}=FCTX;try{await Store.remove(FORMS[kind].table,id);closeForm();toast('Deleted');render()}catch(e){toast(errMsg(e),true)}}

/* ---------- settings save ---------- */
async function saveCfg(key,obj){await Store.upsert('config',{id:key,value:JSON.stringify(obj)});toast('Settings saved');render()}
function settingsSave(which){const c=CFG();
  if(which==='roles'){const roles=JSON.parse(JSON.stringify(c.roles));document.querySelectorAll('[data-rm]').forEach(s=>{const[r,k]=s.dataset.rm.split('|');(roles[r]??={})[k]=s.value});saveCfg('roles',roles).catch(e=>toast(errMsg(e),true))}
  else if(which==='pages'){const pages={};for(const[k]of SCREENS){if(['settings','access'].includes(k))continue;const sh=document.getElementById('pgs-'+k);pages[k]={title:(document.getElementById('pg-'+k)?.value||'').trim(),show:sh?sh.checked:true}}saveCfg('pages',pages).catch(e=>toast(errMsg(e),true))}
  else if(which==='lists'){const lists={};for(const k in LIST_DEF){const v=[...new Set((document.getElementById('ls-'+k).value||'').split('\n').map(x=>x.trim()).filter(Boolean))];if(!v.length){toast(LIST_DEF[k][0]+' needs at least one option.',true);return}if(LOCKED[k]&&!v.includes(LOCKED[k])){toast(`Keep "${LOCKED[k]}" in ${LIST_DEF[k][0]}.`,true);return}lists[k]=v}saveCfg('lists',lists).catch(e=>toast(errMsg(e),true))}
  else if(which==='hours')saveCfg('hours',{day:+$('#h-day').value||8,half:+$('#h-half').value||0,perm:+$('#h-perm').value||0}).catch(e=>toast(errMsg(e),true));
}
async function addCustomField(){const t=$('#cf-t').value,label=$('#cf-l').value.trim(),type=$('#cf-ty').value;if(!label){toast('Give the column a name.',true);return}
  const key=label.toLowerCase().replace(/[^a-z0-9]+/g,'').slice(0,24)||'field';const custom=JSON.parse(JSON.stringify(CFG().custom));const arr=custom[t]||(custom[t]=[]);
  if(arr.some(f=>f.key===key)){toast('That page already has a column with this name.',true);return}
  const options=type==='select'?$('#cf-o').value.split(',').map(x=>x.trim()).filter(Boolean):[];if(type==='select'&&!options.length){toast('Add the dropdown options, separated by commas.',true);return}
  try{await Store.ensureField(t,'x_'+key,label,type==='number'?N:type==='long'?M:T);arr.push({key,label,type,options});await saveCfg('custom',custom)}catch(e){toast(errMsg(e),true)}}

/* ---------- export ---------- */
function csv(a){if(!a.length)return'';const h=Object.keys(a[0]);const q=v=>{v=String(v??'');return/[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};return[h.map(q).join(','),...a.map(r=>h.map(k=>q(r[k])).join(','))].join('\r\n')}
const us=s=>{if(!s)return'';const d=parse(s);return`${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`};
async function exportData(w){let out=[],fn='';
  if(w==='emp'){fn='Employee_Database';out=Object.values(E()).map(e=>({'Employee ID':e.id,'Employee Name':e.name,'Designation':e.designation,'Org. Role':e.orgRole,'Project Role':e.projectRole,'Employee Mail ID':e.email,'Reporting Manager':e.managerId?empName(e.managerId):e.managerName,'Date Of Joining':us(e.doj),'YoE before NovintiX':e.yoePrior,'Total YoE':n1(totalYoE(e)),'DOB (DD - MMM)':e.dob,'Business Unit':e.bu,'Billable':e.billable,'Billable Type':e.billableType,'Location':e.location,'Country':e.country,'Employment Type':e.empType,'Work Mode':e.workMode,'Shift':e.shift,'Employment Status':e.status,'Seat No.':e.seat,'Remarks':e.remarks,...cfExport('employees',e)}))}
  else if(w==='prj'){fn='Project_Initiatives_Dump';out=Object.values(P()).map(p=>({'Project / Initiative ID':p.id,'Portfolio':p.portfolio,'Program':p.program,'Project':p.project,'Project / Initiative Goal':p.goal,'Project / Initiative Status':p.status,'Start Date':us(p.start),'End Date':us(p.end),'Contract Type':p.contract,'Revenue Type':p.revenue,'SOW Start Date':us(p.sowStart),'SOW End Date':us(p.sowEnd),...cfExport('projects',p)}))}
  else if(w==='alloc'){fn='Billability_Utilization';out=Object.values(AL()).map(a=>({'Project / Initiative ID':a.projectId,'Project/Initiative Goal':prj(a.projectId)?.goal||'','Employee ID':a.empId,'Employee Name':empName(a.empId),'Role':a.role,'Allocation':(+a.alloc||0)+'%','Allocation Type':a.type,'Remarks':a.remarks,...cfExport('alloc',a)}))}
  else if(w==='act'){fn='Activities_Responsibilities';out=Object.values(AC()).map(a=>({'Project ID':a.projectId,'Project / Initiative Goal':prj(a.projectId)?.goal||'','Activity ID':a.actId,'Activity Owner':a.ownerId?empName(a.ownerId):'','Activity Owner: Employee ID':a.ownerId,'Activity Description':a.desc,'Frequency':a.frequency,'Priority':a.priority,'Status':a.status,'Start Date':us(a.start),'Target Date':us(a.target),'Actual End Date':us(a.actualEnd),'Backup #1':a.backup1?empName(a.backup1):'','Backup #2':a.backup2?empName(a.backup2):'','Remarks':a.remarks,...cfExport('activities',a)}))}
  else if(w==='uid'){fn='USER_ID_Details_Customer';out=Object.values(UID()).map(u=>({'Employee ID':u.empId,'Employee Name':empName(u.empId),'Project ID':u.projectId,'Program Name':u.program,'Project Name':u.project,'Project Sponsor':u.sponsor,'User Name':u.userName,'User ID Expiry':us(u.expiry),'User ID Status':u.status,...cfExport('userids',u)}))}
  else if(w==='leave'){fn='Leave_Details';out=Object.values(LV()).sort((a,b)=>a.date.localeCompare(b.date)).map(l=>({'Employee Name':empName(l.empId),'Employee ID':l.empId,'Date':us(l.date),'Leave Type':l.kind,'Client Informed':l.clientInformed,'Clients':l.clients,'Remarks':l.remarks||'',...cfExport('leave',l)}))}
  else if(w==='ts'){fn='Productivity_Tracker';out=Object.values(TS()).sort((a,b)=>a.date.localeCompare(b.date)).map(x=>{const p=prj(x.projectId)||{};return{'Employee Name':empName(x.empId),'Employee ID':x.empId,'Activity ID':x.actId,'Activity Description':x.actDesc,'Project ID':x.projectId,'Category':isClient(x.projectId)?'Client':'Initiative','Program':p.program||'','Project':p.project||'','Sub Activity':x.sub,'Effort Date':us(x.date),'Effort Hours':x.hours,'Remarks':x.remarks||''}})}
  if(!out.length){toast('Nothing to export yet.',true);return}
  const data='﻿'+csv(out),name=`${fn}_${TODAY()}.csv`;
  const dl=window.claude?.use?await window.claude.use('downloads'):null;
  if(dl){dl.save({filename:name,data}).catch(e=>{if(e&&e.code!=='declined')toast('Export is not available here.',true)});return}
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([data],{type:'text/csv'}));a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},500)}

/* =========================================================
   EVENTS
   ========================================================= */
let qT=0;
document.addEventListener('input',e=>{const el=e.target;if(el.dataset.f){(UI.f[el.dataset.f]??={})[el.dataset.fk]=el.value;if(el.tagName==='INPUT'){clearTimeout(qT);qT=setTimeout(render,160)}else render()}});
document.addEventListener('click',async e=>{
  const t=e.target.closest('[data-signin],[data-demosignin],[data-signout],[data-refresh],[data-runsetup],[data-per],[data-grantall],[data-addrole],[data-delrole],[data-setsave],[data-cfdel],[data-go],[data-add],[data-export],[data-sort],[data-row],[data-tab],[data-close],[data-save],[data-del],[data-delyes],[data-delno],[data-wk],[data-day],[data-cal],[data-delts]');if(!t)return;const d=t.dataset;
  if(d.close!==undefined&&e.target!==t&&!e.target.closest('button'))return;
  if(d.signin)signIn();
  else if(d.demosignin){const a=ACC()[$('#demo-who').value];await afterSignIn(a.id,a.name)}
  else if(d.signout)signOut();
  else if(d.refresh){t.disabled=true;try{await Store.loadAll();toast('Up to date')}catch(er){toast(errMsg(er),true)}render()}
  else if(d.runsetup)runSetup();
  else if(d.per){UI.per=d.per;render()}
  else if(d.grantall)grantAll().catch(er=>toast(errMsg(er),true));
  else if(d.addrole){const n=($('#newrole').value||'').trim();if(!n)return;const roles=JSON.parse(JSON.stringify(CFG().roles));if(roles[n]){toast('That role exists.',true);return}roles[n]={...ROLE_DEF['Viewer']};saveCfg('roles',roles).catch(er=>toast(errMsg(er),true))}
  else if(d.delrole){const roles=JSON.parse(JSON.stringify(CFG().roles));delete roles[d.delrole];saveCfg('roles',roles).catch(er=>toast(errMsg(er),true))}
  else if(d.setsave)settingsSave(d.setsave);
  else if(d.cfdel){const[tb,i]=d.cfdel.split(':');const c=JSON.parse(JSON.stringify(CFG().custom));(c[tb]||[]).splice(+i,1);saveCfg('custom',c).catch(er=>toast(errMsg(er),true))}
  else if(d.go)location.hash=d.go;
  else if(d.add)openForm(d.add);
  else if(d.export)exportData(d.export);
  else if(d.sort){const s=UI.sort[d.sort]||{};UI.sort[d.sort]={k:d.k,d:s.k===d.k&&s.d!=='desc'?'desc':'asc'};render()}
  else if(d.row)openForm(d.row,d.id);
  else if(d.tab){UI.tab[d.tab]=d.v;render()}
  else if(d.close)closeForm();
  else if(d.save)saveForm();
  else if(d.del)askDelete();
  else if(d.delyes)doDelete();
  else if(d.delno)$('#confirm').innerHTML='';
  else if(d.wk!==undefined){const t0=TODAY();UI.week=d.wk==='0'?monday(t0):addD(UI.week||monday(t0),+d.wk);if((UI.tab.prod||'mine')==='mine')UI.day=d.wk==='0'?t0:UI.week;render()}
  else if(d.day){UI.day=d.day;render()}
  else if(d.cal){const[y,m]=(UI.cal||TODAY().slice(0,7)).split('-').map(Number);const n=new Date(y,m-1+(+d.cal),1);UI.cal=`${n.getFullYear()}-${pad(n.getMonth()+1)}`;render()}
  else if(d.delts){const x=TS()[d.delts];if(!x||(x.empId!==ME.empId&&!canEdit('productivity')))return;try{await Store.remove('productivity',x.id);toast('Entry deleted');render()}catch(er){toast(errMsg(er),true)}}
});
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&FCTX)closeForm();if(e.key==='Enter'&&e.target.matches('tr[data-row]'))openForm(e.target.dataset.row,e.target.dataset.id)});
document.addEventListener('submit',async e=>{e.preventDefault();
  if(e.target.id==='logform'){const a=AC()[$('#lg-act').value];const date=$('#lg-date').value;const h=+$('#lg-h').value;const me=ME.empId;
    if(!a||!date||!(h>0)){toast('Choose an activity, a date and the hours.',true);return}
    const tot=Object.values(TS()).filter(x=>x.empId===me&&x.date===date).reduce((s,x)=>s+(+x.hours||0),0);if(tot+h>16){toast('That would take this day above 16 hours.',true);return}
    const btn=e.target.querySelector('button[type=submit]');btn.disabled=true;
    try{await Store.upsert('productivity',{id:nid('x'),empId:me,actId:a.actId,actDesc:a.desc,projectId:a.projectId,sub:$('#lg-sub').value.trim(),date,hours:h,remarks:$('#lg-rem').value.trim(),by:ME.email});UI.day=date;toast('Logged '+n1(h)+' h');render()}catch(er){toast(errMsg(er),true)}finally{btn.disabled=false}}
  else if(e.target.id==='drawerform')saveForm();
  else if(e.target.id==='cfform')addCustomField();
});

boot();
})();
