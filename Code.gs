/* ARK Automated Payroll Management System - Code.gs (Google Apps Script backend)
   Run setup() ONCE. DB = Google Sheets (ID kept in Script Properties, never sent to the browser). */
const S={
EMPLOYEES:['EmpID','LastName','FirstName','Dept','Position','PayType','Rate','TIN','SSSNo','PhilHealthNo','PagIBIGNo','ScheduleID','Status','DateHired'],
BIOMETRIC_RAW:['RawID','BatchID','EmpID','Date','Time','Source','ImportedBy','ImportedAt','Key'],
ATTENDANCE:['AttID','PeriodID','EmpID','Date','TimeIn','TimeOut','WorkedHrs','RegularHrs','OTHrs','LateMin','UnderMin','Status','Key'],
WORK_SCHEDULE:['ScheduleID','Name','TimeIn','TimeOut','BreakMin','GraceMin','WorkDays','OTMinThreshold'],
PAYROLL_PERIODS:['PeriodID','Label','StartDate','EndDate','ApplyGov','Status','UpdatedAt'],
PAYROLL_REGISTER:['RegID','PeriodID','EmpID','Name','TIN','PayType','Basic','OT','Allowances','AbsDed','LateDed','UnderDed','OtherDed','SSS_EE','SSS_ER','PH_EE','PH_ER','PI_EE','PI_ER','Tax','Gross','NetPay','Status','Key'],
SSS_TABLE:['Effective','MinComp','MaxComp','EE','ER','EC'],
PHILHEALTH_TABLE:['Effective','Min','Max','Rate','Fixed'],
PAGIBIG_TABLE:['Effective','Min','Max','EERate','ERRate','MaxBase'],
BIR_TAX_TABLE:['Effective','Over','NotOver','BaseTax','Rate'],
DEDUCTIONS:['DedID','EmpID','PeriodID','Description','Amount'],
ALLOWANCES:['AlwID','EmpID','PeriodID','Description','Amount'],
PAYSLIPS:['PayslipNo','PeriodID','EmpID','RegID','GeneratedAt','Data'],
USERS:['UserID','Username','Salt','Hash','Role','Active','FullName'],
AUDIT_LOGS:['LogID','Time','User','Action','Details'],
SETTINGS:['Key','Value']};
// text-formatted columns (1-based) so IDs/dates/times are never auto-converted
const TXT={EMPLOYEES:[1,12,14],BIOMETRIC_RAW:[3,4,5,9],ATTENDANCE:[3,4,5,6,13],WORK_SCHEDULE:[3,4],PAYROLL_PERIODS:[3,4],SSS_TABLE:[1],PHILHEALTH_TABLE:[1],PAGIBIG_TABLE:[1],BIR_TAX_TABLE:[1]};
const DAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'], PP=PropertiesService.getScriptProperties();
const DEF={OT_MULT:'1.25',HOURS_PER_DAY:'8',DIVISOR:'261',PERIODS_PER_MONTH:'2',DAILY_FACTOR:'26',HOLIDAYS:'2026-12-25,2026-12-30',COMPANY:'ARK Technological Institute Education System Inc.',BIO_MAP:'{}'};

/* ---------- SETUP ---------- */
function setup(){
  let id=PP.getProperty('DB_ID'),ss;
  if(id)ss=SpreadsheetApp.openById(id);else{ss=SpreadsheetApp.create('ARK Payroll Database');PP.setProperty('DB_ID',ss.getId());}
  Object.keys(S).forEach(n=>{const sh=ss.getSheetByName(n)||ss.insertSheet(n);
    if(sh.getLastRow()===0){sh.appendRow(S[n]);sh.setFrozenRows(1);sh.getRange(1,1,1,S[n].length).setBackground('#C8102E').setFontColor('#fff').setFontWeight('bold');}
    (TXT[n]||[]).forEach(c=>sh.getRange(2,c,sh.getMaxRows()-1,1).setNumberFormat('@'));});
  const d=ss.getSheetByName('Sheet1');if(d)ss.deleteSheet(d);
  if(!rows_('USERS').length){const pw=Utilities.getUuid().slice(0,10);saveUser_({Username:'admin',FullName:'System Administrator',Role:'ADMIN',Active:'Y',password:pw});Logger.log('ADMIN LOGIN -> username: admin  password: '+pw+'  (change it after first login)');}
  if(!rows_('WORK_SCHEDULE').length)put_('WORK_SCHEDULE',[['SCH-1','Regular 8-5','08:00','17:00',60,10,'Mon,Tue,Wed,Thu,Fri',30]]);
  if(!rows_('SETTINGS').length)put_('SETTINGS',Object.keys(DEF).map(k=>[k,DEF[k]]));
  const E='2025-01-01';
  if(!rows_('SSS_TABLE').length){const t=[[E,0,5249.99,250,500,10]];for(let m=5500;m<=34500;m+=500)t.push([E,m-250,m+249.99,m*.05,m*.10,m>=15000?30:10]);t.push([E,34750,1e9,1750,3500,30]);put_('SSS_TABLE',t);}
  if(!rows_('PHILHEALTH_TABLE').length)put_('PHILHEALTH_TABLE',[[E,0,9999.99,0,500],[E,10000,99999.99,0.05,''],[E,100000,1e9,0,5000]]);
  if(!rows_('PAGIBIG_TABLE').length)put_('PAGIBIG_TABLE',[[E,0,1500,0.01,0.02,10000],[E,1500.01,1e9,0.02,0.02,10000]]);
  if(!rows_('BIR_TAX_TABLE').length)put_('BIR_TAX_TABLE',[[E,0,250000,0,0],[E,250000,400000,0,0.15],[E,400000,800000,22500,0.20],[E,800000,2000000,102500,0.25],[E,2000000,8000000,402500,0.30],[E,8000000,1e12,2202500,0.35]]);
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(2).create();
  Logger.log('Database ready: '+ss.getUrl());
}

/* ---------- DB HELPERS ---------- */
const db_=()=>SpreadsheetApp.openById(PP.getProperty('DB_ID'));
const TZ=()=>Session.getScriptTimeZone(), id_=s=>String(s).trim().replace(/^0+(?=\d)/,''), r2=x=>Math.round((+x+Number.EPSILON)*100)/100;
const cell_=v=>v instanceof Date?Utilities.formatDate(v,TZ(),v.getFullYear()<1950?'HH:mm':'yyyy-MM-dd'):v;
function rows_(n){const v=db_().getSheetByName(n).getDataRange().getValues(),h=v.shift();return v.filter(r=>r[0]!=='').map(r=>{const o={};h.forEach((k,i)=>o[k]=cell_(r[i]));return o;});}
function put_(n,a){if(a.length)db_().getSheetByName(n).getRange(db_().getSheetByName(n).getLastRow()+1,1,a.length,a[0].length).setValues(a);}
function add_(n,o){put_(n,[S[n].map(k=>o[k]??'')]);}
function save_(n,objs){const sh=db_().getSheetByName(n),w=S[n].length;if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,w).clearContent();if(objs.length)sh.getRange(2,1,objs.length,w).setValues(objs.map(o=>S[n].map(k=>o[k]??'')));}
function ups_(n,key,o){const a=rows_(n),i=a.findIndex(x=>String(x[key])===String(o[key]));if(i<0)a.push(o);else a[i]=Object.assign(a[i],o);save_(n,a);}
function cfg_(){const c=Object.assign({},DEF);rows_('SETTINGS').forEach(r=>c[r.Key]=String(r.Value));return c;}
function audit_(u,a,d){add_('AUDIT_LOGS',{LogID:Utilities.getUuid(),Time:Utilities.formatDate(new Date(),TZ(),'yyyy-MM-dd HH:mm:ss'),User:u,Action:a,Details:d});}
const m_=t=>{const p=String(t).split(':');return +p[0]*60+ +p[1];};
const nd_=d=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()+1);return Utilities.formatDate(x,TZ(),'yyyy-MM-dd');};
const per_=i=>{const p=rows_('PAYROLL_PERIODS').find(x=>x.PeriodID===i);if(!p)throw new Error('Select a valid payroll period');return p;};
const setPer_=(p,st)=>ups_('PAYROLL_PERIODS','PeriodID',{PeriodID:p.PeriodID,Status:st,UpdatedAt:new Date().toISOString()});
function eff_(n,date){const r=rows_(n),m=r.map(x=>x.Effective).filter(d=>d<=date).sort().pop();return r.filter(x=>x.Effective===m);}

/* ---------- AUTH ---------- */
const hash_=(s,p)=>Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s+p+PP.getProperty('DB_ID')));
function login_(p){const c=CacheService.getScriptCache(),fk='f_'+p.username;
  if((+c.get(fk)||0)>=5)throw new Error('Account locked for 15 minutes after repeated failures');
  const u=rows_('USERS').find(x=>x.Username===p.username&&x.Active==='Y');
  if(!u||hash_(u.Salt,p.password||'')!==u.Hash){c.put(fk,String((+c.get(fk)||0)+1),900);audit_(p.username||'?','LOGIN_FAILED','');throw new Error('Invalid username or password');}
  c.remove(fk);const t=Utilities.getUuid()+Utilities.getUuid();
  c.put('s_'+t,JSON.stringify({u:u.Username,r:u.Role}),21600);audit_(u.Username,'LOGIN','');
  return {token:t,user:u.Username,role:u.Role,name:u.FullName};}
function session_(t){const v=CacheService.getScriptCache().get('s_'+t);if(!v)throw new Error('Session expired. Please sign in again.');return JSON.parse(v);}
function saveUser_(o){const a=rows_('USERS'),ex=a.find(x=>x.Username===o.Username);let r={UserID:ex?ex.UserID:'U-'+(a.length+1),Username:o.Username,FullName:o.FullName,Role:o.Role,Active:o.Active||'Y'};
  if(!/^(ADMIN|PAYROLL|HR)$/.test(r.Role))throw new Error('Invalid role');
  if(o.password){if(o.password.length<8)throw new Error('Password must be at least 8 characters');r.Salt=Utilities.getUuid();r.Hash=hash_(r.Salt,o.password);}else if(!ex)throw new Error('Password required for new users');
  ups_('USERS','Username',r);}

/* ---------- WEB APP ENTRY ---------- */
function doGet(){return ContentService.createTextOutput(JSON.stringify({ok:true,service:'ARK Payroll API'})).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){let r;try{const b=JSON.parse(e.postData.contents);r=api(b.token,b.action,b.payload);}catch(x){r={ok:false,error:x.message};}
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);}
function api(token,action,p){
  try{p=p||{};if(action==='login')return {ok:true,data:login_(p)};
    const s=session_(token),a=AC[action];if(!a||a[0].indexOf(s.r[0]==='A'?'A':s.r[0])<0)throw new Error('Not authorized');
    if(!a[1])return {ok:true,data:a[2](p,s)};
    const L=LockService.getScriptLock();L.waitLock(20000);try{return {ok:true,data:a[2](p,s)};}finally{L.releaseLock();}
  }catch(e){return {ok:false,error:e.message};}}
// roles: A=ADMIN P=PAYROLL H=HR ; [roles, isWrite, handler]
const GOV=['SSS_TABLE','PHILHEALTH_TABLE','PAGIBIG_TABLE','BIR_TAX_TABLE'],TV=GOV.concat(['ALLOWANCES','DEDUCTIONS','AUDIT_LOGS']);
const AC={
dashboard:['AHP',0,()=>{const e=rows_('EMPLOYEES'),p=rows_('PAYROLL_PERIODS'),r=rows_('PAYROLL_REGISTER').filter(x=>x.Status==='APPROVED');return{employees:e.filter(x=>x.Status==='Active').length,periods:p.length,pending:p.filter(x=>x.Status!=='APPROVED').length,approvedNet:r2(r.reduce((a,x)=>a+ +x.NetPay,0)),recent:p.slice(-5)};}],
listEmployees:['AHP',0,()=>rows_('EMPLOYEES')],
saveEmployee:['AH',1,(p,s)=>{if(!p.EmpID||!p.LastName||!p.FirstName)throw new Error('Employee ID and name are required');if(!/^(Monthly|Daily|Hourly)$/.test(p.PayType)||!(+p.Rate>=0))throw new Error('Valid pay type and rate required');ups_('EMPLOYEES','EmpID',p);audit_(s.u,'SAVE_EMPLOYEE',p.EmpID);return 1;}],
listSchedules:['AHP',0,()=>rows_('WORK_SCHEDULE')],
saveSchedule:['AH',1,(p,s)=>{if(!p.ScheduleID)throw new Error('Schedule ID required');ups_('WORK_SCHEDULE','ScheduleID',p);audit_(s.u,'SAVE_SCHEDULE',p.ScheduleID);return 1;}],
listPeriods:['AHP',0,()=>rows_('PAYROLL_PERIODS')],
savePeriod:['AP',1,(p,s)=>{if(!p.StartDate||!p.EndDate||p.EndDate<p.StartDate)throw new Error('Invalid period dates');p.PeriodID=p.PeriodID||'PP-'+p.StartDate+'_'+p.EndDate;const ex=rows_('PAYROLL_PERIODS').find(x=>x.PeriodID===p.PeriodID);if(ex&&ex.Status==='APPROVED')throw new Error('Approved periods are locked');p.Status=ex?ex.Status:'OPEN';p.UpdatedAt=new Date().toISOString();ups_('PAYROLL_PERIODS','PeriodID',p);audit_(s.u,'SAVE_PERIOD',p.PeriodID);return 1;}],
previewBio:['AHP',0,p=>vrows_(p.rows,per_(p.periodId))],
importBio:['AHP',1,(p,s)=>{const per=per_(p.periodId);if(per.Status==='APPROVED')throw new Error('Period already approved');const v=vrows_(p.rows,per).filter(x=>x.status==='OK'),b='B-'+Date.now(),now=new Date().toISOString();
  put_('BIOMETRIC_RAW',v.map((x,i)=>[b+'-'+i,b,x.empId,x.date,x.time,p.source||'UPLOAD',s.u,now,x.key]));audit_(s.u,'IMPORT_BIOMETRIC',v.length+' rows into '+per.PeriodID);return v.length;}],
processAttendance:['AHP',1,(p,s)=>{const per=per_(p.periodId);if(per.Status==='APPROVED')throw new Error('Period already approved');const n=procAtt_(per);setPer_(per,'OPEN');audit_(s.u,'PROCESS_ATTENDANCE',per.PeriodID);return n;}],
listAttendance:['AHP',0,p=>rows_('ATTENDANCE').filter(a=>a.PeriodID===p.periodId)],
verifyAttendance:['AP',1,(p,s)=>{const per=per_(p.periodId),a=rows_('ATTENDANCE').filter(x=>x.PeriodID===per.PeriodID);if(!a.length)throw new Error('Upload and process biometric data first');
  const inc=a.filter(x=>x.Status==='INCOMPLETE').length;if(inc)throw new Error(inc+' incomplete record(s): add the missing punch on the Attendance page');setPer_(per,'ATTENDANCE_VERIFIED');audit_(s.u,'VERIFY_ATTENDANCE',per.PeriodID);return 1;}],
calculatePayroll:['AP',1,(p,s)=>calc_(per_(p.periodId),s)],
listRegister:['AP',0,p=>rows_('PAYROLL_REGISTER').filter(x=>x.PeriodID===p.periodId)],
approvePayroll:['AP',1,(p,s)=>approve_(per_(p.periodId),s)],
listPayslips:['AP',0,p=>rows_('PAYSLIPS').filter(x=>x.PeriodID===p.periodId).map(x=>({PayslipNo:x.PayslipNo,EmpID:x.EmpID,Data:JSON.parse(x.Data)}))],
listTable:['AP',0,(p,s)=>{if(TV.indexOf(p.name)<0||(p.name==='AUDIT_LOGS'&&s.r!=='ADMIN'))throw new Error('Not authorized');let r=rows_(p.name);return{cols:S[p.name],rows:p.name==='AUDIT_LOGS'?r.slice(-300).reverse():r};}],
addRow:['AP',1,(p,s)=>{if(['ALLOWANCES','DEDUCTIONS'].concat(GOV).indexOf(p.name)<0||(GOV.indexOf(p.name)>=0&&s.r!=='ADMIN'))throw new Error('Not authorized');const o=p.row;if(p.name==='ALLOWANCES')o.AlwID='A-'+Date.now();if(p.name==='DEDUCTIONS')o.DedID='D-'+Date.now();if(S[p.name].some(k=>!/ID$|Fixed|PeriodID/.test(k)&&o[k]===undefined))throw new Error('Complete all fields');add_(p.name,o);audit_(s.u,'ADD_'+p.name,JSON.stringify(o));return 1;}],
listUsers:['A',0,()=>rows_('USERS').map(u=>({Username:u.Username,FullName:u.FullName,Role:u.Role,Active:u.Active}))],
saveUser:['A',1,(p,s)=>{saveUser_(p);audit_(s.u,'SAVE_USER',p.Username);return 1;}],
getSettings:['AHP',0,()=>cfg_()],
saveSettings:['A',1,(p,s)=>{Object.keys(p).forEach(k=>{if(k in DEF)ups_('SETTINGS','Key',{Key:k,Value:p[k]});});audit_(s.u,'SAVE_SETTINGS',Object.keys(p).join(','));return 1;}],
backup:['A',1,(p,s)=>{const n=backupNow();audit_(s.u,'BACKUP',n);return n;}]};

/* ---------- BIOMETRIC VALIDATION ---------- */
function vrows_(rows,per){const emp={},keys={};rows_('EMPLOYEES').forEach(e=>emp[id_(e.EmpID)]=1);rows_('BIOMETRIC_RAW').forEach(r=>keys[r.Key]=1);
  return rows.map(r=>{const k=id_(r.empId)+'|'+r.date+'|'+r.time;let st='OK',m='';
    if(!/^\d{4}-\d{2}-\d{2}$/.test(r.date)||!/^\d{2}:\d{2}$/.test(r.time)){st='ERROR';m='Invalid date/time';}
    else if(!emp[id_(r.empId)]){st='ERROR';m='Unknown employee ID';}
    else if(r.date<per.StartDate||r.date>per.EndDate){st='SKIPPED';m='Outside payroll period';}
    else if(keys[k]){st='DUPLICATE';m='Already imported or repeated in file';}else keys[k]=1;
    return{empId:r.empId,date:r.date,time:r.time,status:st,msg:m,key:k};});}

/* ---------- ATTENDANCE ENGINE ---------- */
function procAtt_(per){
  const cf=cfg_(),hol=cf.HOLIDAYS.split(',').map(x=>x.trim()),sc={},g={},out=[];
  rows_('WORK_SCHEDULE').forEach(x=>sc[x.ScheduleID]=x);const def=Object.values(sc)[0];
  rows_('BIOMETRIC_RAW').forEach(r=>{if(r.Date>=per.StartDate&&r.Date<=per.EndDate)(g[id_(r.EmpID)+'|'+r.Date]=g[id_(r.EmpID)+'|'+r.Date]||[]).push(r.Time);});
  rows_('EMPLOYEES').filter(e=>e.Status==='Active').forEach(e=>{
    const s=sc[e.ScheduleID]||def,wd=String(s.WorkDays).split(','),st=m_(s.TimeIn),en=m_(s.TimeOut),br=+s.BreakMin,gr=+s.GraceMin,th=+s.OTMinThreshold,sm=en-st-br;
    for(let d=per.StartDate;d<=per.EndDate;d=nd_(d)){
      const t=(g[id_(e.EmpID)+'|'+d]||[]).sort(),rest=wd.indexOf(DAYS[new Date(d+'T00:00:00').getDay()])<0,h=hol.indexOf(d)>=0;
      const o={AttID:'AT-'+per.PeriodID+'-'+e.EmpID+'-'+d,PeriodID:per.PeriodID,EmpID:e.EmpID,Date:d,TimeIn:t[0]||'',TimeOut:'',WorkedHrs:0,RegularHrs:0,OTHrs:0,LateMin:0,UnderMin:0,Key:e.EmpID+'|'+d};
      if(!t.length)o.Status=rest?'REST':h?'HOLIDAY':'ABSENT';
      else if(t.length<2)o.Status='INCOMPLETE';
      else{const i=m_(t[0]),u=m_(t[t.length-1]),span=u-i,w=Math.max(0,span-(span>300?br:0));o.TimeOut=t[t.length-1];o.WorkedHrs=r2(w/60);
        if(rest){o.Status='REST_WORKED';o.OTHrs=r2(w/60);}
        else{o.Status='PRESENT';o.RegularHrs=r2(Math.min(w,sm)/60);o.LateMin=i-st>gr?i-st:0;o.UnderMin=u<en?en-u:0;o.OTHrs=u-en>=th?r2((u-en)/60):0;}}
      out.push(o);}});
  save_('ATTENDANCE',rows_('ATTENDANCE').filter(a=>a.PeriodID!==per.PeriodID).concat(out));return out.length;}

/* ---------- PAYROLL ENGINE ---------- */
function tax_(bir,x){const r=bir.filter(b=>x>+b.Over).sort((a,b)=>a.Over-b.Over).pop();return r?+r.BaseTax+(x-+r.Over)*+r.Rate:0;}
function calc_(per,s){
  if(per.Status==='APPROVED')throw new Error('Approved payroll is locked');if(per.Status==='OPEN')throw new Error('Verify attendance first');
  const c=cfg_(),ppm=+c.PERIODS_PER_MONTH,hpd=+c.HOURS_PER_DAY,dv=+c.DIVISOR,att=rows_('ATTENDANCE').filter(a=>a.PeriodID===per.PeriodID),al=rows_('ALLOWANCES'),dd=rows_('DEDUCTIONS');
  const sss=eff_('SSS_TABLE',per.EndDate),ph=eff_('PHILHEALTH_TABLE',per.EndDate),pi=eff_('PAGIBIG_TABLE',per.EndDate),bir=eff_('BIR_TAX_TABLE',per.EndDate),reg=[];
  if(per.ApplyGov==='Y'&&!(sss.length&&ph.length&&pi.length&&bir.length))throw new Error('Contribution tables missing for this period date');
  rows_('EMPLOYEES').filter(e=>e.Status==='Active').forEach(e=>{
    const a=att.filter(x=>id_(x.EmpID)===id_(e.EmpID)),cn=t=>a.filter(x=>x.Status===t).length,sm=k=>a.reduce((n,x)=>n+ +x[k],0),rate=+e.Rate,ty=e.PayType;
    const hr=ty==='Monthly'?rate*12/dv/hpd:ty==='Daily'?rate/hpd:rate,day=hr*hpd;let basic,abs=0;
    if(ty==='Monthly'){basic=rate/ppm;abs=cn('ABSENT')*day;}else if(ty==='Daily')basic=day*(cn('PRESENT')+cn('HOLIDAY'));else basic=hr*sm('RegularHrs');
    const late=ty==='Hourly'?0:sm('LateMin')/60*hr,und=ty==='Hourly'?0:sm('UnderMin')/60*hr,ot=sm('OTHrs')*hr*(+c.OT_MULT);
    const mine=l=>l.filter(x=>id_(x.EmpID)===id_(e.EmpID)&&(!x.PeriodID||x.PeriodID===per.PeriodID)).reduce((n,x)=>n+ +x.Amount,0),alw=mine(al),od=mine(dd);
    const g={sE:0,sR:0,pE:0,pR:0,iE:0,iR:0,tax:0},base=ty==='Monthly'?rate:ty==='Daily'?day*+c.DAILY_FACTOR:basic*ppm;
    if(per.ApplyGov==='Y'&&base>0){ // contributions & tax computed on monthly-equivalent pay, deducted once per month (ApplyGov period)
      const r1=sss.find(r=>base>=+r.MinComp&&base<=+r.MaxComp)||sss[sss.length-1];g.sE=+r1.EE;g.sR=+r1.ER+ +r1.EC;
      const r2_=ph.find(r=>base>=+r.Min&&base<=+r.Max)||ph[ph.length-1],tot=r2_.Fixed!==''?+r2_.Fixed:base*+r2_.Rate;g.pE=g.pR=tot/2;
      const r3=pi.find(r=>base>=+r.Min&&base<=+r.Max)||pi[pi.length-1],pb=Math.min(base,+r3.MaxBase);g.iE=pb*+r3.EERate;g.iR=pb*+r3.ERRate;
      g.tax=tax_(bir,Math.max(0,base-g.sE-g.pE-g.iE)*12)/12;} // TIN on employee record is the tax-reporting ID, not a deduction
    const gross=basic+ot+alw,net=gross-abs-late-und-od-g.sE-g.pE-g.iE-g.tax,k=per.PeriodID+'|'+e.EmpID;
    reg.push({RegID:'RG-'+per.PeriodID+'-'+e.EmpID,PeriodID:per.PeriodID,EmpID:e.EmpID,Name:e.LastName+', '+e.FirstName,TIN:e.TIN,PayType:ty,Basic:r2(basic),OT:r2(ot),Allowances:r2(alw),AbsDed:r2(abs),LateDed:r2(late),UnderDed:r2(und),OtherDed:r2(od),SSS_EE:r2(g.sE),SSS_ER:r2(g.sR),PH_EE:r2(g.pE),PH_ER:r2(g.pR),PI_EE:r2(g.iE),PI_ER:r2(g.iR),Tax:r2(g.tax),Gross:r2(gross),NetPay:r2(net),Status:'DRAFT',Key:k});});
  save_('PAYROLL_REGISTER',rows_('PAYROLL_REGISTER').filter(x=>x.PeriodID!==per.PeriodID).concat(reg));setPer_(per,'CALCULATED');audit_(s.u,'CALCULATE_PAYROLL',per.PeriodID+' ('+reg.length+' employees)');return reg.length;}
function approve_(per,s){
  if(per.Status!=='CALCULATED')throw new Error('Calculate and review payroll before approving');
  if(rows_('PAYSLIPS').some(x=>x.PeriodID===per.PeriodID))throw new Error('Duplicate payroll transaction: payslips already exist for this period');
  const all=rows_('PAYROLL_REGISTER'),mine=all.filter(x=>x.PeriodID===per.PeriodID),co=cfg_().COMPANY,now=new Date().toISOString();
  mine.forEach(x=>x.Status='APPROVED');save_('PAYROLL_REGISTER',all);
  put_('PAYSLIPS',mine.map(x=>['PS-'+per.PeriodID+'-'+x.EmpID,per.PeriodID,x.EmpID,x.RegID,now,JSON.stringify(Object.assign({Company:co,Period:per.Label+' ('+per.StartDate+' to '+per.EndDate+')'},x))]));
  setPer_(per,'APPROVED');audit_(s.u,'APPROVE_PAYROLL',per.PeriodID);return mine.length;}

/* ---------- BACKUP / RECOVERY ---------- */
function backupNow(){const f=DriveApp.getFileById(PP.getProperty('DB_ID')),it=DriveApp.getFoldersByName('ARK_Payroll_Backups'),fo=it.hasNext()?it.next():DriveApp.createFolder('ARK_Payroll_Backups');
  const n='ARK Payroll Backup '+Utilities.formatDate(new Date(),TZ(),'yyyy-MM-dd HHmm');f.makeCopy(n,fo);
  const fs=[],i=fo.getFiles();while(i.hasNext())fs.push(i.next());fs.sort((a,b)=>b.getDateCreated()-a.getDateCreated()).slice(30).forEach(x=>x.setTrashed(true));return n;}
