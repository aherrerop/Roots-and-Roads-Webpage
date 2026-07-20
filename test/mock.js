/* ---- Minimal Google Apps Script mock with the Sheets value-coercion trap ---- */
const DAYS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const MONS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtDate(d, tz, pat){
  const p2=n=>String(n).padStart(2,'0');
  switch(pat){
    case 'yyyy-MM-dd': return d.getFullYear()+'-'+p2(d.getMonth()+1)+'-'+p2(d.getDate());
    case 'yyyy-MM-dd HH:mm': return fmtDate(d,tz,'yyyy-MM-dd')+' '+p2(d.getHours())+':'+p2(d.getMinutes());
    case 'yyyy-MM-dd HH:mm:ss': return fmtDate(d,tz,'yyyy-MM-dd HH:mm')+':'+p2(d.getSeconds());
    case 'EEE MMM d': return DAYS[d.getDay()].slice(0,3)+' '+MONS[d.getMonth()]+' '+d.getDate();
    case 'EEE, MMM d': return DAYS[d.getDay()].slice(0,3)+', '+MONS[d.getMonth()]+' '+d.getDate();
    case 'EEEE': return DAYS[d.getDay()];
    case 'H:mm': return d.getHours()+':'+p2(d.getMinutes());
    case 'EEE, d MMM yyyy': return DAYS[d.getDay()].slice(0,3)+', '+d.getDate()+' '+MONS[d.getMonth()]+' '+d.getFullYear();
    case 'yyyy_MM': return d.getFullYear()+'_'+p2(d.getMonth()+1);
    default: return d.toISOString();
  }
}
global.Utilities={
  formatDate: fmtDate,
  base64EncodeWebSafe:x=>Buffer.from(String(x)).toString('base64url'),
  base64DecodeWebSafe:x=>Buffer.from(String(x),'base64url'),
  computeDigest:(a,s)=>Buffer.from(String(s)), DigestAlgorithm:{MD5:'md5'},
  newBlob:b=>({getDataAsString:()=>Buffer.from(b).toString()})
};
global.Session={getScriptTimeZone:()=>'Europe/Madrid'};
global.Logger={log:m=>console.log('[Logger] '+m)};
global.console.error=console.log;
global.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})};
global.MailApp={sendEmail:()=>{},getRemainingDailyQuota:()=>99};
global.ContentService={createTextOutput:t=>({setMimeType:()=>({})}),MimeType:{TEXT:1,JSON:2,JAVASCRIPT:3}};
global.UrlFetchApp={fetch:()=>({getResponseCode:()=>200,getContentText:()=>'{}'})};
global.ScriptApp={getProjectTriggers:()=>[]};
global.GmailApp={getUserLabelByName:()=>null,search:()=>[],createLabel:()=>{}};
global.DriveApp={getFolderById:()=>({addFile:()=>{}}),getRootFolder:()=>({removeFile:()=>{}}),getFileById:()=>({makeCopy:()=>{}})};

const PROPS={};
global.PropertiesService={getScriptProperties:()=>({
  getProperty:k=>PROPS[k]||null, setProperty:(k,v)=>{PROPS[k]=v;}, deleteProperty:k=>{delete PROPS[k];}})};

/* value coercion: time-with-AM/PM strings become Dates unless the cell is '@' */
function coerce(v,fmt){
  if(typeof v==='string' && fmt!=='@'){
    const m=v.trim().match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/i);
    if(m){ let h=+m[1]; if(/pm/i.test(m[3])&&h!==12)h+=12; if(/am/i.test(m[3])&&h===12)h=0;
      return new Date(1899,11,30,h,+m[2]); }
  }
  return v;
}
function disp(v){
  if(v instanceof Date){
    if(v.getFullYear()===1899){ let h=v.getHours(); const s=h>=12?'PM':'AM'; h=((h+11)%12)+1;
      return h+':'+String(v.getMinutes()).padStart(2,'0')+' '+s; }
    return fmtDate(v,null,'yyyy-MM-dd');
  }
  return v==null?'':String(v);
}
class MockSheet{
  constructor(name){this.name=name;this.rows=[];this.fmts=[];this.rich=[];this.frozen=0;this.hidden=false;}
  _ensure(r,c){ while(this.rows.length<r){this.rows.push([]);this.fmts.push([]);this.rich.push([]);}
    for(let i=0;i<this.rows.length;i++){ while(this.rows[i].length<c){this.rows[i].push('');this.fmts[i].push('');this.rich[i].push(null);} } }
  getName(){return this.name;}
  getMaxRows(){return Math.max(this.rows.length,100);}
  getMaxColumns(){return Math.max(...this.rows.map(r=>r.length),26);}
  getLastRow(){ for(let i=this.rows.length-1;i>=0;i--){ if(this.rows[i].some(c=>c!==''&&c!=null)) return i+1; } return 0; }
  getLastColumn(){ let m=0; this.rows.forEach(r=>{ for(let j=r.length-1;j>=0;j--){ if(r[j]!==''&&r[j]!=null){ m=Math.max(m,j+1); break; } } }); return m; }
  setFrozenRows(n){this.frozen=n;return this;}
  hideSheet(){this.hidden=true;}
  clear(){this.rows=[];this.fmts=[];this.rich=[];return this;}
  clearFormats(){return this;}
  deleteRow(r){this.rows.splice(r-1,1);this.fmts.splice(r-1,1);this.rich.splice(r-1,1);}
  insertRowBefore(r){this._ensure(r,1);this.rows.splice(r-1,0,[]);this.fmts.splice(r-1,0,[]);this.rich.splice(r-1,0,[]);}
  appendRow(vals){const r=this.getLastRow()+1;this.getRange(r,1,1,vals.length).setValues([vals]);}
  setColumnWidth(){return this;} setColumnWidths(){return this;} setRowHeight(){return this;} setRowHeights(){return this;}
  autoResizeColumns(){return this;}
  getDataRange(){const lr=Math.max(1,this.getLastRow()),lc=Math.max(1,this.getLastColumn());return this.getRange(1,1,lr,lc);}
  getRange(a,b,c,d){
    if(typeof a==='string'){ // 'B:B' style — format-only ops in our code paths
      return new MockRange(this,1,1,this.getMaxRows(),1,true);
    }
    return new MockRange(this,a,b,c||1,d||1,false);
  }
}
class MockRange{
  constructor(sh,r,c,nr,nc,colMode){this.sh=sh;this.r=r;this.c=c;this.nr=nr;this.nc=nc;this.colMode=colMode;}
  setValues(vals){this.sh._ensure(this.r+this.nr-1,this.c+this.nc-1);
    for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++){
      const fmt=this.sh.fmts[this.r-1+i][this.c-1+j];
      this.sh.rows[this.r-1+i][this.c-1+j]=coerce(vals[i][j],fmt);
      this.sh.rich[this.r-1+i][this.c-1+j]=null;}
    return this;}
  setValue(v){return this.setValues([[v]]);}
  getValues(){this.sh._ensure(this.r+this.nr-1,this.c+this.nc-1);
    const out=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++)row.push(this.sh.rows[this.r-1+i][this.c-1+j]);out.push(row);}return out;}
  getValue(){return this.getValues()[0][0];}
  getDisplayValues(){return this.getValues().map(r=>r.map(disp));}
  setNumberFormat(f){this.sh._ensure(this.r+this.nr-1,this.c+this.nc-1);
    for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++)this.sh.fmts[this.r-1+i][this.c-1+j]=f;return this;}
  setRichTextValue(rt){this.sh._ensure(this.r,this.c);
    this.sh.rows[this.r-1][this.c-1]=rt.text;this.sh.rich[this.r-1][this.c-1]=rt;return this;}
  getRichTextValues(){this.sh._ensure(this.r+this.nr-1,this.c+this.nc-1);
    const out=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++){
      const rt=this.sh.rich[this.r-1+i][this.c-1+j];const v=this.sh.rows[this.r-1+i][this.c-1+j];
      row.push(rt?rtObj(rt):rtObj({text:disp(v),runs:[]}));}out.push(row);}return out;}
  insertCheckboxes(){return this;}
  merge(){return this;} breakApart(){return this;}
  setFontWeight(){return this;} setFontSize(){return this;} setFontStyle(){return this;} setFontColor(){return this;}
  setFontFamily(){return this;} setBackground(){return this;} setBorder(){return this;} setWrap(){return this;}
  setHorizontalAlignment(){return this;} setVerticalAlignment(){return this;}
  clearContent(){for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++){if(this.sh.rows[this.r-1+i])this.sh.rows[this.r-1+i][this.c-1+j]='';}return this;}
  getSheet(){return this.sh;}
  getRow(){return this.r;} getColumn(){return this.c;}
  getNumRows(){return this.nr;} getNumColumns(){return this.nc;}
  setNote(n){(this.sh.notes=this.sh.notes||{})[this.r+'|'+this.c]=n;return this;}
  getNote(){return (this.sh.notes||{})[this.r+'|'+this.c]||'';}
  clearDataValidations(){return this;}
}
function rtObj(rt){
  const text=rt.text||'';
  const runs=(rt.runs&&rt.runs.length)?rt.runs:[{start:0,end:text.length,bold:false}];
  return { getText:()=>text,
    getRuns:()=>runs.map(r=>({ getText:()=>text.slice(r.start,r.end),
      getTextStyle:()=>({isBold:()=>!!r.bold}) })) };
}
class MockSS{
  constructor(id){this.id=id;this.sheets=[];}
  getId(){return this.id;}
  getUrl(){return 'mock://'+this.id;}
  getSheetByName(n){return this.sheets.find(s=>s.name===n)||null;}
  insertSheet(n){const s=new MockSheet(n);this.sheets.push(s);return s;}
  getSheets(){return this.sheets.slice();}
  deleteSheet(s){this.sheets=this.sheets.filter(x=>x!==s);}
  toast(){}
}
const SS_BY_ID={};
global.__mock={MockSS,SS_BY_ID,PROPS};
global.SpreadsheetApp={
  _active:null,
  getActiveSpreadsheet(){return this._active;},
  openById(id){ if(!SS_BY_ID[id]) SS_BY_ID[id]=new MockSS(id); return SS_BY_ID[id]; },
  create(name){const ss=new MockSS('created-'+name);SS_BY_ID[ss.getId()]=ss;return ss;},
  newRichTextValue(){ const o={text:'',runs:[]};
    return { setText(t){o.text=t;return this;},
      setTextStyle(s,e,st){o.runs.push({start:s,end:e,bold:st.__bold});return this;},
      build(){return o;} }; },
  newTextStyle(){ const s={__bold:false};
    return { setBold(b){s.__bold=b;return this;}, build(){return s;} }; },
  newDataValidation(){ return { requireValueInList(){return this;}, setAllowInvalid(){return this;}, build(){return {};} }; }
};
