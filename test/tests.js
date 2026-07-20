/* ============ END-TO-END TESTS against the real deployed functions ============ */
let pass=0, fail=0;
const check=(label,cond,got)=>{ if(cond){pass++;console.log('PASS  '+label);}
  else{fail++;console.log('FAIL  '+label+'  (got: '+JSON.stringify(got)+')');} };
const day=(offset)=>{ const d=new Date(); d.setDate(d.getDate()+offset); d.setHours(12,0,0,0); return d; };
const key=(d)=>Utilities.formatDate(d,null,'yyyy-MM-dd');

/* ---------- setup mock spreadsheets ---------- */
const control=new __mock.MockSS('control'); SpreadsheetApp._active=control;
__mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw']=new __mock.MockSS('booking');
const ledger=new __mock.MockSS('ledger'); __mock.SS_BY_ID['LEDGER']=ledger; __mock.PROPS['LEDGER_ID']='LEDGER';

// Guides tab (needed by portal + validate)
const guides=control.insertSheet('Guides');
guides.getRange(1,1,4,10).setValues([
 ['Guide','Active?','Seniority','English','German','Spanish','French','Manager','Email','Password'],
 ['Albert',true,1,true,false,false,false,true,'a@x.com','pw'],
 ['Carlos',true,1,true,false,true,false,true,'c@x.com','pw'],
 ['Mario',true,2,true,true,false,false,'','m@x.com','pw']]);

/* ============ TEST GROUP 1: LEDGER (the duplicate check-in bug) ============ */
console.log('\n--- Ledger dedupe & Date-coercion ---');
const D=key(day(1));
const mkRow=(bid,ck)=>makeLedgerRow_({dateKey:D,day:'Tomorrow',timeLabel:'11:00 AM',language:'English',
  bookingName:'Elisa',phone:'+39',source:'GetYourGuide',guests:2,children:0,checkedIn:ck,
  weOwe:20,theyOwe:0,rrMakes:7,type:'Paid',bookingId:bid});

writeGuideLedger_('Albert', D, '11:00', 'English', [mkRow('GYG1',2), mkRow('GYG2',1)]);
writeGuideLedger_('Albert', D, '11:00', 'English', [mkRow('GYG1',2), mkRow('GYG2',1)]);
writeGuideLedger_('Albert', D, '11:00', 'English', [mkRow('GYG1',2), mkRow('GYG2',1)]);
let tab=ledger.getSheetByName('Albert');
check('3 saves of same shift -> exactly 2 rows (no duplicates)', tab.getLastRow()-1===2, tab.getLastRow()-1);

// Simulate OLD bad rows: Time stored as a real Date (the coercion trap)
tab.getRange(6,1,1,16).setValues([[D,'Tomorrow',new Date(1899,11,30,11,0),'English','Ghost','+1','GetYourGuide',2,0,2,20,0,7,'Paid','GYG1','2026-07-19 09:00']]);
writeGuideLedger_('Albert', D, '11:00', 'English', [mkRow('GYG1',2), mkRow('GYG2',1)]);
const ids=tab.getRange(2,1,tab.getLastRow()-1,16).getValues().map(r=>r[14]).filter(String);
check('re-save removes even Date-coerced old row for same booking', ids.filter(x=>x==='GYG1').length===1, ids);

const cks=readGuideCheckins_('Albert');
check('readGuideCheckins_ finds both check-ins', Object.keys(cks).length===2, cks);
check('checked-in count read back', cks[Object.keys(cks).find(k=>k.endsWith('GYG1'))]===2, cks);

/* ============ TEST GROUP 2: GRID round trip (writer -> portal reader) ============ */
console.log('\n--- Schedule grid round trip ---');
const d1=day(1), d2=day(2);
const shifts=[
 {dateText:key(d1),time:'11:00',language:'English',isPrivate:false,privIndex:0,assignedGuides:['Albert'],lockedGuides:[],status:'OK',hasLockConflict:false,dateTimeObj:d1},
 {dateText:key(d1),time:'17:00',language:'English',isPrivate:false,privIndex:0,assignedGuides:[],lockedGuides:[],status:'Not assigned',hasLockConflict:false,dateTimeObj:d1},
 {dateText:key(d1),time:'10:00',language:'English',isPrivate:true,privIndex:1,assignedGuides:['Carlos'],lockedGuides:['Carlos'],status:'OK',hasLockConflict:false,dateTimeObj:d1},
 {dateText:key(d1),time:'10:00',language:'English',isPrivate:true,privIndex:2,assignedGuides:[],lockedGuides:[],status:'Not assigned',hasLockConflict:false,dateTimeObj:d1},
 {dateText:key(d2),time:'11:00',language:'English',isPrivate:false,privIndex:0,assignedGuides:['Mario'],lockedGuides:[],status:'OK',hasLockConflict:false,dateTimeObj:d2}
];
makeOneLanguageScheduleTab_(control,'English',shifts);
const gsh=control.getSheetByName('Schedule_English');
const hdr=gsh.getRange(2,1,1,gsh.getLastColumn()).getDisplayValues()[0];
check('private groups get their own columns', hdr.join('|').includes('10:00 · Private') && hdr.join('|').includes('10:00 · Private 2'), hdr);

const sched=readSchedule_();
check('portal reads 5 shifts back', sched.length===5, sched.length);
const priv1=sched.find(s=>s.private&&s.privIndex===1);
check('portal sees private #1 with Carlos', priv1&&priv1.assigned.join()==='Carlos', priv1&&priv1.assigned);
const priv2=sched.find(s=>s.private&&s.privIndex===2);
check('portal sees private #2 unassigned', priv2&&priv2.assigned.length===0&&priv2.status==='Not assigned', priv2);
check('regular 11:00 has Albert', (sched.find(s=>!s.private&&s.time==='11:00'&&s.dateKey===key(d1))||{}).assigned.join()==='Albert', null);

const locks=readLockedAssignments_(control);
const lockKeys=Object.keys(locks);
check('bold lock survives write->read (Carlos on private #1)',
  lockKeys.length===1 && locks[lockKeys[0]].join()==='Carlos' && lockKeys[0].endsWith('|P1'), locks);

/* ---- portal assign API writes a bold lock the scheduler will keep ---- */
const res=writeAssignmentToGrid_('English', key(d1), '17:00', false, 1, 'Mario');
check('assign API ok', res.ok===true, res);
const locks2=readLockedAssignments_(control);
check('assign API produced a readable lock (Mario 17:00)',
  Object.keys(locks2).some(k=>k.includes('17:00')&&locks2[k].join()==='Mario'), locks2);
const sched2=readSchedule_();
check('portal now shows Mario on 17:00', (sched2.find(s=>!s.private&&s.time==='17:00')||{}).assigned.join()==='Mario', null);

/* ============ TEST GROUP 3: the 2h cutoff ============ */
console.log('\n--- 2h tour cutoff ---');
const now=new Date();
const past=new Date(now.getTime()-3*3600000);   // started 3h ago today
const soon=new Date(now.getTime()+3600000);     // starts in 1h today
const t24=x=>x.getHours()+':'+String(x.getMinutes()).padStart(2,'0');
const shifts3=[
 {dateText:key(now),time:t24(past),language:'English',isPrivate:false,privIndex:0,assignedGuides:['Albert'],lockedGuides:[],status:'OK',hasLockConflict:false,dateTimeObj:past},
 {dateText:key(now),time:t24(soon),language:'English',isPrivate:false,privIndex:0,assignedGuides:['Albert'],lockedGuides:[],status:'OK',hasLockConflict:false,dateTimeObj:soon}
];
makeOneLanguageScheduleTab_(control,'English',shifts3);
const sched3=readSchedule_().filter(s=>s.dateKey===key(now));
check('tour that started 3h ago is gone from the portal', !sched3.some(s=>s.time===t24(past)), sched3.map(s=>s.time));
check('tour starting in 1h is still shown', sched3.some(s=>s.time===t24(soon)), sched3.map(s=>s.time));

/* ============ TEST GROUP 4: stale grid purge ============ */
console.log('\n--- Stale grid purge ---');
const gsh2=control.insertSheet('Schedule_German');
gsh2.getRange(1,1,3,2).setValues([['German schedule (2026-07-01 to 2026-07-10)',''],['Date','10:00'],['Thu Jul 2','Albert']]);
makeLanguageScheduleTabs_(control, shifts3.map(s=>Object.assign({},s)));  // only English shifts
const after=control.getSheetByName('Schedule_German');
check('German tab with no shifts gets CLEARED (stale Albert gone)', after.getLastRow()<=1, after.getRange(1,1,Math.max(1,after.getLastRow()),2).getDisplayValues());
check('portal no longer reads any German shift', !readSchedule_().some(s=>s.language==='German'), null);

console.log('\n=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
