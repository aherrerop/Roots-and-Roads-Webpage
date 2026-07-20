/* ===== Regression tests for the 2026-07-20 bugs ===== */
let pass=0, fail=0;
const check=(label,cond,got)=>{ if(cond){pass++;console.log('PASS  '+label);}
  else{fail++;console.log('FAIL  '+label+'  (got: '+JSON.stringify(got)+')');} };

const control=new __mock.MockSS('control'); SpreadsheetApp._active=control;

/* --- BUG 1: tours must exist even when the Week tab lacks their column --- */
console.log('--- buildShifts_: rules win over availability columns ---');
const mon=new Date(2026,6,20,12);  // Monday Jul 20 2026
const rules=[
 {day:'Monday',time:'11:00',language:'English',guidesNeeded:1,activeFrom:null,activeUntil:null},
 {day:'Monday',time:'17:00',language:'English',guidesNeeded:1,activeFrom:null,activeUntil:null}
];
// Week tab only exposed 10:00/17:00 columns -> calendar slots only have those
const calendar=[
 {week:'Week 30',dateObj:mon,dateTimeObj:new Date(2026,6,20,10),dateText:'2026-07-20',day:'Monday',time:'10:00'},
 {week:'Week 30',dateObj:mon,dateTimeObj:new Date(2026,6,20,17),dateText:'2026-07-20',day:'Monday',time:'17:00'}
];
const avail=[{week:'Week 30',guideName:'Albert',dateObj:mon,dateTimeObj:new Date(2026,6,20,17),dateText:'2026-07-20',day:'Monday',time:'17:00'}];
const shifts=buildShifts_(avail,rules,calendar);
check('11:00 English EXISTS despite missing availability column',
  shifts.some(s=>s.time==='11:00'&&s.language==='English'), shifts.map(s=>s.time));
check('11:00 has empty availability (Not assigned later)',
  (shifts.find(s=>s.time==='11:00')||{}).availableGuides.length===0, null);
check('17:00 keeps Albert availability',
  (shifts.find(s=>s.time==='17:00')||{}).availableGuides.join()==='Albert', null);

/* --- BUG 2: Weekly_Schedule Date-coerced times (12/30/1899) still parse --- */
console.log('--- readWeeklySchedule_: coerced Date times ---');
const ws=control.insertSheet('Weekly_Schedule');
ws.getRange(1,1,4,6).setValues([
 ['Day','Time','Language','Guides needed','Active from','Active until'],
 ['Monday','11:00','English',1,'',''],
 ['Monday',new Date(1899,11,30,10,0),'German',1,'',''],      // corrupted cell
 ['Wednesday',new Date(1899,11,30,17,0),'German',1,'','']]);
const parsed=readWeeklySchedule_(control);
check('string time rule parsed', parsed.some(r=>r.time==='11:00'&&r.language==='English'), parsed);
check('Date-coerced 10:00 German rule RECOVERED', parsed.some(r=>r.time==='10:00'&&r.language==='German'), parsed.map(r=>r.time+'/'+r.language));
check('Date-coerced 17:00 German rule RECOVERED', parsed.some(r=>r.time==='17:00'&&r.language==='German'), null);

/* --- offer writer: repairs corrupted German rows, adds no Private rows --- */
console.log('--- updateWeeklyScheduleToCurrentOffer ---');
updateWeeklyScheduleToCurrentOffer();
const after=ws.getRange(1,1,ws.getLastRow(),6).getDisplayValues();
check('no Private language rows written', !after.some(r=>/private/i.test(r[2])), after.map(r=>r[2]));
check('German preserved with REPAIRED time (10:00, not 12/30/1899)',
  after.some(r=>r[2]==='German'&&r[1]==='10:00'), after.filter(r=>r[2]==='German').map(r=>r[1]));
check('English 11:00 Mon-Tue-Thu-Fri present',
  after.filter(r=>r[2]==='English'&&r[1]==='11:00').length===4, null);
check('Spanish 10:30 present', after.filter(r=>r[2]==='Spanish'&&r[1]==='10:30').length===4, null);
const reparsed=readWeeklySchedule_(control);
check('round-trip: all repaired rules parse', reparsed.length===after.length-1, {rules:reparsed.length, rows:after.length-1});

/* --- private availability pseudo-rules --- */
const priv=privateAvailabilityRules_();
check('private slots exposed for availability (Mon 10:30+17:00)',
  priv.filter(r=>r.day==='Monday').map(r=>r.time).sort().join()==='10:30,17:00', priv.filter(r=>r.day==='Monday'));
check('Saturday only 17:00 private slot', priv.filter(r=>r.day==='Saturday').map(r=>r.time).join()==='17:00', null);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
