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
ws.getRange(1,1,6,9).setValues([
 ['Day','Time','Language','Guides needed','Active from','Active until','Guide','Hide from availability','Private'],
 ['Monday','11:00','English',1,'','','','',''],
 ['Monday',new Date(1899,11,30,10,0),'German',1,'','','','',''],      // corrupted cell
 ['Wednesday',new Date(1899,11,30,17,0),'German',1,'','','','',''],
 ['Monday','10:00','Private',0,'','','','',''],           // LEGACY: private-as-language
 ['Saturday','14:00','Italian',0,'','','','','yes']]);     // NEW: real language + Private flag
const parsed=readWeeklySchedule_(control);
check('string time rule parsed', parsed.some(r=>r.time==='11:00'&&r.language==='English'), parsed);
check('Date-coerced 10:00 German rule RECOVERED', parsed.some(r=>r.time==='10:00'&&r.language==='German'), parsed.map(r=>r.time+'/'+r.language));
check('Date-coerced 17:00 German rule RECOVERED', parsed.some(r=>r.time==='17:00'&&r.language==='German'), null);
check('legacy "Private" language parses as isPrivate', parsed.some(r=>r.time==='10:00'&&r.isPrivate&&r.language==='Private'), parsed);
check('Private-flag row parses as isPrivate KEEPING its real language',
  parsed.some(r=>r.language==='Italian'&&r.time==='14:00'&&r.isPrivate===true), parsed);

/* --- offer writer: repairs German, regenerates English/Spanish, PRESERVES +
      migrates private (blanks the legacy "Private" label, keeps real languages) --- */
console.log('--- updateWeeklyScheduleToCurrentOffer ---');
updateWeeklyScheduleToCurrentOffer();
const after=ws.getRange(1,1,ws.getLastRow(),9).getDisplayValues();
check('German preserved with REPAIRED time (10:00, not 12/30/1899)',
  after.some(r=>r[2]==='German'&&r[1]==='10:00'), after.filter(r=>r[2]==='German').map(r=>r[1]));
check('English 11:00 Mon-Tue-Thu-Fri present',
  after.filter(r=>r[2]==='English'&&r[1]==='11:00').length===4, null);
check('Spanish 10:30 present', after.filter(r=>r[2]==='Spanish'&&r[1]==='10:30').length===4, null);

// Private rows are PRESERVED (not regenerated), and the legacy "Private" label
// is migrated off the Language column into the Private flag.
check('no legacy "Private" left in the Language column', !after.some(r=>/^private$/i.test(r[2])), after.map(r=>r[2]));
const legacyMig=after.find(r=>r[0]==='Monday'&&r[1]==='10:00'&&/^yes$/i.test(r[8]||''));
check('legacy private 10:00 migrated -> blank Language + Private=yes', !!legacyMig&&legacyMig[2]==='', legacyMig);
const italPriv=after.find(r=>r[0]==='Saturday'&&r[1]==='14:00');
check('Italian private preserved -> Italian language + Private=yes',
  !!italPriv&&italPriv[2]==='Italian'&&/^yes$/i.test(italPriv[8]||''), italPriv);
const reparsed=readWeeklySchedule_(control);
check('round-trip: all rows parse', reparsed.length===after.length-1, {rules:reparsed.length, rows:after.length-1});
const privParsed=reparsed.filter(r=>r.isPrivate);
check('exactly the two private rows survive as isPrivate', privParsed.length===2, privParsed);
check('private rows keep guidesNeeded 0 (stage no group tour)',
  privParsed.every(r=>Number(r.guidesNeeded)===0), privParsed);

/* --- "Hide from availability" (col H): keep the tour on web + portal, off the
      availability sheet (e.g. 4pm Italian) --- */
console.log('--- hideFromAvailability column ---');
ws.clear();   // reuse the real Weekly_Schedule tab readWeeklySchedule_ reads
ws.getRange(1,1,4,8).setValues([
 ['Day','Time','Language','Guides needed','Active from','Active until','Guide','Hide from availability'],
 ['Monday','11:00','English',1,'','','',''],
 ['Monday','16:00','Italian',1,'','','','yes'],     // hidden from availability
 ['Tuesday','16:00','Italian',1,'','','','']]);      // shown as normal
const hideRules=readWeeklySchedule_(control);
const ital=hideRules.filter(r=>r.language==='Italian');
check('flagged Italian row parses hideFromAvailability=true',
  ital.some(r=>r.day==='Monday'&&r.hideFromAvailability===true), ital);
check('unflagged Italian row stays visible (hideFromAvailability=false)',
  ital.some(r=>r.day==='Tuesday'&&r.hideFromAvailability===false), ital);
// The availability sheet builds from the FILTERED list; the portal/website use
// the full list, so the hidden tour still exists everywhere else.
const availVisible=hideRules.filter(r=>!r.hideFromAvailability);
check('availability sheet omits the hidden Mon 16:00 Italian',
  !availVisible.some(r=>r.day==='Monday'&&r.time==='16:00'&&r.language==='Italian'), availVisible.map(r=>r.day+r.time+r.language));
check('portal/website still see the hidden Mon 16:00 Italian',
  hideRules.some(r=>r.day==='Monday'&&r.time==='16:00'&&r.language==='Italian'), null);
check('English default (blank col H) is not hidden',
  hideRules.find(r=>r.language==='English').hideFromAvailability===false, null);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
