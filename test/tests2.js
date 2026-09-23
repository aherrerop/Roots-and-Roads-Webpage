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
// The offer writer now emits the 11-column layout (Tour Type inserted at col D):
// [Day, Time, Language, Tour Type, Guides needed, Active from, Active until,
//  Guide, Hide from availability, Private, Hide from website].
const after=ws.getRange(1,1,ws.getLastRow(),11).getDisplayValues();
check('German preserved with REPAIRED time (10:00, not 12/30/1899)',
  after.some(r=>r[2]==='German'&&r[1]==='10:00'), after.filter(r=>r[2]==='German').map(r=>r[1]));
check('Tour Type column added, header at col D', after[0][3]==='Tour Type', after[0]);
check('every managed offer row is tagged 3h',
  after.slice(1).filter(r=>r[2]==='English'||r[2]==='Spanish').every(r=>r[3]==='3h'),
  after.slice(1).filter(r=>r[2]==='English').map(r=>r[3]));
check('English 11:00 Mon-Tue-Thu-Fri present',
  after.filter(r=>r[2]==='English'&&r[1]==='11:00').length===4, null);
check('Spanish 10:30 present', after.filter(r=>r[2]==='Spanish'&&r[1]==='10:30').length===4, null);

// Private rows are PRESERVED (not regenerated), and the legacy "Private" label
// is migrated off the Language column into the Private flag (now col index 9).
check('no legacy "Private" left in the Language column', !after.some(r=>/^private$/i.test(r[2])), after.map(r=>r[2]));
const legacyMig=after.find(r=>r[0]==='Monday'&&r[1]==='10:00'&&/^yes$/i.test(r[9]||''));
check('legacy private 10:00 migrated -> blank Language + Private=yes', !!legacyMig&&legacyMig[2]==='', legacyMig);
const italPriv=after.find(r=>r[0]==='Saturday'&&r[1]==='14:00');
check('Italian private preserved -> Italian language + Private=yes',
  !!italPriv&&italPriv[2]==='Italian'&&/^yes$/i.test(italPriv[9]||''), italPriv);
check('preserved private row carries Tour Type 3h', !!italPriv&&italPriv[3]==='3h', italPriv);
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

/* --- "Hide from website" (col J): INDEPENDENT of Hide from availability (col H).
      Management can hide a slot from the website but keep it on the availability
      sheet + portal, or the reverse, or both, or neither. --- */
console.log('--- hideFromWebsite column (independent of hideFromAvailability) ---');
ws.clear();
ws.getRange(1,1,5,10).setValues([
 ['Day','Time','Language','Guides needed','Active from','Active until','Guide','Hide from availability','Private','Hide from website'],
 ['Monday','11:00','English',1,'','','','','',''],            // shows on BOTH
 ['Tuesday','16:00','Italian',1,'','','','yes','',''],        // hidden from availability ONLY
 ['Wednesday','16:00','Italian',1,'','','','','','yes'],      // hidden from website ONLY
 ['Thursday','16:00','Italian',1,'','','','yes','','yes']]);  // hidden from BOTH
const webRules=readWeeklySchedule_(control);
const gday=d=>webRules.find(r=>r.day===d)||{};
check('blank col J parses hideFromWebsite=false', gday('Monday').hideFromWebsite===false, gday('Monday'));
check('col J = yes parses hideFromWebsite=true', gday('Wednesday').hideFromWebsite===true, gday('Wednesday'));
check('Hide-from-website row is NOT hidden from availability (independent)', gday('Wednesday').hideFromAvailability===false, gday('Wednesday'));
check('Hide-from-availability row is NOT hidden from website (independent)', gday('Tuesday').hideFromWebsite===false, gday('Tuesday'));
check('a row flagged in BOTH cols is hidden from both', gday('Thursday').hideFromAvailability===true && gday('Thursday').hideFromWebsite===true, gday('Thursday'));
// The website builder (booking project) filters !hideFromWebsite; the availability
// sheet filters !hideFromAvailability — two independent lists off the same rows.
const webVisible=webRules.filter(r=>!r.hideFromWebsite);
check('website list omits the Hide-from-website rows (Wed + Thu)',
  !webVisible.some(r=>r.day==='Wednesday') && !webVisible.some(r=>r.day==='Thursday'), webVisible.map(r=>r.day));
check('website list KEEPS the Hide-from-availability-only row (Tue)',
  webVisible.some(r=>r.day==='Tuesday'), webVisible.map(r=>r.day));
const availVisible2=webRules.filter(r=>!r.hideFromAvailability);
check('availability list KEEPS the Hide-from-website-only row (Wed)',
  availVisible2.some(r=>r.day==='Wednesday'), availVisible2.map(r=>r.day));

console.log('--- Manager CLEAR sticks: the weekly default does not re-fill a cleared slot ---');
const clrSS = new __mock.MockSS('control-clear'); SpreadsheetApp._active = clrSS;
clrSS.insertSheet('Guides').getRange(1, 1, 2, 11).setValues([
  ['Guide','Active?','Seniority','English','German','Spanish','French','Italian','Manager','Email','Password'],
  ['Polina', true, 1, true, false, false, false, false, false, 'p@x.com','pw']]);
// A recurring Sunday 10:30 English slot whose USUAL guide (col G) is Polina.
clrSS.insertSheet('Weekly_Schedule').getRange(1, 1, 2, 10).setValues([
  ['Day','Time','Language','Guides needed','Active from','Active until','Guide','Hide from availability','Private','Hide from website'],
  ['Sunday','10:30','English',1,'','','Polina','','','']]);
__RRX = {}; __RRX.weekly = readWeeklySchedule_(clrSS);   // fresh rules, bypass cache
const _sunA = (function(){ const d=new Date(); d.setHours(12,0,0,0); while(d.getDay()!==0) d.setDate(d.getDate()+1); return Utilities.formatDate(d, Session.getScriptTimeZone(),'yyyy-MM-dd'); })();
const _sunB = (function(){ const d=new Date(_sunA+'T12:00:00'); d.setDate(d.getDate()+7); return Utilities.formatDate(d, Session.getScriptTimeZone(),'yyyy-MM-dd'); })();
const _s1 = [{ dateKey:_sunA, time:'10:30', minutes:630, language:'English', private:false, assigned:[], status:'Not assigned' }];
applyWeeklyDefaults_(_s1);
check('an UNTOUCHED Sunday 10:30 slot auto-fills the usual guide (Polina)', _s1[0].assigned.join()==='Polina', _s1[0].assigned);
const _s2 = [{ dateKey:_sunB, time:'10:30', minutes:630, language:'English', private:false, assigned:[], status:'Not assigned', cleared:true }];
applyWeeklyDefaults_(_s2);
check('a manager-CLEARED slot STAYS empty (usual guide does NOT re-fill it)', _s2[0].assigned.length===0, _s2[0].assigned);
check('guideForShift_ returns nobody for a cleared slot (not the weekly default)',
  guideForShift_([{ dateKey:_sunB, minutes:630, language:'English', private:false, assigned:[], cleared:true }], _sunB, '10:30', 'English', false)==='', 'expected empty');
check('the grid reader flags a "Not assigned (cleared)" cell as cleared, plain "Not assigned" as not',
  /cleared/i.test('Not assigned (cleared)')===true && /cleared/i.test('Not assigned')===false, null);

console.log('--- Availability horizon: rolling window, extended to the seasonal end date ---');
// mondayThis well before AVAILABILITY_UNTIL (2026-12-31): project out to the week
// containing Dec 31, not just the 4-week rolling window.
const _mon = dateOnly_(new Date('2026-09-21T12:00:00'));   // a Monday
const _wa = availabilityWeeksAhead_(_mon);
check('opens availability out to the Dec-31 week (>= 14 weeks from Sep 21, not just 4)',
  _wa >= 14, _wa);
check('never fewer than the rolling minimum', _wa >= ASSIGN_CFG.AVAILABILITY_WEEKS_AHEAD, _wa);
// The last projected Monday must reach the week that contains Dec 31, 2026.
const _lastMon = new Date(_mon); _lastMon.setDate(_mon.getDate() + 7 * _wa);
const _until = dateOnly_(new Date('2026-12-31T12:00:00'));
check('the furthest week tab covers Dec 31, 2026', _lastMon <= _until && (function(){
  const end = new Date(_lastMon); end.setDate(_lastMon.getDate()+6); return dateOnly_(end) >= _until; })(), _lastMon);
// Past the season end -> fall back to the rolling window only.
const _monLate = dateOnly_(new Date('2027-03-01T12:00:00'));
check('after the season end date, it is just the rolling window',
  availabilityWeeksAhead_(_monLate) === ASSIGN_CFG.AVAILABILITY_WEEKS_AHEAD, availabilityWeeksAhead_(_monLate));

console.log('--- makeSchedule preserves manager assignments OUTSIDE its weekly window ---');
const _today = dateOnly_(new Date());
const _fk = (o)=>{ const d=new Date(_today); d.setDate(d.getDate()+o); return Utilities.formatDate(d, Session.getScriptTimeZone(),'yyyy-MM-dd'); };
const _far = _fk(30), _past = _fk(-10), _inwin = _fk(1);
const _pShifts = [{ dateText:_inwin, time:'10:30', language:'English', isPrivate:false, privIndex:1, assignedGuides:['Carlos'], lockedGuides:['Carlos'] }];
const _pLocks = {};
_pLocks[lockKey_(_far, '17:00', 'English', false, 1)] = ['Polina'];       // far future -> PRESERVE
_pLocks[lockKey_(_past, '10:30', 'German', false, 1)] = ['Mar'];          // past -> skip
_pLocks[lockKey_(_inwin, '10:30', 'English', false, 1)] = ['Carlos'];     // already present -> no dupe
const _pCleared = {}; _pCleared[lockKey_(_far, '10:30', 'Spanish', false, 1)] = true;
preserveManagerGridState_(_pShifts, _pLocks, _pCleared, _today);
const _fS = (d,t,l)=>_pShifts.find(s=>s.dateText===d && s.time===t && s.language===l);
check('a far-future manager LOCK is preserved (not wiped by the weekly rebuild)',
  !!_fS(_far,'17:00','English') && _fS(_far,'17:00','English').assignedGuides.join()==='Polina' &&
  _fS(_far,'17:00','English').lockedGuides.join()==='Polina' && _fS(_far,'17:00','English').status==='OK', _fS(_far,'17:00','English'));
check('a far-future manager CLEAR is preserved with its (cleared) marker',
  !!_fS(_far,'10:30','Spanish') && _fS(_far,'10:30','Spanish').cleared===true &&
  _fS(_far,'10:30','Spanish').status==='Not assigned (cleared)', _fS(_far,'10:30','Spanish'));
check('a PAST manager assignment is NOT resurrected', !_fS(_past,'10:30','German'), 'should be absent');
check('an in-window assignment is not duplicated',
  _pShifts.filter(s=>s.dateText===_inwin && s.time==='10:30' && s.language==='English').length===1, _pShifts.length);

console.log('--- Guide vacations: single dates + ranges, and the usual-guide is not auto-filled ---');
const _vac = parseVacationRanges_('3/10, 19/10 - 29/10, 21/12 - 26/12 pending');
check('parses 3 entries (1 single + 2 ranges) from a mixed cell', _vac.length===3, _vac);
check('a single date is a 1-day range (3/10)', _vac[0].from===_vac[0].to && _vac[0].from.slice(5)==='10-03', _vac[0]);
check('a range spans its days (19/10 -> 29/10)', _vac[1].from.slice(5)==='10-19' && _vac[1].to.slice(5)==='10-29', _vac[1]);
check('trailing "pending" is ignored, the range still parses (21/12 -> 26/12)', _vac[2].from.slice(5)==='12-21' && _vac[2].to.slice(5)==='12-26', _vac[2]);
const _yr = _vac[1].from.slice(0,4);
const _gv = { name:'Francesca', vacations: parseVacationRanges_('19/10 - 29/10') };
check('isGuideOnVacation_: a date inside the range is blocked', isGuideOnVacation_(_gv, _yr+'-10-22')===true, _gv.vacations);
check('isGuideOnVacation_: a date outside the range is free', isGuideOnVacation_(_gv, _yr+'-10-30')===false, _gv.vacations);
check('isGuideOnVacation_: a guide with no vacation column is never blocked', isGuideOnVacation_({name:'X', vacations:[]}, _yr+'-10-22')===false, null);
// weeklyDefaultGuide_ integration: Polina is the usual Sunday 10:30 guide but is on
// vacation THAT Sunday -> no auto-fill; a Sunday she is NOT off -> Polina.
const _vSunA = (function(){ const d=new Date(); d.setHours(12,0,0,0); d.setDate(d.getDate()+7); while(d.getDay()!==0) d.setDate(d.getDate()+1); return d; })();
const _vSunB = new Date(_vSunA); _vSunB.setDate(_vSunA.getDate()+7);
const _vKey = d => Utilities.formatDate(d, Session.getScriptTimeZone(),'yyyy-MM-dd');
const _dmA = _vSunA.getDate()+'/'+(_vSunA.getMonth()+1);   // block sunA only
__RRX = {};
__RRX.guidesRaw = { header: ['Guide','Active?','Seniority','English','German','Spanish','French','Italian','Manager','Email','Password','Vacation dates'],
  rows: [['Polina', true, 1, true, false, false, false, false, false, 'p@x.com','pw', _dmA]] };
__RRX.weekly = [{ day:'Sunday', time:'10:30', language:'English', guidesNeeded:1, isPrivate:false, activeFrom:null, activeUntil:null, guide:'Polina', hideFromAvailability:false, hideFromWebsite:false }];
check('weeklyDefaultGuide_: usual guide is NOT auto-filled on their vacation Sunday', weeklyDefaultGuide_(_vKey(_vSunA),'10:30','English')==='', weeklyDefaultGuide_(_vKey(_vSunA),'10:30','English'));
check('weeklyDefaultGuide_: usual guide IS filled on a Sunday they are not off', weeklyDefaultGuide_(_vKey(_vSunB),'10:30','English')==='Polina', weeklyDefaultGuide_(_vKey(_vSunB),'10:30','English'));

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
