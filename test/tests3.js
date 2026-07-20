/* ===== Tests for orphan shifts, eligibility, moves, grid-edit locks ===== */
let pass=0, fail=0;
const check=(l,c,g)=>{ if(c){pass++;console.log('PASS  '+l);} else{fail++;console.log('FAIL  '+l+'  (got: '+JSON.stringify(g)+')');} };
const day=(o)=>{const d=new Date();d.setDate(d.getDate()+o);d.setHours(12,0,0,0);return d;};
const key=(d)=>Utilities.formatDate(d,null,'yyyy-MM-dd');

const control=new __mock.MockSS('control'); SpreadsheetApp._active=control;
const booking=new __mock.MockSS('booking');
__mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw']=booking;
const guides=control.insertSheet('Guides');
guides.getRange(1,1,4,10).setValues([
 ['Guide','Active?','Seniority','English','German','Spanish','French','Manager','Email','Password'],
 ['Albert',true,1,true,false,false,false,true,'a@x.com','pw'],
 ['Carlos',true,1,true,false,true,false,true,'c@x.com','pw'],
 ['Mario',true,2,true,true,false,false,'','m@x.com','pw']]);

/* --- 1. ORPHAN SHIFTS: a booking with no Weekly_Schedule rule still gets a tour --- */
console.log('--- Orphan shifts ---');
const d3=key(day(3));
const en=booking.insertSheet('English Tours');
en.getRange(1,1,3,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Web Guest','+34',4,d3,'4:30 PM','Website',0,'RRX1',''],           // 16:30 — NOT in any rule
 ['Priv Guest','+34',5,d3,'10:00 AM','GetYourGuide',99,'GYGP1','Private']]);
const shifts0=[{dateText:d3,time:'11:00',language:'English',isPrivate:false,privIndex:0,availableGuides:['Albert'],dateObj:day(3),dateTimeObj:day(3),day:'X',guidesNeeded:1,week:''}];
const availIdx={}; availIdx[d3+'|17:00']=['Carlos'];
const withOrphans=expandOrphanShifts_(shifts0, availIdx, day(0), day(10));
const orphan=withOrphans.find(s=>s.time==='16:30');
check('16:30 website tour becomes a shift', !!orphan, withOrphans.map(s=>s.time));
check('orphan flagged as extra', orphan&&orphan.extra===true, orphan);
check('orphan staffed from nearest availability (Carlos @17:00, 30min off)', orphan&&orphan.availableGuides.join()==='Carlos', orphan&&orphan.availableGuides);
check('private booking NOT duplicated as orphan', !withOrphans.some(s=>s.time==='10:00'&&!s.isPrivate), null);
check('existing 11:00 not duplicated', withOrphans.filter(s=>s.time==='11:00').length===1, null);

/* --- 2. ELIGIBILITY: conflicting guides are filtered from the dropdown --- */
console.log('--- Eligibility ---');
const gbl={English:['Albert','Carlos','Mario'],Spanish:['Carlos']};
const sched=[
 {dateKey:d3,minutes:17*60,language:'English',private:false,assigned:['Carlos']},
 {dateKey:d3,minutes:11*60,language:'English',private:false,assigned:[]}];
const busy=buildBusyMap_(sched);
const elig17es=eligibleGuidesForShift_({dateKey:d3,minutes:17*60,language:'Spanish',private:false},busy,gbl);
check('Carlos EXCLUDED from Spanish 17:00 (already on English 17:00)', elig17es.length===0, elig17es);
const elig11=eligibleGuidesForShift_({dateKey:d3,minutes:11*60,language:'English',private:false},busy,gbl);
check('Albert+Mario eligible for English 11:00; Carlos excluded (17:00 is <5h? no, 6h -> included)',
  elig11.indexOf('Albert')>-1 && elig11.indexOf('Mario')>-1 && elig11.indexOf('Carlos')>-1, elig11);
const elig15=eligibleGuidesForShift_({dateKey:d3,minutes:15*60,language:'English',private:false},busy,gbl);
check('Carlos excluded at 15:00 (2h from his 17:00)', elig15.indexOf('Carlos')===-1, elig15);
/* self-exclusion: guide keeps own shift in dropdown */
const eligSelf=eligibleGuidesForShift_({dateKey:d3,minutes:17*60,language:'English',private:false},busy,gbl);
check('Carlos still eligible for HIS OWN 17:00 shift', eligSelf.indexOf('Carlos')>-1, eligSelf);

/* --- 3. MOVE BOOKING between language tabs --- */
console.log('--- Language move ---');
const de=booking.insertSheet('German Tours');
de.getRange(1,1,2,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Hans','+49',2,d3,'10:00 AM','GetYourGuide',30,'GYGDE1','']]);
const mv=moveBookingRowBetweenTabs_('GYGDE1','German','English');
check('move reports ok', mv.ok===true&&mv.moved===true, mv);
check('row gone from German', de.getLastRow()===1, de.getLastRow());
const enRows=en.getRange(2,1,en.getLastRow()-1,9).getValues();
const moved=enRows.find(r=>r[7]==='GYGDE1');
check('row present in English with traceable note', moved && /moved from German/.test(moved[8]), moved&&moved[8]);
const mv2=moveBookingRowBetweenTabs_('GYGDE1','German','English');
check('repeat move fails cleanly (not found), no duplicate', mv2.ok===false && enRows.filter(r=>r[7]==='GYGDE1').length===1, mv2);

/* --- 4. GRID EDIT: hand-typed name becomes a bold LOCK + validated --- */
console.log('--- Manual grid edit -> lock + validation ---');
const gsh=control.insertSheet('Schedule_Spanish');
gsh.getRange(1,1,3,2).setValues([
 ['Spanish schedule ('+d3+' to '+d3+')',''],
 ['Date','10:30'],
 [Utilities.formatDate(day(3),null,'EEE MMM d'),'Not assigned']]);
const cell=gsh.getRange(3,2);
cell.setValue('Albert');   // Albert does NOT speak Spanish
handleScheduleEdit({range:cell});
check('typed name auto-bolded (lock)', (()=>{const rt=gsh.getRange(3,2).getRichTextValues()[0][0];return rt.getRuns().some(r=>r.getTextStyle().isBold());})(), null);
check('incompatibility flagged in cell note', /does not speak Spanish/.test(cell.getNote()), cell.getNote());
cell.setValue('Carlos');   // Carlos speaks Spanish, no conflicts
handleScheduleEdit({range:cell});
check('valid edit -> no warning note', cell.getNote()==='', cell.getNote());
const locks=readLockedAssignments_(control);
check('hand edit readable as lock by makeSchedule', Object.keys(locks).some(k=>locks[k].join()==='Carlos'), locks);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
