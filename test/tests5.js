/* ===== Unassigned = tours that ALREADY RAN with no guide ===== */
let pass=0, fail=0;
const check=(l,c,g)=>{ if(c){pass++;console.log('PASS  '+l);} else{fail++;console.log('FAIL  '+l+'  (got: '+JSON.stringify(g)+')');} };
const key=(d)=>Utilities.formatDate(d,null,'yyyy-MM-dd');
const dayOff=(o)=>{const d=new Date();d.setDate(d.getDate()+o);d.setHours(12,0,0,0);return d;};

const control=new __mock.MockSS('control'); SpreadsheetApp._active=control;
const ledger=new __mock.MockSS('ledger'); __mock.SS_BY_ID['LEDGER']=ledger; __mock.PROPS['LEDGER_ID']='LEDGER';
const booking=new __mock.MockSS('booking');
__mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw']=booking;

const PAST=key(dayOff(-2)), FUTURE=key(dayOff(+5));

// Completed Log: 3 past tours (one had a guide, one did not, one private no guide)
const cl=booking.insertSheet('Completed Log');
cl.getRange(1,1,4,12).setValues([
 ['Date','Time','Language','Name','Phone','Adults','Children','Source','Income','Booking ID','Notes','Logged'],
 [PAST,'11:00 AM','English','HadGuide','+1',2,0,'GetYourGuide',27,'GYG-OK','','x'],
 [PAST,'5:00 PM','English','NoGuide','+1',3,0,'GetYourGuide',40,'GYG-ORPHAN','','x'],
 [PAST,'10:00 AM','English','PrivNoGuide','+1',5,0,'Viator',99,'BR-PRIV','Private','x']]);

// Schedule grid covering the PAST date: 11:00 has Albert, 17:00 nobody, 10:00 private nobody
const g=control.insertSheet('Schedule_English');
g.getRange(1,1,3,4).setValues([
 ['English schedule ('+PAST+' to '+FUTURE+')','','',''],
 ['Date','11:00','17:00','10:00 · Private'],
 [Utilities.formatDate(dayOff(-2),null,'EEE MMM d'),'Albert','Not assigned','Not assigned']]);

// Active tabs contain FUTURE bookings — these must NOT appear.
const en=booking.insertSheet('English Tours');
en.getRange(1,1,3,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Future Guest','+1',2,FUTURE,'11:00 AM','GetYourGuide',27,'GYG-FUTURE',''],
 ['Future Priv','+1',4,FUTURE,'10:00 AM','Viator',99,'BR-FUTURE','Private']]);

const n=rebuildUnassignedLedger_();
const sh=ledger.getSheetByName('Unassigned');
const ids=(lastDataRow_(sh)>=3? sh.getRange(3,1,lastDataRow_(sh)-2,11).getValues().map(r=>r[10]) : []);

console.log('Unassigned rows -> '+JSON.stringify(ids));
check('FUTURE bookings are NOT listed', ids.indexOf('GYG-FUTURE')===-1 && ids.indexOf('BR-FUTURE')===-1, ids);
check('past tour WITH a guide is NOT listed', ids.indexOf('GYG-OK')===-1, ids);
check('past tour with NO guide IS listed', ids.indexOf('GYG-ORPHAN')>-1, ids);
check('past PRIVATE with no guide IS listed', ids.indexOf('BR-PRIV')>-1, ids);
check('row count matches', n===2, n);
check('title row present', /RAN WITH NO GUIDE/.test(String(sh.getRange(1,1).getValue())), sh.getRange(1,1).getValue());
check('headers on row 2', sh.getRange(2,1).getValue()==='Date', sh.getRange(2,1).getValue());

// idempotent rebuild
const first=lastDataRow_(sh);
rebuildUnassignedLedger_(); rebuildUnassignedLedger_();
check('rebuild is idempotent (no growth)', lastDataRow_(sh)===first, {first, after:lastDataRow_(sh)});

// portal still hides past tours from guides
check('readSchedule_() default still excludes past', !readSchedule_().some(s=>s.dateKey===PAST), readSchedule_().map(s=>s.dateKey));
check('readSchedule_({includePast}) sees the past shift', readSchedule_({includePast:true}).some(s=>s.dateKey===PAST), null);

// --- a completed tour on a date the grids do NOT cover must be IGNORED ---
const OLD=key(dayOff(-10));
cl.getRange(5,1,1,12).setValues([[OLD,'11:00 AM','English','AncientGuest','+1',2,0,'GetYourGuide',27,'GYG-ANCIENT','','x']]);
rebuildUnassignedLedger_();
const ids2=(lastDataRow_(sh)>=3? sh.getRange(3,1,lastDataRow_(sh)-2,11).getValues().map(r=>r[10]) : []);
check('tour outside grid coverage is NOT accused', ids2.indexOf('GYG-ANCIENT')===-1, ids2);
check('in-coverage orphan still reported', ids2.indexOf('GYG-ORPHAN')>-1, ids2);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
