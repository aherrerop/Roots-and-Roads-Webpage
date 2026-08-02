/* ===== Italian + French support: control side (guides, eligibility, moves) ===== */
let pass=0, fail=0;
const check=(l,c,g)=>{ if(c){pass++;console.log('PASS  '+l);} else{fail++;console.log('FAIL  '+l+'  (got: '+JSON.stringify(g)+')');} };
const day=(o)=>{const d=new Date();d.setDate(d.getDate()+o);d.setHours(12,0,0,0);return d;};
const key=(d)=>Utilities.formatDate(d,null,'yyyy-MM-dd');

const control=new __mock.MockSS('control'); SpreadsheetApp._active=control;
const booking=new __mock.MockSS('booking');
__mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw']=booking;

/* Guides tab: language columns in a DELIBERATELY non-standard order (French
   before Italian, Italian last) to prove readGuides_ resolves BY HEADER NAME. */
const guides=control.insertSheet('Guides');
guides.getRange(1,1,6,11).setValues([
 ['Guide','Active?','Seniority','English','German','Spanish','French','Italian','Manager','Email','Password'],
 ['Albert', true,1,true, false,false,false,false,true, 'a@x.com','pw'],
 ['Giulia', true,1,false,false,false,false,true, false,'g@x.com','pw'],   // Italian only
 ['Pierre', true,1,false,false,false,true, false,false,'p@x.com','pw'],   // French only
 ['Sofia',  true,2,false,false,true, false,true, false,'s@x.com','pw'],   // Spanish + Italian
 ['Hans',   true,2,true, true, false,false,false,false,'h@x.com','pw']]);

console.log('--- Guide recognition (Italian / French) ---');
const gs=readGuides_(control);
const g=n=>gs.find(x=>x.name===n);
check('Italian guide recognized (Giulia)', g('Giulia').languages.Italian===true, g('Giulia').languages);
check('French guide recognized (Pierre)', g('Pierre').languages.French===true, g('Pierre').languages);
check('multi-language guide (Sofia: Spanish+Italian)', g('Sofia').languages.Spanish===true && g('Sofia').languages.Italian===true, g('Sofia').languages);
check('NO silent fallback: English-only guide is not Italian/French', g('Albert').languages.Italian===false && g('Albert').languages.French===false, g('Albert').languages);
check('regression: English/German still read correctly', g('Hans').languages.English===true && g('Hans').languages.German===true, g('Hans').languages);
check('Italian guide does NOT also speak English (no fallback)', g('Giulia').languages.English===false, g('Giulia').languages);

console.log('--- Eligibility: Italian / French, no cross-language fallback ---');
const d3=key(day(3));
const gbl={English:['Albert','Hans'],Spanish:['Sofia'],Italian:['Giulia','Sofia'],French:['Pierre']};
const busy=buildBusyMap_([]);
const eligIt=eligibleGuidesForShift_({dateKey:d3,minutes:11*60,language:'Italian',private:false},busy,gbl);
check('Italian 11:00 eligible = Italian speakers only', eligIt.slice().sort().join()==='Giulia,Sofia', eligIt);
check('English-only guide NOT eligible for Italian shift', eligIt.indexOf('Albert')===-1, eligIt);
const eligFr=eligibleGuidesForShift_({dateKey:d3,minutes:17*60,language:'French',private:false},busy,gbl);
check('French 17:00 eligible = Pierre', eligFr.join()==='Pierre', eligFr);
check('non-French guide NOT eligible for French shift', eligFr.indexOf('Giulia')===-1, eligFr);
const eligEn=eligibleGuidesForShift_({dateKey:d3,minutes:11*60,language:'English',private:false},busy,gbl);
check('regression: English 11:00 eligible = Albert,Hans', eligEn.slice().sort().join()==='Albert,Hans', eligEn);

console.log('--- Booking routing: move into Italian / French Tours ---');
const it=booking.insertSheet('Italian Tours');
it.getRange(1,1,1,9).setValues([['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes']]);
const fr=booking.insertSheet('French Tours');
fr.getRange(1,1,1,9).setValues([['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes']]);
const en=booking.insertSheet('English Tours');
en.getRange(1,1,2,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Marco Rossi','+39',2,d3,'11:00 AM','GetYourGuide',40,'GYGIT1','']]);
const mvIt=moveBookingRowBetweenTabs_('GYGIT1','English','Italian');
check('move to Italian ok', mvIt.ok===true && mvIt.moved===true, mvIt);
check('row now in Italian Tours', it.getLastRow()===2 && String(it.getRange(2,8).getValue())==='GYGIT1', it.getRange(2,1,1,9).getValues());
en.getRange(2,1,1,9).setValues([['Pierre Dupont','+33',3,d3,'5:00 PM','Viator',70,'BRFR1','']]);
const mvFr=moveBookingRowBetweenTabs_('BRFR1','English','French');
check('move to French ok', mvFr.ok===true && mvFr.moved===true, mvFr);
check('row now in French Tours', fr.getLastRow()===2 && String(fr.getRange(2,8).getValue())==='BRFR1', fr.getRange(2,1,1,9).getValues());

console.log('--- Live orphan surfacing: bookings show even without a grid slot ---');
const tomorrow=Utilities.formatDate(day(1),null,'yyyy-MM-dd');
let sched=[];
const bbk={};
bbk[shiftKey_(tomorrow,660,'Italian')]=[{name:'Testy',guests:1,note:'Test'}];      // Italian, no grid
bbk[shiftKey_(tomorrow,16*60+30,'English')]=[{name:'Polina guest',guests:2,note:''}]; // English 16:30 orphan
appendOrphanBookingShifts_(sched,bbk);
const itLive=sched.find(s=>s.language==='Italian'&&s.minutes===660);
check('Italian orphan surfaced as a live shift', !!itLive && itLive.status==='Not assigned', sched);
check('surfaced Italian shift has correct time/label', itLive&&itLive.time==='11:00'&&itLive.timeLabel==='11:00 AM', itLive);
check('English orphan surfaced too (regression)', sched.some(s=>s.language==='English'&&s.minutes===16*60+30), sched);
// No duplicate when the grid already has the shift.
let sched2=[{dateKey:tomorrow,minutes:660,language:'Italian',time:'11:00',private:false,assigned:['Giulia'],status:'OK'}];
appendOrphanBookingShifts_(sched2,bbk);
check('no duplicate when grid already has that shift', sched2.filter(s=>s.language==='Italian'&&s.minutes===660).length===1, sched2);
// Past bookings are not surfaced.
const longAgo=Utilities.formatDate(day(-5),null,'yyyy-MM-dd');
let sched3=[]; const bbk3={}; bbk3[shiftKey_(longAgo,660,'Italian')]=[{name:'x',guests:1,note:''}];
appendOrphanBookingShifts_(sched3,bbk3);
check('past booking not surfaced', sched3.length===0, sched3);

// Ordering: a surfaced 11:00 orphan must sort ABOVE an existing 17:00 grid shift
// (same comparator apiTours_ applies after appending).
const byTime=(a,b)=>(a.dateKey<b.dateKey?-1:a.dateKey>b.dateKey?1:0)||(a.minutes-b.minutes)||
  String(a.language).localeCompare(String(b.language));
let sched4=[
  {dateKey:tomorrow,minutes:17*60,language:'English',time:'17:00',private:false,assigned:['Carlos'],status:'OK'},
  {dateKey:tomorrow,minutes:17*60,language:'German',time:'17:00',private:false,assigned:[],status:'Not assigned'}];
const bbk4={};
bbk4[shiftKey_(tomorrow,660,'Italian')]=[{name:'Testy',guests:1,note:'Test'}];
bbk4[shiftKey_(tomorrow,660,'French')]=[{name:'Fr guest',guests:1,note:''}];
appendOrphanBookingShifts_(sched4,bbk4);
sched4.sort(byTime);
check('surfaced orphans carry minutes (sort key)', sched4.every(s=>Number.isFinite(s.minutes)), sched4.map(s=>s.minutes));
check('11:00 IT/FR sort ABOVE 17:00 EN/DE', sched4.map(s=>s.language).join(',')==='French,Italian,English,German', sched4.map(s=>s.timeLabel+' '+s.language));

console.log('--- Guide visibility: assigned + takeable unassigned only ---');
const giulia={name:'Giulia',languages:{English:false,German:false,Spanish:false,French:false,Italian:true}};
const mineG=[{dateKey:tomorrow,minutes:11*60,language:'Italian',assigned:['Giulia']}];
const pool=[
  {dateKey:tomorrow,minutes:11*60,language:'Italian',assigned:['Giulia'],status:'OK'},      // mine
  {dateKey:tomorrow,minutes:17*60,language:'English',assigned:['Carlos'],status:'OK'},      // assigned, other lang
  {dateKey:tomorrow,minutes:17*60,language:'Italian',assigned:[],status:'Not assigned'},    // open, my lang, no clash
  {dateKey:tomorrow,minutes:12*60,language:'Italian',assigned:[],status:'Not assigned'},    // open, my lang, CLASHES (1h)
  {dateKey:tomorrow,minutes:17*60,language:'German',assigned:[],status:'Not assigned'}];    // open, not my lang
const seenG=visibleShiftsForGuide_(pool,mineG,giulia,false);
check('guide sees assigned tours (any language)', seenG.filter(s=>s.assigned.length).length===2, seenG);
check('guide sees open tour in THEIR language with no clash',
  seenG.some(s=>s.language==='Italian'&&s.minutes===17*60&&!s.assigned.length), seenG);
check('guide does NOT see open tour clashing with their own shift',
  !seenG.some(s=>s.minutes===12*60), seenG.map(s=>s.minutes));
check('guide does NOT see open tour in a language they do not run',
  !seenG.some(s=>s.language==='German'), seenG.map(s=>s.language));
check('manager sees everything unfiltered', visibleShiftsForGuide_(pool,mineG,giulia,true).length===pool.length, null);

console.log('--- Availability grid: times stay TEXT (no 12/30/1899) ---');
const gss=new __mock.MockSS('guides'); __mock.SS_BY_ID['GUIDEFILE']=gss;
const wk=gss.insertSheet('Week 99');
const monday=day(1);
const rules=[{day:fullDayName_(monday),time:'10:00',language:'English',guidesNeeded:1,activeFrom:null,activeUntil:null},
             {day:fullDayName_(monday),time:'17:00',language:'Italian',guidesNeeded:1,activeFrom:null,activeUntil:null}];
const weekDates=[{dateKey:key(monday),dateObj:monday,dayName:fullDayName_(monday),shortLabel:shortDateLabel_(monday)}];
rebuildAvailabilityWeekSheet_(wk,weekDates,rules,['Giulia'],{});
const hdr=wk.getRange(4,2,1,2).getValues()[0];
check('time header written as text "10:00" (not a Date)', hdr[0]==='10:00' && !(hdr[0] instanceof Date), hdr);
check('second time header text too', hdr[1]==='17:00' && !(hdr[1] instanceof Date), hdr);

console.log('--- Availability is IDENTICAL week to week (no vanishing 10:00) ---');
// Weekly_Schedule holds only real tours; private slots come from ASSIGN_CFG.
const offer=[{day:'Monday',time:'11:00',language:'English',guidesNeeded:1,activeFrom:null,activeUntil:null}];
const allRules=offer.concat(privateAvailabilityRules_());
const mkWeek=(name,mondayOffset)=>{
  const d=day(mondayOffset); const sh=gss.insertSheet(name);
  rebuildAvailabilityWeekSheet_(sh,[{dateKey:key(d),dateObj:d,dayName:'Monday',shortLabel:shortDateLabel_(d)}],
    allRules,['Giulia'],{});
  return sh.getRange(4,2,1,sh.getLastColumn()-1).getValues()[0].filter(String);
};
const wA=mkWeek('Week A',7), wB=mkWeek('Week B',14), wC=mkWeek('Week C',21);
check('Monday offers 10:00 (private slot) in week A', wA.indexOf('10:00')!==-1, wA);
check('Monday offers 10:00 in week B TOO (was disappearing)', wB.indexOf('10:00')!==-1, wB);
check('Monday offers 10:00 in week C as well', wC.indexOf('10:00')!==-1, wC);
check('same weekday = identical columns across all three weeks',
  wA.join()===wB.join() && wB.join()===wC.join(), {wA,wB,wC});
check('private slots present without any Weekly_Schedule "Private" row',
  wA.indexOf('10:30')!==-1 && wA.indexOf('17:00')!==-1, wA);

console.log('--- Ledger: free-tour commission is per-person, per-platform ---');
const rates={paid:10, free:6, privatePay:75,
  freeCommissions:{guruwalk:4.7, 'free tour':2, website:0, '':0}, paidSources:['Viator','GetYourGuide','Airbnb']};
PORTAL._paidSources=rates.paidSources;
const gw=computeMoney_('Guruwalk',4,false,0,rates);
check('Guruwalk free: guide owes 6x4=24', gw.theyOwe===24, gw);
check('Guruwalk free: R&R makes (6-4.7)x4 = 5.2 (was 19.3)', Math.abs(gw.rrMakes-5.2)<0.001, gw.rrMakes);
const ft=computeMoney_('Free Tour',3,false,0,rates);
check('Free Tour commission 2/person: R&R makes (6-2)x3=12', Math.abs(ft.rrMakes-12)<0.001, ft.rrMakes);
const web=computeMoney_('Website',2,false,0,rates);
check('Website 0 commission: R&R makes 6x2=12', Math.abs(web.rrMakes-12)<0.001, web.rrMakes);
const gyg=computeMoney_('GetYourGuide',2,false,27,rates);
check('Paid tour unchanged: R&R makes income-weOwe = 27-20 = 7', Math.abs(gyg.rrMakes-7)<0.001, gyg);

console.log('--- Assignment always works: grid grows (tab/column/row created) ---');
const ac=new __mock.MockSS('control-assign'); SpreadsheetApp._active=ac;
const d1=key(day(2));
let res=writeAssignmentToGrid_('Italian',d1,'10:00',false,1,'Giulia');   // no Schedule_Italian yet
check('assign to a brand-new language/slot succeeds', res.ok===true && res.assigned==='Giulia', res);
check('Schedule_Italian tab was created', !!ac.getSheetByName('Schedule_Italian'), null);
let res2=writeAssignmentToGrid_('Italian',d1,'17:00',false,1,'Marco');   // new time column, same date
check('assigning a new TIME adds a column, still ok', res2.ok===true, res2);
const itSched=readSchedule_().filter(s=>s.language==='Italian');
check('both created slots read back from the grid', itSched.length===2 &&
  itSched.some(s=>s.time==='10:00'&&s.assigned.join()==='Giulia') &&
  itSched.some(s=>s.time==='17:00'&&s.assigned.join()==='Marco'), itSched);

console.log('--- Year resolver: a Jul row in an Aug-titled grid must NOT jump to next year ---');
const augAnchor=gridAnchor_('Italian schedule (2026-08-07 to 2026-08-07)');
check('Jul 31 in an Aug-titled grid -> 2026-07-31 (was 2027, the sticking bug)',
  gridLabelToKey_('Fri Jul 31',augAnchor)==='2026-07-31', gridLabelToKey_('Fri Jul 31',augAnchor));
check('Aug 7 in the same grid -> 2026-08-07', gridLabelToKey_('Fri Aug 7',augAnchor)==='2026-08-07', null);
const decAnchor=gridAnchor_('English schedule (2026-12-28 to 2027-01-05)');
check('Jan 3 across a Dec->Jan boundary -> 2027-01-03', gridLabelToKey_('Sat Jan 3',decAnchor)==='2027-01-03', gridLabelToKey_('Sat Jan 3',decAnchor));
check('Dec 30 across the same boundary -> 2026-12-30', gridLabelToKey_('Wed Dec 30',decAnchor)==='2026-12-30', null);

console.log('--- Assignment STICKS + no duplicate rows + chronological order ---');
const gc=new __mock.MockSS('control-grid'); SpreadsheetApp._active=gc;
const fmtLabel=k=>Utilities.formatDate(new Date(k+'T12:00:00'),Session.getScriptTimeZone(),'EEE MMM d');
const gA=key(day(3)), gB=key(day(10));   // gA earlier than gB, both upcoming
// Seed a messy grid like production: title anchored to the LATER date, a 17:00
// column and only the later row — exactly the shape that made assigns append.
const gsh=gc.insertSheet('Schedule_Italian');
gsh.getRange(1,1).setValue('Italian schedule ('+gB+' to '+gB+')');
gsh.getRange(2,1,1,2).setValues([['Date','17:00']]);
gsh.getRange(3,1).setValue(fmtLabel(gB));
const GA1=writeAssignmentToGrid_('Italian',gA,'10:00',false,1,'Miguel');
check('assigning the earlier date returns Miguel', GA1.assigned==='Miguel', GA1);
let gv=gsh.getDataRange().getDisplayValues();
const gh1=parseGridTimeHeader_(gv[1][1]), gh2=parseGridTimeHeader_(gv[1][2]);
check('new 10:00 column is inserted BEFORE 17:00 (chronological)',
  !!gh1&&!!gh2&&gh1.time==='10:00'&&gh2.time==='17:00', gv[1]);
check('earlier date row is inserted BEFORE the later one', gv[2][0]===fmtLabel(gA)&&gv[3][0]===fmtLabel(gB), [gv[2][0],gv[3][0]]);
check('Miguel is written into the earlier date / 10:00 cell', gv[2][1]==='Miguel', gv[2][1]);
// Re-assign the SAME shift: must update in place, never add a second row.
writeAssignmentToGrid_('Italian',gA,'10:00',false,1,'Miguel');
gv=gsh.getDataRange().getDisplayValues();
check('re-assigning the same shift does NOT create a duplicate row',
  gv.filter((r,i)=>i>=2 && r[0]===fmtLabel(gA)).length===1, gv.map(r=>r[0]));
// The portal must read it back at the correct date, assigned to Miguel.
const gp=readSchedule_().filter(s=>s.language==='Italian'&&s.dateKey===gA&&s.time==='10:00');
check('portal reads the assignment at the RIGHT date, assigned to Miguel',
  gp.length===1 && gp[0].assigned.indexOf('Miguel')!==-1, gp);
check('the title self-healed to span the real date range',
  /\(\d{4}-\d{2}-\d{2} to \d{4}-\d{2}-\d{2}\)/.test(String(gsh.getRange(1,1).getDisplayValue())), gsh.getRange(1,1).getDisplayValue());
// Re-assigning a DIFFERENT guide (Albert -> Carlos) must reflect immediately on
// the very next read, on the same row, with no trace of the old guide.
writeAssignmentToGrid_('Italian',gA,'10:00',false,1,'Albert');
let gpA=readSchedule_().filter(s=>s.language==='Italian'&&s.dateKey===gA&&s.time==='10:00');
check('after assigning Albert, the shift reads back as Albert', gpA.length===1 && gpA[0].assigned.join()==='Albert', gpA);
writeAssignmentToGrid_('Italian',gA,'10:00',false,1,'Carlos');
let gpC=readSchedule_().filter(s=>s.language==='Italian'&&s.dateKey===gA&&s.time==='10:00');
check('reassign Albert->Carlos reflects on the NEXT read (no delay)', gpC.length===1 && gpC[0].assigned.join()==='Carlos', gpC);
check('the old guide Albert is gone (not merged/duplicated)', gpC[0].assigned.indexOf('Albert')===-1, gpC[0].assigned);
gv=gsh.getDataRange().getDisplayValues();
check('still a SINGLE grid row for the reassigned shift', gv.filter((r,i)=>i>=2 && r[0]===fmtLabel(gA)).length===1, gv.map(r=>r[0]));

console.log('--- Portal keeps a tour until the evening of its day ---');
check('same-day morning tour is NOT over', shiftIsOver_(key(new Date()),10*60)===false, null);
check('a tour on a past day IS over', shiftIsOver_(key(day(-1)),10*60)===true, null);

console.log('--- Ledger fallback: a completed tour keeps its reservations (durable) ---');
const lc=new __mock.MockSS('control-ledger'); SpreadsheetApp._active=lc;
lc.insertSheet('Guides').getRange(1,1,2,11).setValues([
 ['Guide','Active?','Seniority','English','German','Spanish','French','Italian','Manager','Email','Password'],
 ['Marco',true,1,false,false,false,false,true,false,'m@x.com','pw']]);
const led=new __mock.MockSS('ledger-fb'); __mock.SS_BY_ID['LEDFB']=led; __mock.PROPS['LEDGER_ID']='LEDFB';
const gt=led.insertSheet('Marco');
gt.getRange(1,1,2,LEDGER_HEADERS.length).setNumberFormat('@');
gt.getRange(1,1,2,LEDGER_HEADERS.length).setValues([
 LEDGER_HEADERS,
 ['2026-07-28','Tue','11:00','Italian','Guest Uno','+3912345','GetYourGuide',2,1,2,40,0,7,'Paid','GYGIT9','2026-07-28 13:00']]);
const lr=readLedgerReservations_();
const lk=shiftKey_('2026-07-28',11*60,'Italian');
check('ledger reservation indexed by shift key', !!lr[lk] && lr[lk].length===1, Object.keys(lr));
const rb=(lr[lk]||[])[0]||{};
check('reservation carries name/phone/guests/children/checkedIn from ledger',
  rb.name==='Guest Uno' && rb.phone==='+3912345' && Number(rb.guests)===2 && Number(rb.children)===1 && Number(rb.checkedIn)===2, rb);
check('no duplicate booking ids in the ledger index', (lr[lk]||[]).filter(b=>b.bookingId==='GYGIT9').length===1, lr[lk]);

console.log('--- Completed Log fallback: un-checked-in guests still show ---');
const bkCL=new __mock.MockSS('booking-cl'); __mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw']=bkCL;
const cl=bkCL.insertSheet('Completed Log');
cl.getRange(1,1,2,12).setNumberFormat('@');
cl.getRange(1,1,2,12).setValues([
 ['Date','Time','Language','Name','Phone','Adults','Children','Source','Income','Booking ID','Notes','Logged'],
 ['2026-07-28','11:00','French','No Show Guy','+33999','2','1','Viator','0','BRFR9','','2026-07-28 13:00']]);
const cr=readCompletedLogReservations_();
const clk=shiftKey_('2026-07-28',11*60,'French');
check('completed-log reservation indexed by shift', !!cr[clk] && cr[clk].length===1, Object.keys(cr));
const clb=(cr[clk]||[])[0]||{};
check('completed-log carries name/phone/guests/children (checked-in unknown)',
  clb.name==='No Show Guy'&&clb.phone==='+33999'&&Number(clb.guests)===2&&Number(clb.children)===1, clb);

console.log('--- Per-booking notes: written to the BookingSheet (col J) ---');
const bnss=new __mock.MockSS(PORTAL.BOOKING_SHEET_ID); __mock.SS_BY_ID[PORTAL.BOOKING_SHEET_ID]=bnss;
const bnEn=bnss.insertSheet('English Tours');
bnEn.getRange(1,1,3,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Kai','+49',2,'2026-08-05','11:00 AM','Guruwalk',0,'GW1',''],
 ['Edgar','+49',1,'2026-08-05','11:00 AM','GetYourGuide',15,'GYG1','']]);
const wn=writeBookingNote_('GYG1','English','Bringing a wheelchair');
check('note write succeeds for a real booking', wn.ok===true && wn.note==='Bringing a wheelchair', wn);
check('note lands in column J of the right row', String(bnEn.getRange(3,10).getValue())==='Bringing a wheelchair', bnEn.getRange(3,10).getValue());
check('column J header set to Note', String(bnEn.getRange(1,10).getValue())==='Note', bnEn.getRange(1,10).getValue());
check('note write fails for an unknown booking id', writeBookingNote_('NOPE','English','x').ok===false, null);
check('readBookingsIndex_ surfaces the per-booking note as manualNote',
  (readBookingsIndex_()[shiftKey_('2026-08-05',660,'English')]||[]).find(b=>b.bookingId==='GYG1').manualNote==='Bringing a wheelchair', null);
// The note rides into the ledger row (last column).
const lrow=makeLedgerRow_({dateKey:'2026-08-05',day:'Wednesday',timeLabel:'11:00 AM',language:'English',
  bookingName:'Edgar',phone:'+49',source:'GetYourGuide',guests:1,children:0,checkedIn:1,
  weOwe:10,theyOwe:0,rrMakes:0,type:'Paid',bookingId:'GYG1',note:'Bringing a wheelchair'});
check('ledger row carries the note in the Note column', lrow[LEDGER_NOTE_COL]==='Bringing a wheelchair' && lrow.length===LEDGER_HEADERS.length, lrow);
check('ledger Updated column is still the timestamp (not the note)', /\d{4}-\d{2}-\d{2}/.test(String(lrow[LEDGER_UPDATED_COL])), lrow[LEDGER_UPDATED_COL]);

console.log('--- #4 Weekly_Schedule rules surface in the portal immediately ---');
const cw=new __mock.MockSS('control-weekly'); SpreadsheetApp._active=cw;
const _plus=n=>{ const d=new Date(); d.setDate(d.getDate()+n); return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd'); };
const tgtKey=_plus(7), tgtDay=dayNameFromKey_(tgtKey);
cw.insertSheet('Weekly_Schedule').getRange(1,1,4,6).setValues([
 ['Day','Time','Language','Guides needed','Active from','Active until'],
 [tgtDay,'11:00','Italian',1,'',''],
 [tgtDay,'17:00','Private',1,'',''],               // availability window: must NOT surface
 ['Monday','09:00','French',1,'2000-01-01','2000-01-02']]); // expired: must NOT surface
const wsched=[];
appendWeeklyScheduleShifts_(wsched);
const itShift=wsched.find(s=>s.language==='Italian'&&s.dateKey===tgtKey);
check('Italian 11:00 rule surfaced as an unassigned shift',
  !!itShift && itShift.status==='Not assigned' && itShift.assigned.length===0 && !itShift.private, wsched);
check('Private availability row is NOT surfaced', !wsched.some(s=>/private/i.test(s.language)), wsched);
check('expired French rule is NOT surfaced', !wsched.some(s=>s.language==='French'), wsched);
const wsched2=[{dateKey:tgtKey, minutes:timeToMinutes_('11:00'), language:'Italian', private:false, assigned:['Someone'], status:'OK'}];
appendWeeklyScheduleShifts_(wsched2);
check('a shift already present is not duplicated by the rule',
  wsched2.filter(s=>s.language==='Italian'&&s.dateKey===tgtKey).length===1, wsched2);

console.log('--- #5 Management closes a schedule (durable, source-agnostic hide) ---');
const cc2=new __mock.MockSS('control-close'); SpreadsheetApp._active=cc2;
cc2.insertSheet('Closed_Shifts').getRange(1,1,2,3).setValues([
 ['Tour id','Closed by','Closed at'],
 ['2026-07-30|660|italian','Albert','2026-07-29 10:00']]);
const closedSet=readClosedShifts_();
check('closed shift id is read', closedSet['2026-07-30|660|italian']===true, closedSet);
check('shiftDomId_ regular = shiftKey_',
  shiftDomId_({dateKey:'2026-07-30',minutes:660,language:'Italian',private:false})==='2026-07-30|660|italian', null);
check('shiftDomId_ private = key|P<idx>',
  shiftDomId_({dateKey:'2026-07-30',minutes:660,language:'Italian',private:true,privIndex:2})==='2026-07-30|660|italian|P2', null);
check('an unclosed shift is not in the closed set', !closedSet['2026-07-30|660|french'], closedSet);
// New semantics: a closed shift with a booking must reappear.
const _closeFilter=(shifts,closed,bbk)=>shifts.filter(s=>{
  if(!closed[shiftDomId_(s)]) return true;
  const list=bbk[shiftKey_(s.dateKey,s.minutes,s.language)]||[];
  return list.some(b=>s.private?/privat/i.test(b.note||''):!/privat/i.test(b.note||''));
});
const _s1={dateKey:'2026-07-30',minutes:660,language:'Italian',private:false};
const _s2={dateKey:'2026-07-31',minutes:660,language:'Italian',private:false};
const _closed={'2026-07-30|660|italian':true,'2026-07-31|660|italian':true};
const _bbk={}; _bbk[shiftKey_('2026-07-31',660,'Italian')]=[{bookingId:'X',note:''}];
const _kept=_closeFilter([_s1,_s2],_closed,_bbk);
check('closed shift with NO booking stays hidden', !_kept.some(s=>s.dateKey==='2026-07-30'), _kept);
check('closed shift WITH a booking comes back', _kept.some(s=>s.dateKey==='2026-07-31'), _kept);

console.log('--- Portal timing report: percentiles + per-action stats ---');
check('percentile_ p50 (nearest-rank)', percentile_([10,20,30,40,50],50)===30, percentile_([10,20,30,40,50],50));
check('percentile_ p95', percentile_([10,20,30,40,50],95)===50, percentile_([10,20,30,40,50],95));
check('percentile_ empty -> 0', percentile_([],50)===0, null);
check('percentile_ single', percentile_([100],95)===100, null);
const _now='2026-08-02 10:00:00';
const _rows=[
 [_now,'assign',1000,'OK',''],
 [_now,'assign',9000,'OK',''],           // slow -> p95 high -> REVIEW
 [_now,'save',2000,'OK',''],
 [_now,'save',2500,'ERROR','boom'],      // an error -> REVIEW
 [_now,'move',1500,'OK','']];
const _stats=summarisePortalTimings_(_rows, PORTAL.SLOW_MS);
const _a=_stats.find(s=>s.action==='assign'), _s=_stats.find(s=>s.action==='save'), _m=_stats.find(s=>s.action==='move');
check('assign aggregated: count 2, max 9000, flagged REVIEW (p95>budget)', _a.count===2 && _a.max===9000 && _a.status==='REVIEW', _a);
check('save aggregated: 1 error -> REVIEW', _s.errors===1 && _s.status==='REVIEW', _s);
check('move aggregated: fast + no errors -> OK', _m.count===1 && _m.status==='OK', _m);
check('slow tours read logging threshold is configured', typeof PORTAL.SLOW_MS==='number' && PORTAL.SLOW_MS>0, PORTAL.SLOW_MS);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
