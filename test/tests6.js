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

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
