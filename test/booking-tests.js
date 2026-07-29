/* ===== Italian + French support: booking side (routing, parsers, website availability) =====
   Bundled with mock.js + booking/bookingList_v2.gs + booking/websiteAvailabilityUpdate.gs. */
let pass=0, fail=0;
const check=(l,c,g)=>{ if(c){pass++;console.log('PASS  '+l);} else{fail++;console.log('FAIL  '+l+'  (got: '+JSON.stringify(g)+')');} };

console.log('--- Language routing (recognition, regression, unsupported, no fallback) ---');
check('IT: Italian recognized',            normalizeLanguage_('Italian')==='Italian', normalizeLanguage_('Italian'));
check('IT: Italiano recognized',           normalizeLanguage_('Italiano (Live tour guide)')==='Italian', null);
check('IT: German word Italienisch',       normalizeLanguage_('Italienisch')==='Italian', null);
check('IT: French word Italien',           normalizeLanguage_('Italien')==='Italian', null);
check('FR: French recognized',             normalizeLanguage_('French')==='French', null);
check('FR: Français recognized',           normalizeLanguage_('Français')==='French', null);
check('FR: Spanish word Francés',          normalizeLanguage_('Francés')==='French', null);
check('FR: Italian word Francese',         normalizeLanguage_('Francese')==='French', null);
check('FR: German word Französisch',       normalizeLanguage_('Französisch')==='French', null);
check('no fallback: Italian !== English',  normalizeLanguage_('Italian')!=='English', null);
check('no fallback: French !== English',   normalizeLanguage_('French')!=='English', null);
check('regression EN/DE/ES',               normalizeLanguage_('Inglés')==='English'&&normalizeLanguage_('Deutsch')==='German'&&normalizeLanguage_('Español')==='Spanish', null);
check('unsupported -> English (default)',   normalizeLanguage_('Klingon')==='English', normalizeLanguage_('Klingon'));
check('languageToSheet_ IT',               languageToSheet_('Italian')==='Italian Tours', languageToSheet_('Italian'));
check('languageToSheet_ FR',               languageToSheet_('French')==='French Tours', languageToSheet_('French'));
check('languageToSheet_ regression',       languageToSheet_('German')==='German Tours'&&languageToSheet_('Spanish')==='Spanish Tours', null);
check('sheetToLanguage_ IT',               sheetToLanguage_('Italian Tours')==='Italian', null);
check('sheetToLanguage_ FR',               sheetToLanguage_('French Tours')==='French', null);
check('activeSheetNames_ includes IT/FR',  activeSheetNames_().includes('Italian Tours')&&activeSheetNames_().includes('French Tours'), activeSheetNames_());

console.log('--- OTA parsers: Italian / French insertion, cancellation, dedup ---');
const gygIt=parseGygMessage_(makeFakeMsg_(RNR_FIXTURES_.gygItalian.subject, RNR_FIXTURES_.gygItalian.body),'confirm');
check('GYG Italian parsed',                !!gygIt, gygIt);
check('GYG Italian language Italian',      gygIt&&gygIt.language==='Italian', gygIt&&gygIt.language);
check('GYG Italian routes to Italian Tours', gygIt&&languageToSheet_(gygIt.language)==='Italian Tours', null);
check('GYG Italian 2 adults',              gygIt&&gygIt.guests===2, gygIt&&gygIt.guests);
const vFr=parseViatorMessage_(makeFakeMsg_(RNR_FIXTURES_.viatorFrench.subject, RNR_FIXTURES_.viatorFrench.body),'confirm');
check('Viator French parsed',              !!vFr, vFr);
check('Viator French language French',     vFr&&vFr.language==='French', vFr&&vFr.language);
check('Viator French routes to French Tours', vFr&&languageToSheet_(vFr.language)==='French Tours', null);
check('Viator French 3 adults',            vFr&&vFr.guests===3, vFr&&vFr.guests);
check('Italian cancel rejected in confirm mode', parseGygMessage_(makeFakeMsg_(RNR_FIXTURES_.gygItalianCancel.subject, RNR_FIXTURES_.gygItalianCancel.body),'confirm')===null, null);
const itCancel=parseGygMessage_(makeFakeMsg_(RNR_FIXTURES_.gygItalianCancel.subject, RNR_FIXTURES_.gygItalianCancel.body),'cancel');
check('Italian cancel parsed with isCancellation', !!itCancel&&itCancel.isCancellation===true, itCancel);
const dup=uniqueBookings_([gygIt, parseGygMessage_(makeFakeMsg_(RNR_FIXTURES_.gygItalian.subject, RNR_FIXTURES_.gygItalian.body),'confirm')]);
check('dedup: same Italian booking id -> 1 unique', dup.length===1, dup.length);

console.log('--- Website availability + capacity (Italian / French) ---');
const controlSS=new __mock.MockSS(WEBSITE_CONTROL_SPREADSHEET_ID);
__mock.SS_BY_ID[WEBSITE_CONTROL_SPREADSHEET_ID]=controlSS;
const ws=controlSS.insertSheet('Weekly_Schedule');
ws.getRange(1,1,3,6).setValues([
 ['Day','Time','Language','Guides needed','Active from','Active until'],
 ['Monday','11:00','Italian',1,'',''],
 ['Monday','17:00','French',1,'','']]);
const bss=new __mock.MockSS('booking'); SpreadsheetApp._active=bss;
bss.insertSheet('Italian Tours').getRange(1,1,2,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Marco','+39',5,'2026-09-07','11:00 AM','GetYourGuide',40,'GYGIT1','']]);
bss.insertSheet('French Tours').getRange(1,1,2,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['Pierre','+33',4,'2026-09-07','5:00 PM','Viator',70,'BRFR1','']]);
const avail=websiteBuildMonthAvailability_('2026-09');
const mon=avail['2026-09-07']||[];
const itSlot=mon.find(s=>s.language==='Italian');
const frSlot=mon.find(s=>s.language==='French');
check('Italian slot present on Mon 2026-09-07', !!itSlot, mon);
check('Italian capacity 20-5=15',           itSlot&&itSlot.spotsLeft===15, itSlot);
check('French slot present on Mon 2026-09-07', !!frSlot, mon);
check('French capacity 20-4=16',            frSlot&&frSlot.spotsLeft===16, frSlot);

console.log('--- Trigger guard: a failing run must never reach Apps Script ---');
SpreadsheetApp._active=new __mock.MockSS('booking-guard');
let threw=false;
try { safeTriggerRun_('unitTest', function(){ throw new Error('Service Spreadsheets timed out while accessing document'); }); }
catch(e){ threw=true; }
check('safeTriggerRun_ swallows a Spreadsheets timeout (no failure email)', threw===false, threw);
let ran=false;
safeTriggerRun_('unitTest2', function(){ ran=true; });
check('safeTriggerRun_ still runs the work normally', ran===true, ran);

console.log('--- Modification threads must not log a bogus parse failure ---');
const mkThread=(subj,body)=>({getMessages:()=>[makeFakeMsg_(subj,body)],getFirstMessageSubject:()=>subj});
check('"Booking detail change" thread is owned by the modify pass',
  threadIsModifyOrCancel_(mkThread('Booking detail change: - S779080 - GYGMX4FWXAZY','Fecha 20 de julio de 2026'))===true, null);
check('cancellation thread is owned elsewhere too',
  threadIsModifyOrCancel_(mkThread('Reserva cancelada - GYG123','Tu reserva ha sido cancelada.'))===true, null);
check('a plain confirmation is NOT treated as modify/cancel (still logged if it fails)',
  threadIsModifyOrCancel_(mkThread('Booking - S779080 - GYG996ZAK7R7','Se ha reservado tu producto'))===false, null);

console.log('--- Booking rows stay active until the tour DAY is over (portal keeps reservations) ---');
const _today=new Date(); _today.setHours(9,0,0,0);
const _yest=new Date(); _yest.setDate(_yest.getDate()-1); _yest.setHours(9,0,0,0);
const _mk=d=>({date:d,time:'9:00 AM',name:'Guest',bookingId:'BK1',language:'English',source:'Website',guests:2});
check('a tour EARLIER today is not yet moved to Done (reservations persist in portal)', tourDayIsOver_(_mk(_today))===false, null);
check('a tour YESTERDAY is done (row moves to Done)', tourDayIsOver_(_mk(_yest))===true, null);
check('start+2h isCompleted_ still true earlier today (Gmail/invariants unchanged)', isCompleted_(_mk(_today))===true, null);

console.log('--- Website alert is a parseable backup source (Processed gap fix) ---');
const webBody =
  'A new website reservation has been received.\n\n' +
  'Language: English\n' +
  'Name: Jane Doe\n' +
  'Email: jane@example.com\n' +
  'Phone: +34123456789\n' +
  'Guests: 3\n' +
  'Tour date: Sat, 5 Sep 2026\n' +
  'Time: 11:00 AM\n' +
  'Source: Website\n' +
  'Booking ID: RRABC1234\n' +
  'DateKey: 2026-09-05\n' +
  'Time24: 11:00\n\n' +
  'Message:\nSee you there';
const wb = parseWebsiteAlert_(makeFakeMsg_('NEW WEBSITE RESERVATION - Roots & Roads', webBody));
check('website alert parses into a booking',       !!wb, wb);
check('website alert booking id preserved',        wb && wb.bookingId==='RRABC1234', wb && wb.bookingId);
check('website alert language English',            wb && wb.language==='English', wb && wb.language);
check('website alert 3 guests',                    wb && wb.guests===3, wb && wb.guests);
check('website alert date from DateKey',           wb && dateKey_(wb.date)==='2026-09-05', wb && wb.date);
check('website alert routes to English Tours',     wb && languageToSheet_(wb.language)==='English Tours', null);
check('website alert is a valid booking',          wb && isValidBooking_(wb), wb);
check('a non-website email is not mis-parsed',     parseWebsiteAlert_(makeFakeMsg_('Booking - GYG', 'Se ha reservado tu producto'))===null, null);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
