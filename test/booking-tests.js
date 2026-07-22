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

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
