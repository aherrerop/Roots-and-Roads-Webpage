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

console.log('--- Second GetYourGuide account (destinationstewards) tagged source "GYG" ---');
// The 2nd GYG account auto-forwards here; the forwarded copy keeps the original
// "To:" (that 2nd address), so a message addressed to it is tagged "GYG" — same
// parser + pipeline, only the source name differs. No getTo -> "GetYourGuide".
const _gygMkTo=(to)=>{ const m=makeFakeMsg_(RNR_FIXTURES_.gygItalian.subject, RNR_FIXTURES_.gygItalian.body); m.getTo=()=>to; m.getCc=()=>''; return m; };
const gyg2=parseGygMessage_(_gygMkTo('DestinationStewards@gmail.com'),'confirm');   // case-insensitive
check('email to destinationstewards -> source "GYG"', gyg2 && gyg2.source==='GYG' && gyg2.source===RNR.SOURCE.GYG2, gyg2 && gyg2.source);
const gyg1=parseGygMessage_(_gygMkTo('rootsandroadstours@gmail.com'),'confirm');
check('email to the main inbox -> source "GetYourGuide"', gyg1 && gyg1.source==='GetYourGuide', gyg1 && gyg1.source);
check('a message with no readable recipient -> "GetYourGuide" (safe default)', gygIt.source==='GetYourGuide', gygIt.source);
check('"GYG" is a PAID source-model (we owe the guide)', RNR.MODEL[RNR.SOURCE.GYG2]==='paid', RNR.MODEL[RNR.SOURCE.GYG2]);

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

console.log('--- A confirmation sharing a thread with a modification is NOT lost ---');
// Mirrors the real case: order S779080 has a confirmation (GYGN6BZYAAHH) and a
// modification for a DIFFERENT item (GYGMX4FWXAZY) threaded together.
const _next=(o)=>{const d=new Date();d.setDate(d.getDate()+o);return (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear();};
const confBody=[
 '¡Hola! Buenas noticias.','Se ha reservado tu producto',
 'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town','Tour en inglés',
 'Número de referencia GYGN6BZYAAHH',
 'Fecha August 4, 2027 10:00 AM',
 'Número de participantes','5 x Adults (Edad 14 - 99)',
 'Cliente principal','Vijay Patel customer-pecme5uoyscbtib4@reply.getyourguide.com Teléfono: +447961286234 Idioma: English',
 'Idioma del tour','Inglés (Live tour guide)'
].join('\n');
const confMsg=makeFakeMsg_('Booking - S779080 - GYGN6BZYAAHH', confBody);
const modMsg =makeFakeMsg_('Booking detail change: - S779080 - GYGMX4FWXAZY','Fecha August 6, 2027 10:00 AM');
const mixedThread={getMessages:()=>[confMsg,modMsg],getFirstMessageSubject:()=>'Booking - S779080 - GYGN6BZYAAHH'};
const confParsed=uniqueBookings_(parseThread_(mixedThread, RNR.SOURCE.GYG, 'confirm'));
check('the confirmation still parses in confirm mode (modification does not swallow it)',
  confParsed.length===1 && confParsed[0].bookingId==='GYGN6BZYAAHH', confParsed.map(b=>b.bookingId));
check('the parsed confirmation is a VALID, registerable booking',
  confParsed.length===1 && isValidBooking_(confParsed[0]), confParsed[0]);
check('it is the right guest (Vijay Patel) and 5 adults',
  confParsed.length===1 && confParsed[0].name==='Vijay Patel' && confParsed[0].guests===5, confParsed[0]);
// The reconcile pass parses in 'any' mode (bypasses classification) — it must
// still surface the confirmation as a recoverable booking.
const anyValid=uniqueBookings_(parseThread_(mixedThread, RNR.SOURCE.GYG, 'any'))
  .filter(b=>isValidBooking_(b) && !b.isCancellation);
check('reconcile (any mode) surfaces the confirmation for recovery',
  anyValid.some(b=>b.bookingId==='GYGN6BZYAAHH'), anyValid.map(b=>b.bookingId));

console.log('--- GYG: the TOUR date (labelled Fecha) wins over an earlier stray date ---');
// An email carrying an earlier "booking made" date must not make an upcoming
// tour look completed (which is what marked it Processed but never listed it).
const strayBody=[
 'Reserva realizada el August 1, 2026',      // earlier, past date (booking made)
 'Número de referencia GYGSTRAY99',
 'Fecha August 4, 2027 10:00 AM',            // the real, future tour date
 'Número de participantes','3 x Adults (Edad 14 - 99)',
 'Cliente principal','Test Guest Teléfono: +34600000000 Idioma: English',
 'Idioma del tour','Inglés (Live tour guide)'
].join('\n');
const strayB=parseGygMessage_(makeFakeMsg_('Booking - S999 - GYGSTRAY99', strayBody),'confirm');
check('GYG picks the labelled Fecha (Aug 4 2027), not the earlier Aug 1 date',
  strayB && dateKey_(strayB.date)==='2027-08-04', strayB && (strayB.date&&dateKey_(strayB.date)));
check('so the upcoming tour is NOT wrongly treated as completed',
  strayB && isCompleted_(strayB)===false, strayB && isCompleted_(strayB));
check('it is still a valid, registerable booking', strayB && isValidBooking_(strayB), strayB);

console.log('--- activeBookingIdSet_ reads ids across the language tabs ---');
const idSS=new __mock.MockSS('booking-ids'); SpreadsheetApp._active=idSS;
const idEn=idSS.insertSheet('English Tours');
idEn.getRange(1,1,3,9).setValues([
 ['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes'],
 ['A','+34',2,'2027-08-04','11:00 AM','GetYourGuide',0,'GYGAAA111',''],
 ['B','+34',2,'2027-08-05','11:00 AM','Viator',0,'BR-222','']]);
const idset=activeBookingIdSet_();
check('active id set contains existing booking ids', idset['GYGAAA111']===true && idset['BR-222']===true, Object.keys(idset));
check('a missing id is correctly absent', !idset['GYGN6BZYAAHH'], Object.keys(idset));

console.log('--- Booking rows stay active until the tour DAY is over (portal keeps reservations) ---');
const _today=new Date(); _today.setHours(9,0,0,0);
const _yest=new Date(); _yest.setDate(_yest.getDate()-1); _yest.setHours(9,0,0,0);
const _mk=d=>({date:d,time:'9:00 AM',name:'Guest',bookingId:'BK1',language:'English',source:'Website',guests:2});
check('a tour EARLIER today is not yet moved to Done (reservations persist in portal)', tourDayIsOver_(_mk(_today))===false, null);
check('a tour YESTERDAY is done (row moves to Done)', tourDayIsOver_(_mk(_yest))===true, null);
// A tour that STARTED 3h ago: start+2h is an hour in the past -> completed.
// Built from wall-clock offset so the assertion never depends on the run time.
const _past=new Date(Date.now()-3*3600000);
let _ph=_past.getHours(); const _pap=_ph>=12?'PM':'AM'; let _p12=_ph%12; if(_p12===0)_p12=12;
const _pmin=_past.getMinutes();
const _ptime=_p12+':'+(_pmin<10?'0'+_pmin:_pmin)+' '+_pap;
check('start+2h isCompleted_ is true once a tour has run (Gmail/invariants unchanged)',
  isCompleted_({date:_past,time:_ptime,name:'Guest',bookingId:'BK1',language:'English',source:'Website',guests:2})===true, _ptime);

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

console.log('--- A manually-added past tour with NO Booking ID must still move to Done ---');
const mvSS=new __mock.MockSS('booking-move'); SpreadsheetApp._active=mvSS;
const mvEn=mvSS.insertSheet('English Tours');
mvEn.getRange(1,1,1,9).setValues([['Name','Phone','Number of Guests','Tour date','Time','Source','Income','Booking ID','Notes']]);
const _yd=new Date(); _yd.setDate(_yd.getDate()-1); _yd.setHours(12,0,0,0);
const _fd=new Date(); _fd.setDate(_fd.getDate()+3); _fd.setHours(12,0,0,0);
mvEn.getRange(2,1,2,9).setValues([
 ['Kleiton Reis','+34600',1,_yd,'4:00 PM','Guruwalk',6,'',''],        // manual, NO id, yesterday
 ['Future Guest','+34600',2,_fd,'11:00 AM','Website',0,'RRFUT1','']]); // upcoming, must stay
const _moved=moveCompletedBookingRowsToDone_();
check('the no-id past tour is recognised as completed', _moved.some(b=>b.name==='Kleiton Reis'), _moved.map(b=>b.name));
check('a synthetic Booking ID was assigned so Done/Log dedupe works',
  _moved.filter(b=>b.name==='Kleiton Reis').every(b=>/^MAN/.test(b.bookingId)), _moved);
const _act=mvEn.getDataRange().getDisplayValues().map(r=>r[0]);
check('the manual past row was removed from the active tab', _act.indexOf('Kleiton Reis')===-1, _act);
check('the upcoming tour stays active', _act.indexOf('Future Guest')!==-1, _act);

console.log('--- REAL emails (captured from the live inbox) parse correctly ---');
// Airbnb experience confirmation — exact fields from a real "booked your
// experience" email (guest first name only; the parser needs no more).
const airbnbReal = [
  'Esteban booked your experience for July 7',
  'Send them a message to confirm all the details.',
  'Esteban', 'Identity verified · 1 review', 'Bogotá, Colombia',
  'Barcelona Walking Tour: Gaudí, Modernism, & Gothic', 'Hosted by Albert',
  'Date and time', 'Tue, July 7, 2026 · 11:00 AM – 2:00 PM CET',
  'Guests', '2 adults',
  'Confirmation code', 'TAFPNM3A',
  '€15.00 x 2 guests', '€30.00', '-€6.00', '-€5.81', 'Total (EUR)', '€18.19'
].join('\n');
const ab = parseAirbnbMessage_(makeFakeMsg_('Fwd: Confirmed: Esteban booked your experience for July 7', airbnbReal), 'confirm');
check('Airbnb real: parsed + valid', !!ab && isValidBooking_(ab), ab);
check('Airbnb real: name Esteban',           ab && ab.name==='Esteban', ab && ab.name);
check('Airbnb real: date 2026-07-07',        ab && dateKey_(ab.date)==='2026-07-07', ab && (ab.date&&dateKey_(ab.date)));
check('Airbnb real: time 11:00 AM',          ab && ab.time==='11:00 AM', ab && ab.time);
check('Airbnb real: 2 adults',               ab && ab.guests===2, ab && ab.guests);
check('Airbnb real: net income €18.19',      ab && Math.abs(ab.income-18.19)<0.001, ab && ab.income);
check('Airbnb real: confirmation code',      ab && ab.bookingId==='TAFPNM3A', ab && ab.bookingId);

// Airbnb CANCELLATION — guest-initiated "<Name> canceled their reservation"
// (real forwarded email). It carries NO confirmation code, so it must be matched
// to its booking by name + date + time. Regression: this format was unrecognised
// (isCancel=false) so the cancelled tour was never removed from the list.
const airbnbCancelReal = [
  '---------- Forwarded message ---------',
  'From: Airbnb <automated@airbnb.com>',
  'Subject: Vladimir canceled their reservation: Barcelona Walking Tour:',
  'Vladimir canceled their reservation',
  'They’ll receive a full refund because they canceled at least 24 hours before the start time.',
  'Canceled reservation',
  'Barcelona Walking Tour: Gaudí, Modernism, & Gothic', 'Hosted by Albert',
  'Friday, August 28 11:00 AM · 2 guests'
].join('\n');
const abcSubj = 'Fwd: Vladimir canceled their reservation: Barcelona Walking Tour: Gaudí, Modernism, & Gothic';
check('Airbnb cancel: rejected in confirm mode', parseAirbnbMessage_(makeFakeMsg_(abcSubj, airbnbCancelReal),'confirm')===null, null);
const abc = parseAirbnbMessage_(makeFakeMsg_(abcSubj, airbnbCancelReal), 'cancel');
// (validity's future-date gate depends on "now"; the real email is future-dated
// and valid — here we assert the parse + fields + matching, which is what removal needs.)
check('Airbnb cancel: detected + parsed',     !!abc && abc.source===RNR.SOURCE.AIRBNB, abc);
check('Airbnb cancel: name Vladimir',         abc && abc.name==='Vladimir', abc && abc.name);
check('Airbnb cancel: time 11:00 AM',         abc && abc.time==='11:00 AM', abc && abc.time);
check('Airbnb cancel: date is Aug 28 (with the email\'s year, not JS default 2001)', abc && abc.date && abc.date.getMonth()===7 && abc.date.getDate()===28 && abc.date.getFullYear()>=2026, abc && abc.date);
check('Airbnb cancel: flagged isCancellation',abc && abc.isCancellation===true, abc && abc.isCancellation);
// It must MATCH the active Vladimir booking (no id/phone -> by name+date+time)…
const abActive = normalizeBooking_({ name:'Vladimir', phone:'', guests:2, date:abc&&abc.date, time:'11:00 AM', language:'English', source:RNR.SOURCE.AIRBNB, bookingId:'TA4XBEYM' });
check('Airbnb cancel: matches its booking',   abc && cancellationMatchesBooking_(abc, abActive)===true, null);
// …and must NOT sweep away a different guest sharing the slot.
check('Airbnb cancel: no false match',        abc && cancellationMatchesBooking_(abc, normalizeBooking_({ name:'Someone Else', phone:'', guests:2, date:abc&&abc.date, time:'11:00 AM', language:'English', source:RNR.SOURCE.AIRBNB, bookingId:'ZZZ' }))===false, null);
// htmlToText_ can split the date and time onto SEPARATE lines, and the forward
// carries its own "Date: Fri, Aug 21 … 11:27 AM" header — the parser must still
// read the TOUR date/time (Aug 28, 11:00 AM), never the header's (Aug 21, 11:27).
const airbnbCancelSplit = [
  '---------- Forwarded message ---------',
  'From: Airbnb <automated@airbnb.com>',
  'Date: Fri, Aug 21, 2026 at 11:27 AM',
  'Subject: Vladimir canceled their reservation: Barcelona Walking Tour: Gaudí, Modernism, & Gothic',
  'Vladimir canceled their reservation',
  'Barcelona Walking Tour: Gaudí, Modernism, & Gothic', 'Hosted by Albert',
  'Friday, August 28', '11:00 AM', '2 guests'
].join('\n');
const abc2 = parseAirbnbMessage_(makeFakeMsg_(abcSubj, airbnbCancelSplit), 'cancel');
check('Airbnb cancel (split lines): date is Aug 28', abc2 && abc2.date && abc2.date.getMonth()===7 && abc2.date.getDate()===28, abc2 && abc2.date);
check('Airbnb cancel (split lines): time 11:00 AM (not the 11:27 header)', abc2 && abc2.time==='11:00 AM', abc2 && abc2.time);
// PRODUCTION PATH: getBestMessageText_ prefers the HTML body (getBody), which the
// default makeFakeMsg_ leaves empty — so exercise the real HTML->text path. Airbnb
// renders "Friday, August 28\r\n11:00 AM · 2 guests" in a white-space:pre-line cell.
const abHtmlMsg = {
  getId:()=>'ABHTML1', getSubject:()=>abcSubj, getPlainBody:()=>'x',
  // Long enough (>200 chars of text) that pickBestBody_ chooses this HTML over the
  // plain body — the same choice production makes on the real (long) Airbnb email.
  getBody:()=>'<h1>Vladimir canceled their reservation</h1><div>They will receive a full refund because they canceled at least 24 hours before the start time. Show reservation. This booking has been removed from your calendar.</div><h2>Canceled reservation</h2><p>Barcelona Walking Tour: Gaudi, Modernism, and Gothic</p><p>Hosted by Albert</p><p style="white-space:pre-line">Friday, August 28\r\n11:00 AM · 2 guests</p>',
  getDate:()=>new Date('2026-08-21T11:15:00Z')
};
const abc3 = parseAirbnbMessage_(abHtmlMsg, 'cancel');
check('Airbnb cancel (HTML path): parsed', !!abc3 && abc3.isCancellation===true, abc3);
check('Airbnb cancel (HTML path): full date 2026-08-28', abc3 && abc3.date && abc3.date.getMonth()===7 && abc3.date.getDate()===28 && abc3.date.getFullYear()===2026, abc3 && abc3.date);
check('Airbnb cancel (HTML path): matches its 2026 booking',
  abc3 && cancellationMatchesBooking_(abc3, normalizeBooking_({ name:'Vladimir', phone:'', guests:2, date:new Date(2026,7,28), time:'11:00 AM', language:'English', source:RNR.SOURCE.AIRBNB, bookingId:'TA4XBEYM' }))===true, abc3 && abc3.date);
// removeActiveBooking_ must use the LENIENT (cancellation) matcher: a phone-less
// Airbnb cancellation vs the phone-less active Airbnb booking FAILS the strict
// sameBooking_ (it requires BOTH name AND phone) but matches cancellationMatchesBooking_.
// Using the strict matcher is exactly why the cancelled tour was never removed.
const abCancelNoPhone = normalizeBooking_({ name:'Vladimir', phone:'', guests:2, date:new Date(2026,7,28), time:'11:00 AM', language:'English', source:RNR.SOURCE.AIRBNB, bookingId:'' });
const abActiveNoPhone = normalizeBooking_({ name:'Vladimir', phone:'', guests:2, date:new Date(2026,7,28), time:'11:00 AM', language:'English', source:RNR.SOURCE.AIRBNB, bookingId:'TA4XBEYM' });
check('strict sameBooking_ does NOT match a phone-less cancellation (the old bug)', !sameBooking_(abActiveNoPhone, abCancelNoPhone), null);
check('lenient cancellationMatchesBooking_ DOES match it (the fix path)', cancellationMatchesBooking_(abCancelNoPhone, abActiveNoPhone)===true, null);
// THE ROOT BLOCKER: cancellationBookingsFromThread_ ran uniqueBookings_ →
// isValidBooking_, which REQUIRES a bookingId. A codeless Airbnb cancellation was
// dropped BEFORE it could reach removal, so the tour never left the list. This
// exercises the full thread path and asserts the codeless cancellation survives.
const abCancelThread = { getId:()=>'ABT1', getFirstMessageSubject:()=>abcSubj, getMessages:()=>[abHtmlMsg] };
const abCxs = cancellationBookingsFromThread_(abCancelThread, RNR.SOURCE.AIRBNB);
check('codeless Airbnb cancellation SURVIVES cancellationBookingsFromThread_', abCxs.length===1 && abCxs[0].name==='Vladimir' && !abCxs[0].bookingId, abCxs);
check('  …and carries the tour date+time so it can be matched', abCxs.length===1 && abCxs[0].date && abCxs[0].date.getDate()===28 && abCxs[0].time==='11:00 AM', abCxs[0]);

// GetYourGuide ITALIAN confirmation — exact fields from a real GYG email
// (as htmlToText_ renders it): tests language routing + the labelled-date fix.
const gygItalReal = [
  '¡Hola! Buenas noticias.', 'Se ha reservado tu producto',
  'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town', 'Tour en italiano',
  'Número de referencia', 'GYG7VKRVHNN2',
  'Fecha', 'December 24, 2026 5:00 PM',
  'Número de participantes', '1 x Child (Edad 0 - 13)', '3 x Adults (Edad 14 - 99)',
  'Cliente principal', 'Oujia Wu', 'customer-b34m5e3nb7n25iji@reply.getyourguide.com',
  'Teléfono: +393397905338', 'Idioma: Italian',
  'Idioma del tour', 'Italiano (Live tour guide)',
  'Precio', '61,00 €'
].join('\n');
const gg = parseGygMessage_(makeFakeMsg_('Booking - S779080 - GYG7VKRVHNN2', gygItalReal), 'confirm');
check('GYG real IT: parsed + valid',         !!gg && isValidBooking_(gg), gg);
check('GYG real IT: booking id',             gg && gg.bookingId==='GYG7VKRVHNN2', gg && gg.bookingId);
check('GYG real IT: name Oujia Wu',          gg && gg.name==='Oujia Wu', gg && gg.name);
check('GYG real IT: language Italian',       gg && gg.language==='Italian', gg && gg.language);
check('GYG real IT: language is NOT flagged uncertain', gg && gg.languageUncertain===false, gg && gg.languageUncertain);
check('GYG real IT: date 2026-12-24 (labelled Fecha)', gg && dateKey_(gg.date)==='2026-12-24', gg && (gg.date&&dateKey_(gg.date)));
check('GYG real IT: time 5:00 PM',           gg && gg.time==='5:00 PM', gg && gg.time);
check('GYG real IT: 3 adults',               gg && gg.guests===3, gg && gg.guests);
check('GYG real IT: routes to Italian Tours', gg && languageToSheet_(gg.language)==='Italian Tours', null);

// Viator confirmation — exact fields from a real "New Booking" email. The tour
// TIME comes from the tour grade ("English Tour 17:00"), not an explicit clock.
const viatorReal = [
  'No action is required. This booking is confirmed.', 'Booking Confirmation',
  'You have a new reservation for Barcelona Walking Tour: Sagrada Familia, Gaudi and Gothic Quarter.',
  'Booking Details',
  'Booking Reference: BR-1430762545',
  'Tour Name: Barcelona Walking Tour: Sagrada Familia, Gaudi and Gothic Quarter',
  'Travel Date: Fri, Sep 11, 2026',
  'Lead Traveler Name: Ruben Demen',
  'Traveler Names: Ruben Demen, Passenger Two',
  'Travelers: 2 Adults',
  'Product Code: 5631527P3', 'Tour Grade: English Tour 17:00', 'Tour Grade Code: TG4~17:00',
  'Tour Language: English - Guide', 'Net Rate: EUR €23,04',
  'Phone: (Alternate Phone)US+1 2817482866'
].join('\n');
const vi = parseViatorMessage_(makeFakeMsg_('New Booking for Fri, Sep 11, 2026 (#BR-1430762545)', viatorReal), 'confirm');
check('Viator real: parsed + valid',         !!vi && isValidBooking_(vi), vi);
check('Viator real: booking id',             vi && vi.bookingId==='BR-1430762545', vi && vi.bookingId);
check('Viator real: name Ruben Demen',       vi && vi.name==='Ruben Demen', vi && vi.name);
check('Viator real: date 2026-09-11',        vi && dateKey_(vi.date)==='2026-09-11', vi && (vi.date&&dateKey_(vi.date)));
check('Viator real: time 5:00 PM (from tour grade 17:00)', vi && vi.time==='5:00 PM', vi && vi.time);
check('Viator real: 2 adults',               vi && vi.guests===2, vi && vi.guests);
check('Viator real: net income €23.04',      vi && Math.abs(vi.income-23.04)<0.001, vi && vi.income);
check('Viator real: language English',       vi && vi.language==='English', vi && vi.language);

// Guruwalk confirmation — real "1 adult, 1 child" German booking. This caught a
// real bug: the child count used to be dropped (only the first number was read).
const guruReal = [
  'Hi Roots & roads,', 'You have a new booking on a Tour',
  'Congratulations, you have a confirmed booking on your guruwalk:',
  'Barcelona Highlights Free Tour: Sagrada Família, Passeig de Gràcia & Gothic Quarter',
  'Booking details:',
  'Walker: Yuliya', 'Booking code: BAR12477290', 'Phone +49 17661529180',
  'Attendees: 1 adult, 1 child', 'Language: German',
  'Date: Friday, 7 Aug 2026', 'Time: 10:00'
].join('\n');
const gw = parseGuruwalkMessage_(makeFakeMsg_('Confirmed booking 12477290 on your tour', guruReal), 'confirm')[0];
check('Guruwalk real: parsed + valid',        !!gw && isValidBooking_(gw), gw);
check('Guruwalk real: booking code',          gw && gw.bookingId==='BAR12477290', gw && gw.bookingId);
check('Guruwalk real: name Yuliya',           gw && gw.name==='Yuliya', gw && gw.name);
check('Guruwalk real: 1 adult',               gw && gw.guests===1, gw && gw.guests);
check('Guruwalk real: 1 CHILD (was dropped before)', gw && gw.children===1, gw && gw.children);
check('Guruwalk real: language German',       gw && gw.language==='German', gw && gw.language);
check('Guruwalk real: routes to German Tours', gw && languageToSheet_(gw.language)==='German Tours', null);
check('Guruwalk real: date 2026-08-07',       gw && dateKey_(gw.date)==='2026-08-07', gw && (gw.date&&dateKey_(gw.date)));
check('Guruwalk real: time 10:00 AM',         gw && gw.time==='10:00 AM', gw && gw.time);

// Guruwalk EVERY tour language. The operator email is ALWAYS English (lang="en");
// only the "Language:" VALUE changes. This exact body is a real confirmation
// (Emanuele Preti BAR12526571): day-first date "Sunday, 9 Aug 2026", a "Phone"
// label with NO colon, "Attendees: N adults". Proves the parse never depends on
// the tour language, and each maps to its own Tours tab.
console.log('--- Guruwalk parses every tour language (email is always English) ---');
const gwLangBody = lang => [
  'Hi Roots & roads,', 'You have a new booking on a Tour',
  'Congratulations, you have a confirmed booking on your guruwalk:',
  'Walker: Emanuele Preti', 'Booking code: BAR9526571', 'Phone +39 3402318218',
  'Attendees: 2 adults', 'Language: ' + lang, 'Date: Sunday, 9 Aug 2026', 'Time: 16:00'
].join('\n\n');
[['English','English Tours'],['German','German Tours'],['French','French Tours'],
 ['Italian','Italian Tours'],['Spanish','Spanish Tours']].forEach(function(pair){
  const lang=pair[0], sheet=pair[1];
  const g = parseGuruwalkMessage_(makeFakeMsg_('Confirmed booking 9526571 on your tour', gwLangBody(lang)), 'confirm')[0];
  check('Guruwalk '+lang+': language + routes to '+sheet, g && g.language===lang && languageToSheet_(g.language)===sheet, g && [g&&g.language, g&&languageToSheet_(g.language)]);
  check('Guruwalk '+lang+': id/name/2 adults/phone(no colon)/date/time all parse',
    g && g.bookingId==='BAR9526571' && g.name==='Emanuele Preti' && g.guests===2 &&
    /3402318218/.test(g.phone||'') && dateKey_(g.date)==='2026-08-09' && g.time==='4:00 PM', g);
});

console.log('--- REAL cancellation / modification emails (per source) ---');
// GYG cancellation (real format).
const gygCancelReal = ['GYGRFQLWK73H ha sido cancelada', 'Hola, proveedor:',
  'Te escribimos para informarte de que la siguiente reserva ha sido cancelada.',
  'Número de referencia: GYGRFQLWK73H', 'Tour: Tour definitivo por Barcelona'].join('\n');
const gyc = parseGygMessage_(makeFakeMsg_('Se ha cancelado una reserva - S779080 - GYGRFQLWK73H', gygCancelReal), 'cancel');
check('GYG cancel real: id + isCancellation', gyc && gyc.bookingId==='GYGRFQLWK73H' && gyc.isCancellation===true, gyc);
check('GYG cancel real: rejected in confirm mode', parseGygMessage_(makeFakeMsg_('Se ha cancelado una reserva - S779080 - GYGRFQLWK73H', gygCancelReal), 'confirm')===null, null);

// GYG MODIFICATION (real): date change 16 Aug -> 15 Aug; must pick the NEW date.
const gygModReal = ['Hola, Albert', 'Nos gustaría informarte de que la siguiente reserva se ha modificado.',
  'Italian Tour', 'Código de reserva', 'GYGBLHHH3G2Y',
  'Fecha Nuevo', '15 de agosto de 2026 a las 17:00', '16 de agosto de 2026 a las 17:00',
  'Número de participantes', '4', 'Idioma', 'Italiano'].join('\n');
const gym = parseGygMessage_(makeFakeMsg_('Booking detail change: - S779080 - GYGBLHHH3G2Y', gygModReal), 'modify');
check('GYG modify real: picks the NEW date 2026-08-15 (not the struck 08-16)', gym && dateKey_(gym.date)==='2026-08-15', gym && (gym.date&&dateKey_(gym.date)));
check('GYG modify real: 5:00 PM, 4 adults, Italian, not a cancellation',
  gym && gym.time==='5:00 PM' && gym.guests===4 && gym.language==='Italian' && gym.isCancellation===false, gym);

// Guruwalk CANCELLATION (real, label-then-newline format) — used to be dropped.
const gwCancelReal = ['SEBASTIAN HAS CANCELED A BOOKING FOR TUESDAY, JULY 28, 2026, 11:00H', 'Cancelled',
  'BOOKING DETAILS', 'Walker', 'Sebastian', 'Booking code', 'BAR12441706', 'Phone', '+49 17664031469',
  'Attendees', '1', 'Language', 'German', 'Date and time', 'Tuesday, July 28, 2026, 11:00h'].join('\n');
const gwc = parseGuruwalkMessage_(makeFakeMsg_('Sebastian has canceled booking BAR12441706 for the event on Tuesday, July 28, 2026, 11:00h', gwCancelReal), 'cancel')[0];
check('Guruwalk cancel real: parses the booking id (was DROPPED before)', gwc && gwc.bookingId==='BAR12441706', gwc && (gwc && gwc.bookingId));
check('Guruwalk cancel real: isCancellation + date + time', gwc && gwc.isCancellation===true && dateKey_(gwc.date)==='2026-07-28' && gwc.time==='11:00 AM', gwc);
check('Guruwalk cancel real: thread-id fallback also finds BAR id',
  extractBookingIdFromThread_({getFirstMessageSubject:()=>'Sebastian has canceled booking BAR12441706 for the event',getMessages:()=>[makeFakeMsg_('','')] }, RNR.SOURCE.GURUWALK)==='BAR12441706', null);

// Viator CANCELLATION (real): "Booking Canceled", #BR-…, Italian Tour 17:00.
const viCancelReal = ['Booking Canceled', 'Booking Details', 'Booking Reference: #BR-1429181493', 'Canceled',
  'Barcelona Walking Tour: Sagrada Familia, Gaudi and Gothic Quarter', 'Tour Option: Italian Tour 17:00',
  'Location: Barcelona, Spain', 'Travel Date: Sun, Aug 30, 2026', 'Travelers: 2 Adults',
  'Lead Traveler Name: VALERIA TORRE'].join('\n');
const vic = parseViatorMessage_(makeFakeMsg_('Cancelled Booking: Sun, Aug 30, 2026', viCancelReal), 'cancel');
check('Viator cancel real: id + isCancellation', vic && vic.bookingId==='BR-1429181493' && vic.isCancellation===true, vic);
check('Viator cancel real: rejected in confirm mode', parseViatorMessage_(makeFakeMsg_('Cancelled Booking: Sun, Aug 30, 2026', viCancelReal), 'confirm')===null, null);

console.log('--- Deduplication: different ids are DIFFERENT bookings (never merged) ---');
const _b1=normalizeBooking_({source:'GetYourGuide',bookingId:'GYGAAA1',name:'Jon Doe',phone:'+34600',date:'2027-08-04',time:'11:00 AM'});
const _b2=normalizeBooking_({source:'GetYourGuide',bookingId:'GYGBBB2',name:'Jon Doe',phone:'+34600',date:'2027-08-04',time:'11:00 AM'});
check('same guest + same slot but DIFFERENT ids -> NOT the same booking', sameBooking_(_b1,_b2)===false, null);
const _b3=normalizeBooking_({source:'GetYourGuide',bookingId:'GYGAAA1',name:'Jonathan Doe',phone:'+34600',date:'2027-08-04',time:'11:00 AM'});
check('same GYG id -> same booking even if the name text differs', sameBooking_(_b1,_b3)===true, null);
const _n1=normalizeBooking_({source:'Free Tour',name:'Ann Lee',phone:'+34611',date:'2027-08-04',time:'11:00 AM'});
const _n2=normalizeBooking_({source:'Free Tour',name:'Ann Lee',phone:'+34611',date:'2027-08-04',time:'11:00 AM'});
check('no ids: same name+phone+date+time -> same booking (soft match still works)', sameBooking_(_n1,_n2)===true, null);

console.log('--- Modification matching is lenient (id, or date+time+phone-OR-name) ---');
const _m0=normalizeBooking_({source:'GetYourGuide',name:'Anna Rossi',phone:'+34600111',date:'2027-08-15',time:'11:00 AM',guests:2});
const _mNoPhone=normalizeBooking_({source:'GetYourGuide',name:'Anna Rossi',phone:'',date:'2027-08-15',time:'11:00 AM',guests:3});
check('lenient: same date+time+name but NO phone -> match (updates, not duplicates)', sameBooking_(_m0,_mNoPhone,{lenient:true})===true, null);
check('strict (confirmations) still needs phone too -> no match', !sameBooking_(_m0,_mNoPhone), null);
const _mNoName=normalizeBooking_({source:'GetYourGuide',name:'A. Rossi typo',phone:'+34600111',date:'2027-08-15',time:'11:00 AM'});
check('lenient: same date+time+phone but tweaked name -> match', sameBooking_(_m0,_mNoName,{lenient:true})===true, null);
const _mIdA=normalizeBooking_({source:'GetYourGuide',bookingId:'GYGAAA1',name:'Anna Rossi',phone:'+34600111',date:'2027-08-15',time:'11:00 AM'});
const _mIdB=normalizeBooking_({source:'GetYourGuide',bookingId:'GYGBBB2',name:'Anna Rossi',phone:'+34600111',date:'2027-08-15',time:'11:00 AM'});
check('lenient STILL refuses two different platform ids (never merges paid bookings)', sameBooking_(_mIdA,_mIdB,{lenient:true})===false, null);
const _mOther=normalizeBooking_({source:'GetYourGuide',name:'Bob Other',phone:'+34600999',date:'2027-08-15',time:'11:00 AM'});
check('lenient does NOT match a different guest at the same slot', sameBooking_(_m0,_mOther,{lenient:true})===false, null);

console.log('--- Language recognition flag ---');
check('recognised: Inglés', languageRecognised_('Inglés (Live tour guide)')===true, null);
check('recognised: empty means platform-default, treated as fine', languageRecognised_('')===true, null);
check('NOT recognised: Portuguese', languageRecognised_('Portuguese')===false, null);
check('NOT recognised: gibberish', languageRecognised_('Xyzzy 123')===false, null);

console.log('--- Portal Feed rebuild (one read-optimised tab, preserves check-ins) ---');
(function(){
  const ss=new __mock.MockSS('booking'); SpreadsheetApp._active=ss;
  const mk=n=>{ const s=ss.insertSheet(n); s.getRange(1,1,1,RNR.ACTIVE_HEADERS.length).setValues([RNR.ACTIVE_HEADERS]);
    s.getRange(1,5,1,1).setNumberFormat('@'); s.getRange(1,8,1,1).setNumberFormat('@'); return s; };
  const en=mk('English Tours'), fr=mk('French Tours');
  const day=o=>{ const d=new Date(); d.setDate(d.getDate()+o); return d; };
  en.getRange(2,5,3,1).setNumberFormat('@'); en.getRange(2,8,3,1).setNumberFormat('@');
  fr.getRange(2,5,1,1).setNumberFormat('@'); fr.getRange(2,8,1,1).setNumberFormat('@');
  // in-window English (today+2) with a mixed note + a manager note in col J
  en.getRange(2,1,1,10).setValues([['Jane','+34600',3,day(2),'11:00 AM','GetYourGuide',30,'GYGFEED1','2 children; Aldo','VIP']]);
  // in-window French (today+3)
  fr.getRange(2,1,1,9).setValues([['Pierre','+33600',2,day(3),'10:00 AM','GetYourGuide',27,'GYGFEED2','']]);
  // out-of-window English (today+400) -> excluded
  en.getRange(3,1,1,9).setValues([['Old','+34',1,day(400),'11:00 AM','GetYourGuide',13.5,'GYGFAR','']]);
  // pre-existing check-in already in the feed for GYGFEED1 -> must be preserved
  const feed=ss.insertSheet('Portal Feed');
  feed.getRange(1,1,1,RNR.PORTAL_FEED_HEADERS.length).setValues([RNR.PORTAL_FEED_HEADERS]);
  feed.getRange(2,10,1,1).setNumberFormat('@'); feed.getRange(2,14,1,1).setNumberFormat('@');
  // 15-col seed: pre-existing check-in AND a pre-existing Guide (col O) for GYGFEED1.
  feed.getRange(2,1,1,15).setValues([['x','x','English','x','x',0,0,'x',0,'GYGFEED1','','',2,'10:05','Carlos']]);

  rebuildPortalFeed_();
  const fv=feed.getRange(2,1,feed.getLastRow()-1,16).getValues();
  const byId=id=>fv.find(r=>String(r[9])===id);
  check('feed has both in-window bookings only', fv.length===2, fv.length);
  check('feed excludes the out-of-window booking', !byId('GYGFAR'), null);
  const j=byId('GYGFEED1');
  check('feed row carries Language', j&&j[2]==='English', j&&j[2]);
  check('feed row carries children from the notes', j&&Number(j[6])===2, j&&j[6]);
  check('feed Notes strips the child count but keeps other tags (Aldo)', j&&j[10]==='Aldo', j&&j[10]);
  check('feed carries the Manager note (col L)', j&&j[11]==='VIP', j&&j[11]);
  check('rebuild PRESERVES an existing check-in by id (M,N)', j&&Number(j[12])===2&&String(j[13])==='10:05', j&&[j[12],j[13]]);
  check('rebuild PRESERVES the assigned Guide by id (O)', j&&String(j[14])==='Carlos', j&&j[14]);
  check('a reservation row is Type=booking (P)', j&&String(j[15])==='booking', j&&j[15]);
  const p=byId('GYGFEED2');
  check('a booking with no prior check-in has blank check-in cols', p&&(p[12]===''||p[12]==null), p&&p[12]);
  check('a booking with no prior guide has a blank Guide col', p&&(p[14]===''||p[14]==null), p&&p[14]);
})();

// REGRESSION (real incident): Viator names the customer's chosen language in the
// TOUR GRADE ("Italian Tour" / "Italian-language tour"), while a generic
// "Tour Language: English - Guide" line lies. The grade MUST win, or an Italian
// booking is sent to the English tour. Time still comes from the grade code.
const vItBody=[
  'Booking Details','Booking Reference: BR-1438955061',
  'Tour Name: Barcelona Walking Tour: Sagrada Familia, Gaudi and Gothic Quarter',
  'Travel Date: Sun, Aug 23, 2026','Lead Traveler Name: Emanuele Trotta',
  'Traveler Names: Emanuele Trotta, Emanuele Trotta','Travelers: 2 Adults',
  'Product Code: 5631527P3','Tour Grade: Italian Tour 16:00','Tour Grade Code: TG8~16:00',
  'Tour Grade Description: Italian-language tour','Tour Language: English - Guide'
].join('\n');
const vIt=parseViatorMessage_(makeFakeMsg_('Confirmed Booking: Sun, Aug 23, 2026', vItBody),'confirm');
check('Viator: "Tour Grade: Italian Tour" wins over "Tour Language: English"', vIt && vIt.language==='Italian', vIt && vIt.language);
check('Viator: time still from Tour Grade Code TG8~16:00 (stored 12h)', vIt && vIt.time==='4:00 PM', vIt && vIt.time);
check('Viator: no language in grade -> falls back to Tour Language', (function(){ const b=parseViatorMessage_(makeFakeMsg_('Confirmed Booking: Sun, Aug 23, 2026', 'Booking Reference: BR-1\nTravel Date: Sun, Aug 23, 2026\nLead Traveler Name: X\nTravelers: 1 Adult\nTour Grade Code: TG1~11:00\nTour Language: German - Guide'),'confirm'); return b && b.language==='German'; })(), null);

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
