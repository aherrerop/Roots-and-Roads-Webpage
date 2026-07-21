/* ===== Audit round: quota fix, queue layout+clear button, no emails ===== */
let pass=0, fail=0;
const check=(l,c,g)=>{ if(c){pass++;console.log('PASS  '+l);} else{fail++;console.log('FAIL  '+l+'  (got: '+JSON.stringify(g)+')');} };
const day=(o)=>{const d=new Date();d.setDate(d.getDate()+o);d.setHours(12,0,0,0);return d;};
const key=(d)=>Utilities.formatDate(d,null,'yyyy-MM-dd');

const control=new __mock.MockSS('control'); SpreadsheetApp._active=control;
const ledger=new __mock.MockSS('ledger'); __mock.SS_BY_ID['LEDGER']=ledger; __mock.PROPS['LEDGER_ID']='LEDGER';
__mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw']=new __mock.MockSS('booking');
const guides=control.insertSheet('Guides');
guides.getRange(1,1,3,10).setValues([
 ['Guide','Active?','Seniority','English','German','Spanish','French','Manager','Email','Password'],
 ['Albert',true,1,true,false,false,false,true,'a@x.com','pw'],
 ['Carlos',true,1,true,false,true,false,true,'c@x.com','pw']]);

console.log('--- Queue tab layout (button / headers / data) ---');
ensureQueueTabs_(ledger);
const gw=ledger.getSheetByName(QUEUE_TABS.GURUWALK);
check('row1 = CLEAR button label', /^CLEAR/.test(String(gw.getRange(1,1).getValue())), gw.getRange(1,1).getValue());
check('row1 col2 = unchecked box', gw.getRange(1,2).getValue()===false, gw.getRange(1,2).getValue());
check('row2 = headers', gw.getRange(2,1).getValue()==='Tour date', gw.getRange(2,1).getValue());
check('empty tab: lastQueueRow_ = header row', lastQueueRow_(gw)===QUEUE_HEADER_ROW, lastQueueRow_(gw));

console.log('--- Entries land at row 3+, no duplicates on rerun ---');
const D=key(day(-1));
const albert=ledger.insertSheet('Albert');
albert.getRange(1,1,1,16).setValues([LEDGER_HEADERS]);
albert.getRange(2,1,1,16).setValues([[D,'Mon','11:00 AM','English','Olga','+49','Guruwalk',3,0,3,0,18,13.3,'Free','BAR123','2026-07-20 12:00']]);
updateGuruwalkCheckinQueue_();
check('entry written at row 3', String(gw.getRange(3,4).getValue())==='BAR123', gw.getRange(3,1,1,5).getValues());
check('attendance = All', String(gw.getRange(3,8).getValue())==='All', gw.getRange(3,8).getValue());
const before=lastQueueRow_(gw);
updateGuruwalkCheckinQueue_(); updateGuruwalkCheckinQueue_();
check('two re-runs add nothing', lastQueueRow_(gw)===before, {before, after:lastQueueRow_(gw)});

console.log('--- CLEAR button ---');
gw.getRange(1,2).setValue(true);
handleLedgerEdit({range: gw.getRange(1,2)});
check('entries cleared', lastQueueRow_(gw)===QUEUE_HEADER_ROW, lastQueueRow_(gw));
check('headers survived', gw.getRange(2,1).getValue()==='Tour date', gw.getRange(2,1).getValue());
check('button survived', /^CLEAR/.test(String(gw.getRange(1,1).getValue())), gw.getRange(1,1).getValue());
check('checkbox auto-reset', gw.getRange(1,2).getValue()===false, gw.getRange(1,2).getValue());
check('status shows count cleared', /Cleared 1 entr/.test(String(gw.getRange(1,3).getValue())), gw.getRange(1,3).getValue());

console.log('--- Legacy layout upgrade preserves entries ---');
const vn=ledger.getSheetByName(QUEUE_TABS.VIATOR_NOSHOW);
vn.clear();
vn.getRange(1,1,1,NOSHOW_HEADERS.length).setValues([NOSHOW_HEADERS]);
vn.getRange(2,1,1,NOSHOW_HEADERS.length).setValues([[D,'5:00 PM','English','Viator','BR-1','Martha',1,0,'Albert','','Not checked in',false,'','']]);
ensureQueueTabs_(ledger);
check('headers moved to row 2', vn.getRange(2,1).getValue()==='Tour date', vn.getRange(2,1).getValue());
check('legacy entry preserved at row 3', String(vn.getRange(3,5).getValue())==='BR-1', vn.getRange(3,1,1,6).getValues());

console.log('--- Recurring notification emails removed ---');
check('sendGuruwalkCheckinReminder gone', typeof sendGuruwalkCheckinReminder === 'undefined', 'still defined');
check('dailyScheduleSelfTest gone', typeof dailyScheduleSelfTest === 'undefined', 'still defined');
check('archiveLedgerMonthly does not email', !/MailApp/.test(archiveLedgerMonthly.toString()), 'still emails');

console.log('=================================');
console.log('RESULT: '+pass+' passed, '+fail+' failed');
process.exit(fail?1:0);
