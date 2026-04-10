// ================================================================
//  AKI GATE CHECK-IN — Google Apps Script (3-scan flow)
//
//  DEPLOY:
//  1. script.google.com → New project → paste this file
//  2. Set NOTIFY_EMAIL and REPORT_TIME_HOUR below
//  3. Deploy → New Deployment → Web App
//     Execute as: Me  |  Access: Anyone
//  4. Copy Web App URL → paste into driver_checkin.html SCRIPT_URL
//  5. Run setupDailyTrigger() once from the editor to activate email
// ================================================================

const SHEET_NAME        = 'Gate Log';
const NOTIFY_EMAIL      = '';          // 'ops@aki.ae' or 'a@aki.ae,b@aki.ae'
const REPORT_TIME_HOUR  = 8;           // 8 = 8:00 AM GST daily report
const FLAG_HOURS        = 10;          // trips longer than this get flagged

// ── COLUMN LAYOUT ────────────────────────────────────────────────
const COLS = [
  'Trip ID',              // A
  'Plate',                // B
  'Driver',               // C
  'Vendor',               // D
  'Route',                // E
  'Department',           // F
  'City',                 // G
  'Entry Time (GST)',     // H  — Scan 1
  'Exit Time (GST)',      // I  — Scan 2
  'Return Time (GST)',    // J  — Scan 3
  'Out-for-Delivery Hrs', // K  exit → return
  'Total Gate Hrs',       // L  entry → return
  'Status',               // M  ENTRY / EXITED / RETURNED
  'Flagged'               // N
];

// ── POST HANDLER ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const d      = JSON.parse(e.postData.contents);
    const action = (d.action || '').toLowerCase();
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheet  = getOrCreate(ss);

    if (action === 'entry')  handleEntry(sheet, d);
    if (action === 'exit')   handleExit(sheet, d);
    if (action === 'return') handleReturn(sheet, d);

    refreshSummary(ss, sheet);
    return json({ status: 'ok', action });
  } catch(err) {
    return json({ status: 'error', message: err.message });
  }
}

// ── SCAN 1: ENTRY ────────────────────────────────────────────────
function handleEntry(sheet, d) {
  sheet.appendRow([
    d.tripId || '',
    d.plate  || '',
    d.driver || '',
    d.vendor || '',
    '', '', '',                         // route, dept, city — filled at scan 2
    d.entryTimestamp || '',
    '', '',                             // exit, return — filled later
    '', '',                             // durations
    'ENTRY',
    ''
  ]);
  colorRow(sheet, sheet.getLastRow(), '#E1F5EE');   // light teal — entered
}

// ── SCAN 2: EXIT ─────────────────────────────────────────────────
function handleExit(sheet, d) {
  const row = findRow(sheet, d.tripId, d.plate);
  if (row > 0) {
    sheet.getRange(row, 5).setValue(d.route  || '');   // Route
    sheet.getRange(row, 6).setValue(d.dept   || '');   // Dept
    sheet.getRange(row, 7).setValue(d.city   || '');   // City
    sheet.getRange(row, 9).setValue(d.exitTimestamp || '');  // Exit time
    sheet.getRange(row, 13).setValue('EXITED');
    colorRow(sheet, row, '#FAEEDA');   // amber — out for delivery
  } else {
    // Fallback: write new row with all available data
    sheet.appendRow([
      d.tripId || '', d.plate || '', d.driver || '', d.vendor || '',
      d.route || '', d.dept || '', d.city || '',
      '', d.exitTimestamp || '', '',
      '', '', 'EXITED', 'WARN — no entry row'
    ]);
    colorRow(sheet, sheet.getLastRow(), '#FAC775');
  }
}

// ── SCAN 3: RETURN ───────────────────────────────────────────────
function handleReturn(sheet, d) {
  const row = findRow(sheet, d.tripId, d.plate);
  const flagged = (d.totalDurationHrs || 0) > FLAG_HOURS
    ? `YES — ${d.totalDurationHrs}h total` : '';

  if (row > 0) {
    sheet.getRange(row, 10).setValue(d.returnTimestamp  || '');  // Return time
    sheet.getRange(row, 11).setValue(d.delivDurationHrs || '');  // Delivery hrs
    sheet.getRange(row, 12).setValue(d.totalDurationHrs || '');  // Total hrs
    sheet.getRange(row, 13).setValue('RETURNED');
    sheet.getRange(row, 14).setValue(flagged);
    colorRow(sheet, row, '#D5F0E4');   // green — returned
  } else {
    sheet.appendRow([
      d.tripId || '', d.plate || '', d.driver || '', d.vendor || '',
      d.route || '', d.dept || '', d.city || '',
      d.entryTimestamp || '', d.exitTimestamp || '', d.returnTimestamp || '',
      d.delivDurationHrs || '', d.totalDurationHrs || '',
      'RETURNED', flagged || 'WARN — no prior rows'
    ]);
    colorRow(sheet, sheet.getLastRow(), '#D5F0E4');
  }

  if (NOTIFY_EMAIL && flagged) {
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: `[AKI FLAG] Long trip: ${d.plate} — ${d.totalDuration}`,
        body: `Vehicle: ${d.plate}\nDriver: ${d.driver}\nVendor: ${d.vendor}\nRoute: ${d.route}\nTotal time: ${d.totalDuration}\nOut for delivery: ${d.delivDuration}`
      });
    } catch(e) {}
  }
}

// ── GET: CONTROL TOWER POLLING ───────────────────────────────────
function doGet(e) {
  const action = (e.parameter.action || 'today').toLowerCase();
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return json({ error: 'Sheet not found' });

  const raw  = sheet.getDataRange().getValues();
  const hdrs = raw[0];
  const all  = raw.slice(1).map(r => { const o={}; hdrs.forEach((h,i)=>o[h]=r[i]); return o; });

  const todayShort = Utilities.formatDate(new Date(), 'Asia/Dubai', 'dd MMM');

  if (action === 'out') {
    const res = all.filter(r => String(r['Status']).toUpperCase() === 'EXITED');
    return json({ count: res.length, trips: res });
  }
  if (action === 'flagged') {
    const res = all.filter(r => String(r['Flagged']).toUpperCase().startsWith('YES'));
    return json({ count: res.length, trips: res });
  }
  if (action === 'today') {
    const res = all.filter(r => String(r['Entry Time (GST)']).includes(todayShort));
    return json({ count: res.length, trips: res });
  }
  if (action === 'all') { const limit = parseInt(e.parameter.limit || '2000'); const res = all.slice(-limit).reverse(); return json({ count: res.length, trips: res }); }
  return json({ count: Math.min(all.length,100), trips: all.slice(-100).reverse() });
}

// ── SUMMARY TAB ──────────────────────────────────────────────────
function refreshSummary(ss, logSheet) {
  try {
    let s = ss.getSheetByName('Summary');
    if (!s) { s = ss.insertSheet('Summary'); ss.setActiveSheet(logSheet); }
    s.clearContents();

    const raw = logSheet.getDataRange().getValues();
    const all = raw.slice(1);
    const tod = Utilities.formatDate(new Date(), 'Asia/Dubai', 'dd MMM');
    const today = all.filter(r => String(r[7]).includes(tod));

    const byStatus = st => all.filter(r => String(r[12]).toUpperCase() === st.toUpperCase()).length;
    const todSt    = st => today.filter(r => String(r[12]).toUpperCase() === st.toUpperCase()).length;

    const durs = all.filter(r=>r[11]>0).map(r=>parseFloat(r[11])||0);
    const avg  = durs.length ? (durs.reduce((a,b)=>a+b,0)/durs.length).toFixed(1) : '—';

    s.appendRow(['AKI GATE LOG — LIVE SUMMARY']);
    s.appendRow(['Updated', Utilities.formatDate(new Date(),'Asia/Dubai','dd MMM yyyy HH:mm')]);
    s.appendRow([]);
    s.appendRow(['Today',                          Utilities.formatDate(new Date(),'Asia/Dubai','dd MMM yyyy')]);
    s.appendRow(['  Entries today',                today.length]);
    s.appendRow(['  Currently inside (ENTRY only)',todSt('ENTRY')]);
    s.appendRow(['  Out for delivery',             todSt('EXITED')]);
    s.appendRow(['  Returned',                     todSt('RETURNED')]);
    s.appendRow([]);
    s.appendRow(['All time']);
    s.appendRow(['  Total trips',                  all.length]);
    s.appendRow(['  Still inside (entry only)',    byStatus('ENTRY')]);
    s.appendRow(['  Out for delivery',             byStatus('EXITED')]);
    s.appendRow(['  Completed returns',            byStatus('RETURNED')]);
    s.appendRow(['  Avg total trip hrs',           avg]);

    s.getRange(1,1).setBackground('#07111F').setFontColor('#1D9E75').setFontWeight('bold');
  } catch(e) {}
}

// ── DAILY REPORT TRIGGER SETUP ───────────────────────────────────
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased().atHour(REPORT_TIME_HOUR).everyDays(1)
    .inTimezone('Asia/Dubai').create();
  Logger.log(`Trigger set: sendDailyReport at ${REPORT_TIME_HOUR}:00 GST daily`);
}

// ── DAILY REPORT EMAIL ───────────────────────────────────────────
function sendDailyReport() {
  if (!NOTIFY_EMAIL) return;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const raw  = sheet.getDataRange().getValues();
  const hdrs = raw[0];
  const all  = raw.slice(1).map(r => { const o={}; hdrs.forEach((h,i)=>o[h]=r[i]); return o; });

  const yest     = Utilities.formatDate(new Date(Date.now()-86400000), 'Asia/Dubai', 'dd MMM yyyy');
  const yShort   = Utilities.formatDate(new Date(Date.now()-86400000), 'Asia/Dubai', 'dd MMM');
  const dayRows  = all.filter(r => String(r['Entry Time (GST)']).includes(yShort));

  const entered  = dayRows.filter(r => String(r['Status']).toUpperCase() === 'ENTRY');
  const exited   = dayRows.filter(r => String(r['Status']).toUpperCase() === 'EXITED');
  const returned = dayRows.filter(r => String(r['Status']).toUpperCase() === 'RETURNED');
  const flagged  = dayRows.filter(r => String(r['Flagged']).toUpperCase().startsWith('YES'));

  const hc  = dayRows.filter(r => String(r['Department']).toUpperCase() === 'HC');
  const nhc = dayRows.filter(r => String(r['Department']).toUpperCase() === 'NHC');

  const durs = returned.map(r => parseFloat(r['Total Gate Hrs'])||0).filter(v=>v>0);
  const avg  = durs.length ? (durs.reduce((a,b)=>a+b,0)/durs.length).toFixed(1) : '—';
  const max  = durs.length ? Math.max(...durs).toFixed(1) : '—';
  const min  = durs.length ? Math.min(...durs).toFixed(1) : '—';

  const cityCount = {};
  dayRows.forEach(r => { const c=r['City']||'Unknown'; cityCount[c]=(cityCount[c]||0)+1; });

  const html = buildReport({ yest, dayRows, entered, exited, returned, flagged, hc, nhc, avg, max, min, cityCount });
  const csv  = buildCsv(hdrs, dayRows.map(r => hdrs.map(h => r[h]||'')));
  const blob = Utilities.newBlob(csv, 'text/csv', `AKI_Gate_${yShort.replace(' ','_')}.csv`);

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: `AKI Gate Report — ${yest} | ${dayRows.length} trips | ${exited.length} still out`,
    htmlBody: html,
    attachments: [blob]
  });
}

// ── HTML REPORT BUILDER ───────────────────────────────────────────
function buildReport(d) {
  const kpi = (val, lbl, color) => `
    <td width="25%" align="center" style="padding:16px 6px;border-right:1px solid #F2F5F8;">
      <div style="font-size:28px;font-weight:700;color:${color};">${val}</div>
      <div style="font-size:10px;color:#5A6A7A;margin-top:3px;text-transform:uppercase;letter-spacing:1px;">${lbl}</div>
    </td>`;

  const cityHtml = Object.entries(d.cityCount).sort((a,b)=>b[1]-a[1]).map(([c,n])=>`
    <tr><td style="padding:7px 12px;border-bottom:1px solid #F2F5F8;font-size:13px;">${c}</td>
    <td style="padding:7px 12px;border-bottom:1px solid #F2F5F8;font-size:13px;font-weight:700;text-align:right;">${n}</td></tr>`).join('');

  const tripHtml = d.returned.map(r => {
    const dept = String(r['Department']).toUpperCase();
    const dc = dept==='HC' ? '#0C447C' : '#993C1D';
    const db = dept==='HC' ? '#E6F1FB' : '#FAECE7';
    return `<tr>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;">${r['Plate']||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;">${r['Driver']||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;">${r['Vendor']||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;font-weight:600;">${r['Route']||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;text-align:center;">
        <span style="background:${db};color:${dc};padding:2px 7px;border-radius:10px;font-size:11px;font-weight:700;">${dept}</span></td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;">${r['City']||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:11px;color:#5A6A7A;">${String(r['Entry Time (GST)']).slice(11,19)||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:11px;color:#5A6A7A;">${String(r['Exit Time (GST)']).slice(11,19)||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:11px;color:#5A6A7A;">${String(r['Return Time (GST)']).slice(11,19)||'—'}</td>
      <td style="padding:7px 9px;border-bottom:1px solid #F2F5F8;font-size:12px;font-weight:600;text-align:right;">${r['Total Gate Hrs']?r['Total Gate Hrs']+'h':'—'}</td>
    </tr>`;}).join('');

  const stillOutHtml = d.exited.length ? `
    <tr><td colspan="10" style="padding:14px 20px;">
      <div style="background:#FAEEDA;border:1px solid #EF9F27;border-radius:8px;padding:12px 16px;">
        <div style="font-size:11px;font-weight:700;color:#633806;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px;">⚠ Vehicles Still Out — No Return Scanned</div>
        ${d.exited.map(r=>`<div style="font-size:12px;padding:3px 0;color:#633806;">${r['Plate']} — ${r['Driver']} — ${r['Vendor']} — Route ${r['Route']} — Exit: ${String(r['Exit Time (GST)']).slice(0,19)}</div>`).join('')}
      </div>
    </td></tr>` : '';

  const flagHtml = d.flagged.length ? `
    <tr><td colspan="10" style="padding:0 20px 14px;">
      <div style="background:#FCEBEB;border:1px solid #E24B4A;border-radius:8px;padding:12px 16px;">
        <div style="font-size:11px;font-weight:700;color:#791F1F;margin-bottom:6px;text-transform:uppercase;letter-spacing:1px;">🚩 Flagged</div>
        ${d.flagged.map(r=>`<div style="font-size:12px;padding:2px 0;color:#791F1F;">${r['Plate']} — ${r['Driver']} — Route ${r['Route']} — ${r['Flagged']}</div>`).join('')}
      </div>
    </td></tr>` : '';

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#F2F5F8;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#F2F5F8;padding:20px 0;">
<tr><td align="center"><table width="640" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;border:1px solid #CDD8E3;">

  <tr><td style="background:#07111F;padding:18px 24px;">
    <table width="100%"><tr>
      <td><span style="background:#1D9E75;color:#fff;font-weight:900;font-size:12px;padding:4px 10px;border-radius:6px;letter-spacing:1px;">AKI</span>
          <span style="color:#fff;font-size:16px;font-weight:700;margin-left:10px;vertical-align:middle;">Gate Tracking Report</span></td>
      <td align="right" style="color:#6A9EBE;font-size:12px;">${d.yest}</td>
    </tr></table>
  </td></tr>

  <tr><td><table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #CDD8E3;">
    <tr>
      ${kpi(d.dayRows.length,  'Total Entries', '#07111F')}
      ${kpi(d.returned.length, 'Returned',      '#0F6E56')}
      ${kpi(d.exited.length,   'Still Out',     d.exited.length>0?'#BA7517':'#07111F')}
      ${kpi(d.flagged.length,  'Flagged',       d.flagged.length>0?'#A32D2D':'#07111F')}
    </tr>
  </table></td></tr>

  <tr><td style="padding:18px 24px 0;">
    <table width="100%"><tr>
      <td width="32%" valign="top" style="padding-right:10px;">
        <div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#5A6A7A;margin-bottom:8px;">Department</div>
        <table width="100%" style="border:1px solid #CDD8E3;border-radius:8px;overflow:hidden;">
          <tr style="background:#E6F1FB;"><td style="padding:9px 12px;font-size:13px;font-weight:700;color:#0C447C;">HC</td><td style="padding:9px 12px;font-size:17px;font-weight:700;color:#0C447C;text-align:right;">${d.hc.length}</td></tr>
          <tr style="background:#FAECE7;"><td style="padding:9px 12px;font-size:13px;font-weight:700;color:#712B13;">NHC</td><td style="padding:9px 12px;font-size:17px;font-weight:700;color:#712B13;text-align:right;">${d.nhc.length}</td></tr>
        </table>
      </td>
      <td width="32%" valign="top" style="padding-right:10px;">
        <div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#5A6A7A;margin-bottom:8px;">Trip Duration (hrs)</div>
        <table width="100%" style="border:1px solid #CDD8E3;border-radius:8px;overflow:hidden;">
          <tr style="background:#F2F5F8;"><td style="padding:7px 12px;font-size:12px;color:#5A6A7A;">Avg</td><td style="padding:7px 12px;font-size:13px;font-weight:700;text-align:right;">${d.avg}</td></tr>
          <tr><td style="padding:7px 12px;font-size:12px;color:#5A6A7A;border-top:1px solid #F2F5F8;">Max</td><td style="padding:7px 12px;font-size:13px;font-weight:700;text-align:right;border-top:1px solid #F2F5F8;">${d.max}</td></tr>
          <tr style="background:#F2F5F8;"><td style="padding:7px 12px;font-size:12px;color:#5A6A7A;border-top:1px solid #CDD8E3;">Min</td><td style="padding:7px 12px;font-size:13px;font-weight:700;text-align:right;border-top:1px solid #CDD8E3;">${d.min}</td></tr>
        </table>
      </td>
      <td width="36%" valign="top">
        <div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#5A6A7A;margin-bottom:8px;">By City</div>
        <table width="100%" style="border:1px solid #CDD8E3;border-radius:8px;overflow:hidden;">
          ${cityHtml||'<tr><td style="padding:9px 12px;font-size:12px;color:#5A6A7A;">No data</td></tr>'}
        </table>
      </td>
    </tr></table>
  </td></tr>

  <tr><td style="padding:18px 24px 0;">
    <div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#5A6A7A;margin-bottom:10px;">Trip Detail — Completed Returns</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #CDD8E3;border-radius:8px;overflow:hidden;">
      <tr style="background:#07111F;">
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Plate</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Driver</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Vendor</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Route</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Dept</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">City</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Entry</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Exit</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:left;font-size:10px;">Return</th>
        <th style="padding:7px 9px;color:#6A9EBE;font-weight:600;text-align:right;font-size:10px;">Hrs</th>
      </tr>
      ${stillOutHtml}${flagHtml}
      ${tripHtml||'<tr><td colspan="10" style="padding:12px;font-size:12px;color:#5A6A7A;text-align:center;">No completed returns for this date.</td></tr>'}
    </table>
  </td></tr>

  <tr><td style="padding:16px 24px;border-top:1px solid #CDD8E3;margin-top:16px;">
    <table width="100%"><tr>
      <td style="font-size:11px;color:#5A6A7A;">CSV attached · Generated ${Utilities.formatDate(new Date(),'Asia/Dubai','dd MMM yyyy HH:mm')} GST</td>
      <td align="right" style="font-size:11px;color:#5A6A7A;">AKI Last Mile Operations</td>
    </tr></table>
  </td></tr>

</table></td></tr></table>
</body></html>`;
}

// ── HELPERS ───────────────────────────────────────────────────────
function findRow(sheet, tripId, plate) {
  const data = sheet.getDataRange().getValues();
  // Search by tripId first
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === tripId) return i+1;
  }
  // Fallback: most recent open/exited row for this plate
  for (let i = data.length-1; i >= 1; i--) {
    const st = String(data[i][12]).toUpperCase();
    if (String(data[i][1]).toUpperCase() === plate.toUpperCase() && (st==='ENTRY'||st==='EXITED')) return i+1;
  }
  return -1;
}

function colorRow(sheet, rowNum, color) {
  sheet.getRange(rowNum, 1, 1, COLS.length).setBackground(color);
}

function buildCsv(headers, rows) {
  const esc = v => { const s=String(v??''); return (s.includes(',')||s.includes('"')||s.includes('\n'))?`"${s.replace(/"/g,'""')}"`:s; };
  return [headers.map(esc).join(','), ...rows.map(r=>r.map(esc).join(','))].join('\r\n');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLS);
    const h = sheet.getRange(1, 1, 1, COLS.length);
    h.setBackground('#07111F'); h.setFontColor('#fff'); h.setFontWeight('bold'); h.setFontSize(10);
    sheet.setFrozenRows(1);
    [200,120,160,160,80,100,120,190,190,190,120,120,90,180].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  }
  return sheet;
}
