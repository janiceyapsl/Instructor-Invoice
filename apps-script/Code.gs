// ── SKILLS UNION INSTRUCTOR INVOICE — APPS SCRIPT ──────────────────────────
// Paste this entire file into your Apps Script editor and deploy as Web App
// (Execute as: Me, Who has access: Anyone)
// ────────────────────────────────────────────────────────────────────────────

const SECRET_TOKEN  = "SU2026PC!nv0!ce";          // Must match token in index.html
const NOTIFY_EMAILS = ["auto@skillsunion.com"];   // Every submission is emailed here
const ROSTER_SHEET  = "Roster";                   // Column A: instructor emails that get their own copy
const TIME_ZONE     = "Asia/Singapore";
const HEADERS = [
  'Submission Date', 'Invoice No', 'Name', 'Email', 'Phone', 'Invoice Date',
  'Programs (Total for FT & PT roles)', 'FT Roles & Hours', 'PT Roles & Hours',
  'Additional Services', 'Total (SGD)', 'Remarks',
];
const TEXT_FIELDS = [
  'invoiceNo', 'name', 'email', 'phone', 'date', 'programs', 'ftRoles', 'ptRoles',
  'services', 'grandTotal', 'bank', 'accNo', 'paynow', 'remarks',
];

// ── HANDLE POST ──────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const raw  = e.parameter.data || e.postData?.contents || '';
    const body = JSON.parse(raw);

    // ── Token validation ──────────────────────────────────────────────────
    if (body.token !== SECRET_TOKEN) {
      return jsonOut({ status: 'error', message: 'Submission rejected' });
    }

    // ── Honeypot: the hidden field is only ever filled in by bots ─────────
    // Reply "success" so the bot gets no signal, but save and send nothing.
    if (body.honeypot) {
      console.warn('Honeypot triggered; submission dropped');
      return jsonOut({ status: 'success' });
    }

    const data = cleanPayload(body);

    // ── Write to Google Sheet ─────────────────────────────────────────────
    const sheet = getMonthSheet();
    const now   = new Date().toLocaleString('en-SG', { timeZone: TIME_ZONE });

    sheet.appendRow([
      now,                    // A: Submission Date
      data.invoiceNo,         // B: Invoice No
      data.name,              // C: Name
      data.email,             // D: Email
      data.phone,             // E: Phone
      data.date,              // F: Invoice Date
      data.programs,          // G: Programs
      data.ftRoles,           // H: FT Roles
      data.ptRoles,           // I: PT Roles
      data.services,          // J: Services
      data.grandTotal,        // K: Total
      data.remarks,           // L: Remarks
    ].map(asPlainText));

    // ── Send email notification with PDF ─────────────────────────────────
    // Instructors on the Roster get the email directly, with the team in CC so
    // they can see it went to both. Anyone else: team only, flagged in the subject.
    // (The form is public, so we never email an address that isn't on the Roster.)
    try {
      const onRoster = isOnRoster(data.email);
      const subject  = `${onRoster ? '' : '[Not on Roster] '}Invoice Submitted — ${data.name || 'Instructor'} | ${data.grandTotal || ''} | ${data.invoiceNo ? 'Inv #' + data.invoiceNo : ''}`;
      const htmlBody = buildEmailHtml(data, onRoster);

      // Generate PDF from invoice HTML
      const pdfBlob = generatePdf(data);

      const emailOpts = {
        htmlBody: htmlBody,
        name:     'Skills Union Invoice System',
        noReply:  true,
        replyTo:  NOTIFY_EMAILS[0],   // replies from instructors reach the team
      };
      if (onRoster) {
        emailOpts.cc = NOTIFY_EMAILS.join(',');
      }
      if (pdfBlob) {
        emailOpts.attachments = [pdfBlob];
      }
      GmailApp.sendEmail(onRoster ? data.email : NOTIFY_EMAILS.join(','), subject, '', emailOpts);
    } catch (emailErr) {
      // Email failure should not block sheet save — log and continue
      console.error('Email send failed:', emailErr.message);
    }

    return jsonOut({ status: 'success' });

  } catch (err) {
    // Details go to the Apps Script log only, never back to the browser
    console.error('doPost error:', err.stack || err.message);
    return jsonOut({ status: 'error', message: 'Server error' });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── INPUT CLEANING ────────────────────────────────────────────────────────────
// The form already HTML-encodes values (& < > " '), which showed up in the Sheet
// as e.g. "I&#x27;ll". Decode back to plain text here; every output then
// escapes for its own context (asPlainText for the Sheet, esc for HTML).
function cleanPayload(body) {
  const data = {};
  TEXT_FIELDS.forEach(k => { data[k] = decodeEntities(body[k]).trim(); });
  return data;
}

function decodeEntities(v) {
  if (v === null || v === undefined) return '';
  return String(v)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#x27;|&#39;/g, "'")
    .replace(/&amp;/g, '&');
}

// Stop the Sheet treating submitted text as a formula (e.g. "=IMPORTXML(...)").
// A leading apostrophe forces plain text and isn't shown in the cell.
function asPlainText(v) {
  const s = String(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

// Escape text before putting it into the email or PDF HTML
function esc(v) {
  return String(v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ── ROSTER CHECK ─────────────────────────────────────────────────────────────
function isOnRoster(email) {
  const e = String(email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) return false;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROSTER_SHEET);
    if (!sheet || sheet.getLastRow() < 1) return false;
    return sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues()
      .some(r => String(r[0]).trim().toLowerCase() === e);
  } catch (err) {
    console.error('Roster lookup failed:', err.message);
    return false;
  }
}

// ── MONTHLY TAB ("Oct 2026") — created with headers if it doesn't exist ─────
function getMonthSheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = Utilities.formatDate(new Date(), TIME_ZONE, 'MMM yyyy');
  let sheet  = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name, 0);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── BUILD EMAIL HTML BODY ─────────────────────────────────────────────────────
function buildEmailHtml(data, onRoster) {
  const fmt = v => v ? esc(v) : '—';
  const team = NOTIFY_EMAILS.map(esc).join(', ');
  const intro = onRoster
    ? `This invoice has been submitted to Skills Union. It was sent to <strong>${esc(data.email)}</strong> and <strong>${team}</strong> at the same time. Details are below and the PDF is attached.`
    : `A new instructor invoice has been submitted. Details are below and the PDF is attached.`;
  const rosterNote = onRoster ? '' : `
      <div style="margin:0 0 16px;padding:12px 14px;background:#fff4e5;border-left:3px solid #c05500;border-radius:6px;font-size:13px;color:#7a3a00">
        <strong>Not on Roster:</strong> ${data.email ? esc(data.email) : 'no email given'} isn't on the Roster tab, so the instructor was <strong>not</strong> sent a copy.
        If this is a genuine instructor, add their email to the Roster tab and forward this email to them.
      </div>`;
  const rows = [
    ['Invoice No.',    fmt(data.invoiceNo)],
    ['Instructor',     fmt(data.name)],
    ['Email',          fmt(data.email)],
    ['Phone',          fmt(data.phone)],
    ['Invoice Date',   fmt(data.date)],
    ['Programs',       fmt(data.programs)],
    ['FT Roles',       fmt(data.ftRoles)],
    ['PT Roles',       fmt(data.ptRoles)],
    ['Services',       fmt(data.services)],
    ['Grand Total',    fmt(data.grandTotal)],
    ['Remarks',        fmt(data.remarks)],
  ].map(([label, value]) => `
    <tr>
      <td style="padding:7px 12px;font-size:13px;color:#666;white-space:nowrap;border-bottom:1px solid #eee;width:140px">${label}</td>
      <td style="padding:7px 12px;font-size:13px;color:#1a1a2e;border-bottom:1px solid #eee;white-space:pre-line">${value}</td>
    </tr>`).join('');

  return `
  <div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;border:1px solid #e8e8f0;border-radius:8px;overflow:hidden">
    <div style="background:#1a1a2e;padding:24px 28px">
      <div style="font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#8888aa;margin-bottom:6px">Skills Union</div>
      <div style="font-size:22px;font-weight:700;color:#fff">Invoice Submitted</div>
    </div>
    <div style="padding:20px 28px">
      ${rosterNote}
      <p style="font-size:14px;color:#4a4a6a;margin:0 0 16px">
        ${intro}
      </p>
      <table style="width:100%;border-collapse:collapse;border:1px solid #eee;border-radius:6px;overflow:hidden">
        ${rows}
      </table>
      <div style="margin-top:20px;padding:14px 16px;background:#f4f5f7;border-radius:6px;border-left:3px solid #3b5bdb">
        <span style="font-size:13px;color:#4a4a6a">Grand Total: </span>
        <span style="font-size:16px;font-weight:700;color:#3b5bdb">${fmt(data.grandTotal)}</span>
      </div>
    </div>
    <div style="padding:14px 28px;background:#f9f9fc;border-top:1px solid #eee;font-size:11px;color:#8888aa;text-align:center">
      Generated by the Skills Union Instructor Invoice System · ${new Date().getFullYear()}
    </div>
  </div>`;
}

// ── GENERATE PDF FROM INVOICE DATA ────────────────────────────────────────────
function generatePdf(data) {
  try {
    const fmt = v => v ? esc(v) : '—';
    const today = new Date().toLocaleDateString('en-SG', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const invDate = data.date ? new Date(data.date).toLocaleDateString('en-SG', { day: 'numeric', month: 'long', year: 'numeric' }) : '—';

    // ── Helper: parse "Description: X hrs = S$Amount" → {desc, amount}
    function parseRoleLine(line) {
      // Header line e.g. "SCTP AIE Cohort 2F Weekday Lead: 1 session = S$400"
      const amtMatch = line.match(/S\$([\.\d,]+)\s*$/);
      const amt = amtMatch ? 'S$ ' + amtMatch[1] : '';
      // Description = everything before " = S$..."
      const desc = line.replace(/\s*=\s*S\$[\.\d,]+\s*$/, '').trim();
      return { desc, amt };
    }

    // Build program rows — FT and PT roles
    function buildRoleRows(roleStr, sectionLabel) {
      if(!roleStr) return '';
      let html = '';
      const lines = roleStr.split('\n');
      let firstInSection = true;
      lines.forEach(line => {
        if(!line.trim()) return;
        if(line.startsWith('-')) {
          // Date sub-line — indent, no amount
          html += `<tr>
            <td style="padding:2px 12px 2px 24px;font-size:11px;color:#4a4a6a;border-bottom:1px solid #f5f5f5">${esc(line)}</td>
            <td style="border-bottom:1px solid #f5f5f5"></td>
          </tr>`;
        } else {
          // Header line — bold description + amount right-aligned
          const { desc, amt } = parseRoleLine(line);
          const topBorder = firstInSection ? 'border-top:2px solid #e8e8f0;' : 'border-top:1px solid #eee;';
          firstInSection = false;
          html += `<tr>
            <td style="padding:8px 12px 4px;font-size:12px;font-weight:700;color:#1a1a2e;${topBorder}">${esc(desc)}</td>
            <td style="padding:8px 12px 4px;font-size:12px;font-weight:700;color:#3b5bdb;text-align:right;white-space:nowrap;${topBorder}">${amt}</td>
          </tr>`;
        }
      });
      return html;
    }

        // Build service rows — each service on its own row with amount
    function buildServiceRows(servicesStr) {
      if(!servicesStr) return '';
      let html = '';
      // Services format: "Mentoring: 75hrs = S$75 | Ad-hoc: desc (S$50) | Program Lead: S$200 — desc"
      const parts = servicesStr.split(' | ');
      parts.forEach(part => {
        if(!part.trim()) return;
        // Try all amount patterns in order of specificity:
        // 1. "= S$400" (mentoring/training)
        // 2. "(S$50)" (ad-hoc in parentheses)
        // 3. "S$200" anywhere (program lead)
        let amt = '';
        let desc = part.trim();
        const eqMatch  = part.match(/=\s*S\$([\d,\.]+)/);
        const parMatch = part.match(/\(S\$([\d,\.]+)\)/);
        const anyMatch = part.match(/S\$([\d,\.]+)/);
        if(eqMatch) {
          amt = 'S$ ' + eqMatch[1];
          desc = part.replace(/\s*=\s*S\$[\d,\.]+.*$/, '').trim();
        } else if(parMatch) {
          amt = 'S$ ' + parMatch[1];
          desc = part.replace(/\s*\(S\$[\d,\.]+\).*$/, '').trim();
        } else if(anyMatch) {
          amt = 'S$ ' + anyMatch[1];
          desc = part.replace(/\s*S\$[\d,\.]+.*$/, '').trim();
        }
        html += `<tr>
          <td style="padding:7px 12px;font-size:12px;color:#1a1a2e;border-top:1px solid #eee;font-weight:600">${esc(desc)}</td>
          <td style="padding:7px 12px;font-size:12px;font-weight:700;color:#3b5bdb;text-align:right;white-space:nowrap;border-top:1px solid #eee">${amt}</td>
        </tr>`;
      });
      return html;
    }

    const programRows = (buildRoleRows(data.ftRoles, 'FT') + buildRoleRows(data.ptRoles, 'PT'))
      || `<tr><td colspan="2" style="padding:10px 12px;color:#888;font-size:12px">No programs recorded</td></tr>`;
    const serviceRows = buildServiceRows(data.services);

    const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:Arial,sans-serif; font-size:12px; color:#1a1a2e; background:#fff; }
  .header { padding:20px 24px 16px; border-bottom:2px solid #e8e8f0; }
  .header-top { display:flex; justify-content:space-between; align-items:flex-start; }
  .co { font-size:9px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:#888; margin-bottom:4px; }
  .title { font-size:20px; font-weight:700; color:#1a1a2e; }
  .inv-meta { text-align:right; }
  .inv-meta .lbl { font-size:9px; color:#888; margin-bottom:2px; }
  .inv-meta .val { font-size:13px; font-weight:700; }
  .body { padding:16px 24px; }
  .info-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:14px; padding-bottom:14px; border-bottom:1px solid #eee; }
  .info-item .lbl { font-size:9px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:#888; margin-bottom:2px; }
  .info-item .val { font-size:12px; font-weight:600; }
  .section-title { font-size:9px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:#888; margin-bottom:6px; }
  table { width:100%; border-collapse:collapse; margin-bottom:12px; }
  th { padding:5px 12px; font-size:9px; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:#888; background:#f4f5f7; text-align:left; }
  th:last-child { text-align:right; }
  .total-row { display:flex; justify-content:flex-end; margin-bottom:14px; }
  .total-box { width:220px; }
  .total-line { display:flex; justify-content:space-between; padding:8px 0; border-top:2px solid #1a1a2e; }
  .total-label { font-size:13px; font-weight:700; }
  .total-val { font-size:13px; font-weight:800; color:#3b5bdb; }
  .bottom-grid { border-top:1px solid #eee; padding-top:12px; }
  .footer { margin-top:12px; padding-top:8px; border-top:1px solid #eee; text-align:center; font-size:9px; color:#888; }
</style>
</head><body>
<div class="header">
  <div class="header-top">
    <div>
      <div class="co">Skills Union</div>
      <div class="title">Instructor Invoice</div>
    </div>
    <div class="inv-meta">
      <div class="lbl">Invoice No.</div>
      <div class="val">${fmt(data.invoiceNo)}</div>
      <div class="lbl" style="margin-top:6px">Invoice Date</div>
      <div class="val" style="font-size:12px">${invDate}</div>
      <div class="lbl" style="margin-top:4px">Submission Date</div>
      <div class="val" style="font-size:12px">${today}</div>
    </div>
  </div>
</div>
<div class="body">
  <div class="info-grid">
    <div class="info-item"><div class="lbl">Instructor</div><div class="val">${fmt(data.name)}</div></div>
    <div class="info-item"><div class="lbl">Email</div><div class="val">${fmt(data.email)}</div></div>
    <div class="info-item"><div class="lbl">Mobile</div><div class="val">${data.phone ? '+'+esc(data.phone) : '—'}</div></div>
    <div class="info-item"><div class="lbl">Payment</div><div class="val">${data.bank ? esc(data.bank) + (data.accNo ? ' · ' + esc(data.accNo) : '') : (data.paynow ? 'PayNow: ' + esc(data.paynow) : '—')}</div></div>
  </div>
  <div class="section-title">Programs &amp; Services</div>
  <table>
    <thead><tr><th>Description</th><th style="text-align:right">Amount</th></tr></thead>
    <tbody>
      ${programRows}
      ${serviceRows}
    </tbody>
  </table>
  <div class="total-row">
    <div class="total-box">
      <div class="total-line">
        <span class="total-label">Total Payable</span>
        <span class="total-val">${fmt(data.grandTotal)}</span>
      </div>
    </div>
  </div>
  <div class="bottom-grid">
    <div><div class="section-title">Remarks</div><div style="font-size:11px;color:#4a4a6a">${fmt(data.remarks)}</div></div>
  </div>
  <div class="footer">This invoice was generated by the Skills Union Instructor Invoice System · ${new Date().getFullYear()}</div>
</div>
</body></html>`;

    const blob = Utilities.newBlob(html, 'text/html', 'invoice.html');
    const pdf  = blob.getAs('application/pdf');
    const name = `Invoice_${(data.name||'Instructor').replace(/\s+/g,'_')}_${data.invoiceNo||today.replace(/\//g,'-')}.pdf`;
    pdf.setName(name);
    return pdf;

  } catch(err) {
    console.error('PDF generation failed:', err.message);
    return null;
  }
}

// ── HANDLE GET (health check) ─────────────────────────────────────────────────
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', service: 'SU Invoice System' }))
    .setMimeType(ContentService.MimeType.JSON);
}
