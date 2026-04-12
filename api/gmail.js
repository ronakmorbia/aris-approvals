import { google } from 'googleapis';
import * as XLSX from 'xlsx';

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

function cleanTitle(subj) {
  return (subj || '')
    .replace(/^(Re:|RE:|Fwd:|FWD:)\s*/gi, '')
    .replace(/^APP-[A-Z]+-?\s*[-:]\s*/gi, '')
    .replace(/^APP-[A-Z]+\s+/gi, '')
    .replace(/^CRM-APP:\s*/gi, '')
    .replace(/^HR-APP:\s*/gi, '')
    .replace(/\s+[-–]\s+ASL\s+\d{2}-\d{2}-\d{4}/gi, '')
    .replace(/\s+[-–]\s+\d{2}-\d{2}-\d{4}/gi, '')
    .replace(/\s+of\s+Rs\.?\s*[\d,]+\s*\/-?/gi, '')
    .replace(/\s+/g, ' ').trim();
}

function fmtAmt(n) {
  if (!n || isNaN(n)) return '';
  if (n >= 10000000) return '₹' + (n / 10000000).toFixed(2) + ' Cr';
  if (n >= 100000) return '₹' + (n / 100000).toFixed(2) + ' L';
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

function extractAmtFromSubject(subj) {
  const s = subj || '';
  const rs = s.match(/Rs\.?\s*([\d,]+)\s*(?:\/[-]?)?/i);
  if (rs) { const n = parseInt(rs[1].replace(/,/g, '')); if (n > 0) return fmtAmt(n); }
  const cr = s.match(/(\d+\.?\d*)\s*Cr/i); if (cr) return '₹' + cr[1] + ' Cr';
  const l = s.match(/(\d+\.?\d*)\s*L(?:acs?|akhs?)?/i); if (l) return '₹' + l[1] + ' L';
  return '';
}

function detectType(subj) {
  const s = (subj || '').toUpperCase();
  if (s.includes('APP-VPAY') || s.includes('VENDOR PAYMENT')) return 'VPAY';
  if (s.includes('APP-TD') || s.includes('-TD-') || s.includes('DEPOSIT')) return 'TD';
  if (s.includes('APP-TRF') || s.includes('TRANSFER')) return 'TRF';
  if (s.includes('APP-EXP')) return 'EXP';
  if (s.includes('APP-FD') || s.includes('FIXED DEPOSIT')) return 'FD';
  if (s.includes('CRM-APP') || s.includes('CREDIT APPROVAL')) return 'CRM';
  if (s.includes('APP-HR') || s.includes('HR-APP')) return 'HR';
  return 'GEN';
}

function firstName(from) {
  const raw = (from || '').replace(/<[^>]+>/g, '').replace(/"/g, '').trim();
  // If it's a bare email address like trupti.gupta@aris.in
  if (raw.includes('@') && !raw.includes(' ')) {
    const local = raw.split('@')[0];
    const part = local.split('.')[0];
    return part.charAt(0).toUpperCase() + part.slice(1);
  }
  return raw.split(/\s+/)[0];
}

function decodeBody(data) {
  try { return Buffer.from((data || '').replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf-8'); }
  catch { return ''; }
}

function extractPlainText(payload) {
  if (!payload) return '';
  if (payload.mimeType === 'text/plain' && payload.body?.data) return decodeBody(payload.body.data);
  if (payload.parts) {
    for (const p of payload.parts) { const t = extractPlainText(p); if (t) return t; }
  }
  return '';
}

function stripQuotes(text) {
  const cuts = [/\r?\nOn .{10,120}wrote:\r?\n/i, /\r?\nFrom:.*\r?\nSent:/i];
  for (const re of cuts) { const i = text.search(re); if (i > 100) return text.slice(0, i); }
  return text;
}

function categorise(name, purpose, emailType) {
  const s = (name + ' ' + purpose).toLowerCase();
  if (/salary|payroll|wages|stipend|staff/i.test(s)) return 'Salary';
  if (/interest|ncd|bank charge|factoring|od |loan repay|finance|processing fee/i.test(s)) return 'Finance Cost';
  // Technology: only clearly identifiable tech tools/platforms/subscriptions
  if (/\baws\b|\bazure\b|google cloud|techmagify|sazs apps|\bslack\b|\bdropbox\b|\bfigma\b|\bnotion\b|\bcursor\b|openai|anthropic|atlassian|rebrandly|gupshup|relic\b|creativeit|\bsaas\b|corp cc|credit card.*tech|google workspace|microsoft 365|github|datadog|mixpanel|hubspot|salesforce|jira|zoom|webex/i.test(s)) return 'Technology';
  if (emailType === 'VPAY') return 'Payables';
  // EXP emails never have Payables — always Expenses
  return 'Expenses';
}

function parseExcel(base64Data, emailType) {
  try {
    const b64 = (base64Data || '').replace(/-/g, '+').replace(/_/g, '/');
    const buf = Buffer.from(b64, 'base64');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    let hIdx = -1, nameCol = -1, amtCol = -1, purposeCol = -1;
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i].map(c => String(c).toLowerCase().trim());
      const ni = row.findIndex(c => /vendor|party|name|payee|particular|expense head|beneficiary/i.test(c));
      const ai = row.findIndex(c => /^amount$|^amt$|payment amount|debit/i.test(c));
      const ai2 = ai >= 0 ? ai : row.findIndex(c => /amount|amt/i.test(c) && !c.includes('words'));
      if (ni >= 0 && ai2 >= 0) {
        hIdx = i; nameCol = ni; amtCol = ai2;
        purposeCol = row.findIndex(c => /purpose|remark|narration|description|brief|head/i.test(c) && c !== row[ni]);
        break;
      }
    }
    if (hIdx < 0) return [];

    const dataRows = [];
    for (let i = hIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = String(r[nameCol] || '').trim();
      const amtRaw = r[amtCol];
      const amt = typeof amtRaw === 'number' ? amtRaw : parseFloat(String(amtRaw).replace(/[₹,\s]/g, '')) || 0;
      if (name && amt > 0 && !/total|grand total|sub.?total/i.test(name)) {
        const purpose = purposeCol >= 0 ? String(r[purposeCol] || '').trim() : '';
        dataRows.push({ name, amount: amt, purpose, category: categorise(name, purpose, emailType) });
      }
    }
    dataRows.sort((a, b) => b.amount - a.amount);
    return dataRows.slice(0, 5).map(r => ({ ...r, amount: fmtAmt(r.amount) }));
  } catch (e) {
    console.error('Excel parse error:', e.message);
    return [];
  }
}

function parseExpenseRows(body) {
  const rows = [];
  const lines = body.split('\n');
  let section = null; // 'interest' or 'top5'

  for (const line of lines) {
    const l = line.trim();

    // Detect section headers
    if (/interest\s+expenses/i.test(l)) { section = 'interest'; continue; }
    if (/top\s+5\s+expenses/i.test(l)) { section = 'top5'; continue; }

    // Stop parsing after "Best," or signature
    if (/^best,|^regards|^thanks/i.test(l)) break;

    // Skip header rows
    if (!section) continue;
    if (/expense\s+head|amount\s*\(rs/i.test(l)) continue;

    // Parse data rows — format: "Name   amount   purpose"
    const m = l.match(/^(.+?)\s{2,}(\d[\d,]+)\s*(.*)?$/);
    if (m) {
      const name = m[1].trim();
      const amt = parseInt(m[2].replace(/,/g, ''));
      const purpose = (m[3] || '').trim();
      if (!name || amt <= 0) continue;

      // Force Finance Cost for interest section
      const cat = section === 'interest'
        ? 'Finance Cost'
        : categorise(name, purpose, 'EXP');

      rows.push({ name, amount: fmtAmt(amt), purpose, category: cat });
    }
  }

  // Return top 5 by amount across all sections
  rows.sort((a, b) => {
    const pa = parseFloat(a.amount.replace(/[₹,LCr\s]/g,'')) || 0;
    const pb = parseFloat(b.amount.replace(/[₹,LCr\s]/g,'')) || 0;
    return pb - pa;
  });
  return rows.slice(0, 7); // allow up to 7 to show interest + top5
}

function extractCrmRisk(body) {
  const m = (body || '').match(/Recommendation from Credit Team\s*:\s*(High|Medium|Low)/i);
  return m ? m[1] : '';
}

function extractCrmFields(body) {
  const fields = [];
  const patterns = [
    { label: 'Customer', re: /1\.\s+Customer Name\s*:\s*(.+)/i },
    { label: 'Entity', re: /2\.\s+Supplying Entity\s*:\s*(.+)/i },
    { label: 'Credit Limit', re: /6\.\s+Credit Limit.*?:\s*(.+)/i },
    { label: 'Credit Period', re: /8\.\s*Credit Period.*?:\s*(.+)/i },
    { label: 'Location', re: /12\.\s+City\s*:\s*(.+)/i },
    { label: 'Manager', re: /13\.\s+Assigned Manager\s*:\s*(.+)/i },
    { label: 'Risk', re: /18\.\s+Recommendation.*?:\s*(.+)/i },
  ];
  for (const { label, re } of patterns) {
    const m = body.match(re);
    if (m) fields.push({ label, value: m[1].trim().replace(/\r/g, '') });
  }
  return fields;
}

function parseTDAmounts(body) {
  const rows = [];
  let total = 0;
  const lines = body.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  // Try multiline: date on one line, amount on next
  for (let i = 0; i < lines.length; i++) {
    const dateMatch = lines[i].match(/^(\d{2}-\d{2}-\d{4})$/);
    if (dateMatch) {
      let amtStr = '';
      // Check same line first
      const sameLine = lines[i].match(/(\d{2}-\d{2}-\d{4})\s+([\d,]+)/);
      if (sameLine) {
        amtStr = sameLine[2];
      } else if (i + 1 < lines.length && /^[\d,]+$/.test(lines[i + 1])) {
        amtStr = lines[i + 1];
      }
      if (amtStr) {
        const amt = parseInt(amtStr.replace(/,/g, ''));
        if (amt > 0) { total += amt; rows.push({ date: dateMatch[1], amount: amt }); }
      }
    }
  }

  // Fallback: same line pattern
  if (!rows.length) {
    for (const line of lines) {
      const m = line.match(/(\d{2}-\d{2}-\d{4})\s+([\d,]+)/);
      if (m) {
        const amt = parseInt(m[2].replace(/,/g, ''));
        if (amt > 0) { total += amt; rows.push({ date: m[1], amount: amt }); }
      }
    }
  }

  return { rows, total };
}



// Shorten company names for display
function shortName(s) {
  if (!s) return s;
  const map = {
    'arisinfra solutions limited': 'ARIS',
    'arisinfra solutions private limited': 'ARIS',
    'arisinfra solutions': 'ARIS',
    'arisinfra': 'ARIS',
    'buildmex infra private limited': 'BM',
    'buildmex infra': 'BM',
    'buildmex': 'BM',
    'natureresidences real estate development private limited': 'NRDPL',
    'natureresidences real estate development': 'NRDPL',
    'natureresidences realtors private limited': 'NRPL',
    'natureresidences realtors': 'NRPL',
    'arisinfra constructions maharashtra private limited': 'ACMPL',
    'whiteroots infra private limited': 'WR',
    'whiteroots': 'WR',
    'chennai mines private limited': 'CMPL',
    'chennai mines': 'CM',
    'ps blue metals': 'PS Blue',
    'p.s. blue metals': 'PS Blue',
    'apar infra solutions private limited': 'Apar Infra',
    'apar infra solutions': 'Apar Infra',
    'sun-x concrete india private limited': 'Sun-X',
    'netwin roadways': 'Netwin',
    'sarvadnya enterprises': 'Sarvadnya',
    'satyam ventures projects private limited': 'Satyam Ventures',
    'gvee infra pvt ltd': 'GVEE Infra',
  };
  const key = s.trim().toLowerCase();
  if (map[key]) return map[key];
  // Strip legal suffixes
  return s.replace(/\s+(private limited|pvt\.?\s*ltd\.?|limited|llp|solutions private limited)$/i, '').trim();
}

// Generate smart human-readable title from email data
function smartTitle(type, subj, from, body, rows) {
  const d = subj.match(/(\d{2}-\d{2}-\d{4})/);
  const dateStr = d ? (() => {
    const [dd, mm, yyyy] = d[1].split('-');
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return dd + ' ' + months[parseInt(mm)-1];
  })() : '';

  if (type === 'VPAY') {
    const acct = subj.includes('OD & CA') ? 'OD & CA' : subj.includes('OD') ? 'OD' : 'CA';
    return `Vendor Payments — ${acct}${dateStr ? ' · ' + dateStr : ''}`;
  }
  if (type === 'EXP') {
    return `Other Expenses — ASL${dateStr ? ' · ' + dateStr : ''}`;
  }
  if (type === 'TD') {
    const rawParty = subj
      .replace(/^APP-TD[-–\s:]*/i, '')
      .replace(/^\d+\.?\d*\s*Cr?\s*[-–]?\s*Deposit[-–\s]*/i, '')
      .replace(/[-–]?\s*(ASL|BIPL)\s*$/i, '')
      .replace(/Deposit[-–\s]*/i, '')
      .trim();
    const party = shortName(rawParty) || rawParty.split(' ').slice(0,4).join(' ');
    return `Trade Deposit — ${party || 'See email'}`;
  }
  if (type === 'TRF') {
    const fromMatch = (body||'').match(/From:\s*(.+?)(?:\r?\n|To:)/i);
    const toMatch = (body||'').match(/To:\s*(.+?)(?:\r?\n|Amount:)/i);
    if (fromMatch && toMatch) return `Transfer — ${fromMatch[1].trim().split(' ').slice(0,2).join(' ')} → ${toMatch[1].trim().split(' ').slice(0,2).join(' ')}`;
    return `Internal Transfer${dateStr ? ' · ' + dateStr : ''}`;
  }
  if (type === 'CRM') {
    const party = subj.replace(/^.*Credit Approval[-–\s]*/i,'').replace(/[-–]\s*\(.*\)$/,'').trim();
    return `Credit Approval — ${party || 'Customer'}`;
  }
  if (type === 'HR') {
    return subj.replace(/^(Re:|RE:|APP-HR[-–]?|HR-APP:)\s*/gi,'').trim().slice(0, 50);
  }
  if (type === 'FD') {
    return subj.replace(/^APP-FD:?\s*/i,'').replace(/Rs\s+/i,'').trim().slice(0, 50);
  }
  return subj.replace(/^(Re:|RE:|APP-[A-Z]+-?\s*[-:]?\s*)/gi,'').trim().slice(0, 60);
}


// Extract the expense date from subject line (DD-MM-YYYY format)
// Returns true if the email belongs to the current month
function isCurrentMonth(subj, emailDate) {
  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();

  // Try subject date first (most accurate — e.g. "ASL 03-04-2026")
  const subjMatch = (subj || '').match(/(\d{2})-(\d{2})-(\d{4})/);
  if (subjMatch) {
    const month = parseInt(subjMatch[2]) - 1;
    const year  = parseInt(subjMatch[3]);
    return month === curMonth && year === curYear;
  }

  // Fall back to email date
  const d = new Date(emailDate);
  return d.getMonth() === curMonth && d.getFullYear() === curYear;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action } = req.query;

  if (action === 'auth-url') {
    const url = oauth2Client.generateAuthUrl({
      access_type: 'offline', prompt: 'consent',
      scope: ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify']
    });
    return res.json({ url });
  }

  if (action === 'auth-callback') {
    try {
      const { tokens } = await oauth2Client.getToken(req.query.code);
      res.setHeader('Location', `/?tokens=${encodeURIComponent(JSON.stringify(tokens))}`);
      return res.status(302).end();
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  const { tokens } = req.method === 'POST' ? req.body : {};
  if (!tokens) return res.status(400).json({ error: 'No tokens' });
  oauth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

  if (action === 'inbox') {
    try {
      const [fundsRes, crmRes, hrRes, taxRes] = await Promise.all([
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds in:inbox', maxResults: 50 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:CRM in:inbox', maxResults: 30 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:HR in:inbox', maxResults: 20 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:Tax in:inbox', maxResults: 20 }),
      ]);

      const all = [
        ...(fundsRes.data.messages || []).map(m => ({ ...m, cat: 'Funds' })),
        ...(crmRes.data.messages || []).map(m => ({ ...m, cat: 'CRM' })),
        ...(hrRes.data.messages || []).map(m => ({ ...m, cat: 'HR' })),
        ...(taxRes.data.messages || []).map(m => ({ ...m, cat: 'Tax' })),
      ];

      const seen = new Map();
      for (const m of all) { if (!seen.has(m.threadId)) seen.set(m.threadId, m); }
      const msgList = [...seen.values()];

      // Step 1: metadata for all in parallel
      const metaList = await Promise.all(msgList.map(async msg => {
        try {
          const d = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'metadata', metadataHeaders: ['From', 'To', 'Cc', 'Subject', 'Date'] });
          const h = {};
          for (const hdr of (d.data.payload?.headers || [])) h[hdr.name.toLowerCase()] = hdr.value;
          const labelIds = d.data.labelIds || [];
          const isUnread = labelIds.includes('UNREAD');
          const to = (h.to || '').toLowerCase();
          const isCc = !to.includes('ronak@aris.in') && !to.includes('ronak@arisinfra.one');
          const status = !isUnread ? 'done' : isCc ? 'fyi' : 'pending';
          const type = detectType(h.subject || '');
          return { msg, h, status, isUnread, isCc, type, snippet: d.data.snippet || '' };
        } catch { return null; }
      }));

      const valid = metaList.filter(Boolean);
      // EXP done items need full thread processing to get Rohan's category breakdown for KPIs
      // All other done items just get metadata (fast path)
      const pending = valid.filter(m => m.status !== 'done');
      const expDone = valid.filter(m => m.status === 'done' && detectType(m.h.subject||'') === 'EXP');
      const tdDone = valid.filter(m => m.status === 'done' && detectType(m.h.subject||'') === 'TD');
      const done = valid.filter(m => m.status === 'done' && detectType(m.h.subject||'') !== 'EXP' && detectType(m.h.subject||'') !== 'TD');

      const processMsg = async (meta) => {
        const { msg, h, status, isUnread, isCc, type, snippet } = meta;
        const subj = h.subject || '';
        let amount = extractAmtFromSubject(subj);
        let smartTitleStr = '';
        let fields = [];
        let rows = [];
        let note = '';
        let risk = '';
        let rohanApproved = false;
        let approvalPill = '';

        try {
          if (type === 'EXP') {
            const thread = await gmail.users.threads.get({ userId: 'me', id: msg.threadId, format: 'full' });
            const messages = thread.data.messages || [];

            // Check latest message for cancellation
            const latestMsg = messages[messages.length - 1];
            const latestBody = extractPlainText(latestMsg?.payload || {});
            const isCancelled = /please ignore|kindly ignore|disregard/i.test(latestBody);

            if (isCancelled) {
              return {
                tid: msg.threadId, mid: msg.id, cat: msg.cat, subj, type,
                title: smartTitle(type, subj, h.from, '', []),
                amount: extractAmtFromSubject(subj),
                risk: '', rows: [],
                note: 'Sandesh has asked to ignore this request. Please disregard.',
                fields: [{ label: 'Status', value: 'Cancelled — Please Ignore' }, { label: 'From', value: firstName(h.from) }],
                from: h.from || '', date: h.date || '', to: h.to || '', cc: h.cc || '',
                isUnread, isCc, status, rohanApproved: false, approvalPill: ''
              };
            }

            // Find Rohan's reply for expense breakdown
            const rohanMsg = messages.find(m =>
              (m.payload?.headers || []).find(hh => hh.name === 'From' && hh.value.toLowerCase().includes('rohan'))
            );
            if (rohanMsg) {
              const rohanBody = extractPlainText(rohanMsg.payload);
              rows = parseExpenseRows(rohanBody);
              rohanApproved = true;
              note = `${rows.length} expense items pre-approved by Rohan. Review breakdown and approve.`;
            } else {
              note = 'Pre-approved by Rohan. Full list in attached Excel. Awaiting your sign-off.';
              rohanApproved = true;
            }

            // Total from Sandesh's original message
            const sandeshMsg = messages.find(m =>
              (m.payload?.headers || []).find(hh => hh.name === 'From' && hh.value.toLowerCase().includes('sandesh'))
            );
            if (sandeshMsg) {
              const sb = extractPlainText(sandeshMsg.payload);
              const tm = sb.match(/Total Amount\s*[:\s]+([\d,]+)\s*\/-?/i);
              if (tm) amount = fmtAmt(parseInt(tm[1].replace(/,/g, '')));
            }

            fields = [
              { label: 'Company', value: 'Arisinfra Solutions Limited' },
              { label: 'Pre-Approved By', value: 'Rohan' },
              { label: 'Total Amount', value: amount },
            ];

          } else if (type === 'VPAY') {
            const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
            const body = extractPlainText(full.data.payload);

            const findAtts = (parts) => {
              const atts = [];
              if (!parts) return atts;
              for (const p of parts) {
                if (p.filename && p.body?.attachmentId && (p.mimeType?.includes('spreadsheet') || p.filename.endsWith('.xlsx')))
                  atts.push({ name: p.filename, id: p.body.attachmentId });
                if (p.parts) atts.push(...findAtts(p.parts));
              }
              return atts;
            };
            const atts = findAtts([full.data.payload]);

            if (atts.length > 0) {
              const attRes = await gmail.users.messages.attachments.get({ userId: 'me', messageId: msg.id, id: atts[0].id });
              rows = parseExcel(attRes.data.data, 'VPAY');
            }

            const tm = body.match(/Total Amount\s*[:\s]+Rs\.?\s*([\d,]+)/i);
            if (tm) amount = fmtAmt(parseInt(tm[1].replace(/,/g, '')));

            const accts = subj.includes('OD & CA') ? 'OD & CA' : subj.includes('OD') ? 'OD' : 'CA';
            fields = [
              { label: 'Company', value: 'Arisinfra Solutions Limited' },
              { label: 'Accounts', value: accts },
              { label: 'Total', value: amount },
            ];
            note = rows.length ? `${rows.length} vendors across ${accts} accounts. Full list in Excel.` : 'Payment list attached — awaiting your approval.';

          } else if (type === 'TD') {
            const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
            const body = extractPlainText(full.data.payload);

            // Extract party name from subject
            const party = subj
              .replace(/^APP-TD\s*[-–]?\s*(?:\d+\.?\d*\s*Cr?\s*)?(?:Deposit\s*[-–]?\s*)?/i, '')
              .replace(/\s*[-–]\s*(ASL|BIPL)\s*$/i, '')
              .replace(/-ASL$/i, '')
              .trim();

            const { rows: tdRows, total } = parseTDAmounts(body);
            if (total > 0) amount = fmtAmt(total);

            rows = tdRows.map(r => ({
              name: r.date,
              amount: fmtAmt(r.amount),
              purpose: 'Trade Deposit',
              category: 'Payables'
            }));

            const account = body.includes('HDFC') ? 'HDFC Account-9899' : body.includes('Axis') ? 'BIPL-Axis Bank' : 'See email';
            fields = [
              { label: 'Party', value: party || 'See email' },
              { label: 'Company', value: body.includes('Buildmex') ? 'Buildmex Infra Pvt Ltd' : 'Arisinfra Solutions Limited' },
              { label: 'Account', value: account },
              { label: 'Total', value: amount },
            ];
            // Smart TD note — what's being deposited, to whom, from which company
            const tdCompany = body.includes('Buildmex') ? 'BM' : 'ARIS';
            const tdPartyShort = shortName(party || '') || (party || '').split(' ').slice(0,3).join(' ');
            if (rows.length > 1) {
              note = `${rows.length} tranches to ${tdPartyShort || 'party'} via ${tdCompany}. Approval needed for fund release.`;
            } else {
              note = `Deposit to ${tdPartyShort || 'party'} via ${tdCompany}. Approval needed for fund release.`;
            }

          } else if (type === 'CRM') {
            const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
            const body = extractPlainText(full.data.payload);
            const trimmed = stripQuotes(body);
            risk = extractCrmRisk(trimmed);
            fields = extractCrmFields(trimmed);
            // Extract credit limit as the card amount
            const clField = fields.find(f => f.label === 'Credit Limit');
            if (clField) {
              const clMatch = clField.value.match(/([\d,]+)/);
              if (clMatch) {
                const n = parseInt(clMatch[1].replace(/,/g, ''));
                amount = fmtAmt(n);
              }
            }
            // Smart CRM note
            const custField = fields.find(f => f.label === 'Customer Name' || f.label === 'Party Name');
            const custName = custField ? custField.value.trim().split('\n')[0].split(' ').slice(0,3).join(' ') : '';
            const riskNote = risk ? ` Credit team flags ${risk} risk.` : '';
            note = custName
              ? `Credit limit requested for ${custName}.${riskNote} Review fields and approve or reject.`
              : `Credit limit approval required.${riskNote}`;

          } else if (type === 'TRF') {
            const snipLow = snippet.toLowerCase();
            if (snipLow.includes('approved') && status === 'fyi') approvalPill = 'Transfer Approved';
            const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
            const body = extractPlainText(full.data.payload);
            // Extract transfer details
            const fromMatch = body.match(/From:\s*(.+?)(?:\r?\n|To:)/i);
            const toMatch = body.match(/To:\s*(.+?)(?:\r?\n|Amount:)/i);
            const amtMatch = body.match(/Amount:\s*(.+?)(?:\r?\n|Transfer)/i);
            if (fromMatch) fields.push({ label: 'From', value: fromMatch[1].trim() });
            if (toMatch) fields.push({ label: 'To', value: toMatch[1].trim() });
            if (amtMatch) { fields.push({ label: 'Amount', value: amtMatch[1].trim() }); amount = amtMatch[1].trim(); }
            fields.push({ label: 'Approved By', value: 'Nishita' });
            const trfFrom = fromMatch ? fromMatch[1].trim().split('\n')[0] : '';
            const trfTo = toMatch ? toMatch[1].trim().split('\n')[0] : '';
            if (trfFrom && trfTo) {
              note = `${trfFrom} → ${trfTo}. Approved by Nishita — for your awareness.`;
            } else {
              note = 'Internal group transfer. Approved by Nishita — for your awareness.';
            }

          } else if (type === 'HR') {
            const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
            const body = stripQuotes(extractPlainText(full.data.payload));

            // Extract amount
            const amtMatch = body.match(/total payable amount.*?Rs\.?\s*\*?([\d,]+)\*?/i)
              || body.match(/Rs\.?\s*\*?([\d,]+)\*?\s*\/\-/i);
            if (amtMatch) amount = fmtAmt(parseInt(amtMatch[1].replace(/,/g,'')));

            // Build smart note — skip salutation lines, take the actual request
            const hrLines = body.split('\n')
              .filter(l => {
                const t = l.trim();
                return t.length > 10
                  && !t.startsWith('>')
                  && !/^(hi |dear |regards|thanks|--)/i.test(t)
                  && !/^(Arisinfra|Art Guild|L\.B\.S|Mumbai|Web:|Mob:|Unit)/i.test(t);
              });

            // Find the key request line
            const requestLine = hrLines.find(l => /please|approval|process|payable|kindly/i.test(l));
            const totalLine = hrLines.find(l => /total payable|payable amount/i.test(l));

            // Build category breakdown from body
            const catRows = [];
            const catRegex = /^(Fuel|Food|Local Conveyance|Driver|Repairs|Printing|Courier|Work Support|Flight|Accommodation|Mobile|Fastag).*?(\d[\d,]+)/gm;
            let cm;
            while ((cm = catRegex.exec(body)) !== null) {
              catRows.push({ name: cm[1].trim(), amount: fmtAmt(parseInt(cm[2].replace(/,/g,''))) });
            }

            // Department breakdown
            const deptRows = [];
            const deptRegex = /^(Sales and Marketing|Customer Relation|Administration|Operations|Accounts|Legal|Human Resource|Finance|Procurement|Tech).*?(\d[\d,]+)/gm;
            let dm;
            while ((dm = deptRegex.exec(body)) !== null) {
              deptRows.push({ label: dm[1].trim().split(' ').slice(0,2).join(' '), value: fmtAmt(parseInt(dm[2].replace(/,/g,''))) });
            }

            note = requestLine
              ? requestLine.replace(/\s+/g,' ').trim().slice(0,180)
              : (totalLine || 'Employee expense reimbursement request for approval.');

            fields = [
              { label: 'From', value: firstName(h.from) },
              ...(amount ? [{ label: 'Total', value: amount }] : []),
              ...deptRows.slice(0, 5)
            ];

          } else {
            note = snippet.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'");
            fields = [{ label: 'From', value: firstName(h.from) }];
          }

          if (status === 'fyi' && !approvalPill) {
            const sl = snippet.toLowerCase();
            if (sl.includes('approved')) {
              if (type === 'TRF') approvalPill = 'Transfer Approved';
              else if (type === 'CRM') approvalPill = 'Limit Approved';
              else approvalPill = 'Approved';
            }
          }

        } catch (e) {
          note = note || snippet.replace(/&amp;/g, '&').replace(/&#39;/g, "'") || 'Email received.';
          if (!fields.length) fields = [{ label: 'From', value: firstName(h.from) }];
        }

        smartTitleStr = smartTitle(type, subj, h.from, '', rows);
        return {
          tid: msg.threadId, mid: msg.id, cat: msg.cat, subj, type,
          title: smartTitleStr, amount, risk, fields, rows, note,
          from: h.from || '', date: h.date || '', to: h.to || '', cc: h.cc || '',
          isUnread, isCc, status, rohanApproved, approvalPill
        };
      };

      const pendingItems = await Promise.all([...pending, ...expDone, ...tdDone].slice(0, 30).map(processMsg));

      const doneItems = done.map(({ msg, h, status, isUnread, isCc }) => ({
        tid: msg.threadId, mid: msg.id, cat: msg.cat,
        subj: h.subject || '', type: detectType(h.subject || ''),
        title: smartTitle(detectType(h.subject||''), h.subject||'', h.from||'', '', []),
        amount: extractAmtFromSubject(h.subject || ''),
        risk: '', fields: [], rows: [], note: '',
        from: h.from || '', date: h.date || '', to: h.to || '', cc: h.cc || '',
        isUnread: false, isCc, status: 'done', rohanApproved: false, approvalPill: ''
      }));

      const items = [...pendingItems.filter(Boolean), ...doneItems];
      return res.json({ items, refreshedTokens: oauth2Client.credentials });

    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  if (action === 'mark-read') {
    try {
      await gmail.users.threads.modify({ userId: 'me', id: req.body.tid, requestBody: { removeLabelIds: ['UNREAD'] } });
      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  if (action === 'send') {
    const { tid, to, cc, subject, body } = req.body;
    try {
      const thread = await gmail.users.threads.get({ userId: 'me', id: tid, format: 'metadata' });
      const msgs = thread.data.messages || [];
      const last = msgs[msgs.length - 1];
      const lh = {};
      for (const h of (last?.payload?.headers || [])) lh[h.name.toLowerCase()] = h.value;
      const msgId = lh['message-id'] || '';
      const replySubj = subject.startsWith('Re:') ? subject : `Re: ${subject}`;
      const lines = [
        `From: Ronak Morbia <ronak@aris.in>`,
        `To: ${to}`,
        cc ? `Cc: ${cc}` : null,
        `Subject: ${replySubj}`,
        msgId ? `In-Reply-To: ${msgId}` : null,
        msgId ? `References: ${msgId}` : null,
        `Content-Type: text/plain; charset=utf-8`,
        '', body
      ].filter(l => l !== null);
      const raw = Buffer.from(lines.join('\r\n')).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
      await gmail.users.messages.send({ userId: 'me', requestBody: { raw, threadId: tid } });
      await gmail.users.threads.modify({ userId: 'me', id: tid, requestBody: { removeLabelIds: ['UNREAD'] } });
      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  // ── WHOAMI — return logged-in user's email ─────────────────────────────────
  if (action === 'whoami') {
    try {
      const profile = await gmail.users.getProfile({ userId: 'me' });
      return res.json({ email: profile.data.emailAddress, refreshedTokens: oauth2Client.credentials });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  // ── ROHAN INBOX — EXP emails where Rohan is To/CC ─────────────────────────
  if (action === 'rohan-inbox') {
    try {
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const afterDate = `${monthStart.getFullYear()}/${String(monthStart.getMonth()+1).padStart(2,'0')}/01`;

      // Fetch APP-EXP emails in Funds label this month
      const listRes = await gmail.users.messages.list({
        userId: 'me',
        q: `label:Funds APP-EXP after:${afterDate}`,
        maxResults: 50
      });
      const messages = listRes.data.messages || [];

      const items = await Promise.all(messages.map(async (msg) => {
        try {
          const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
          const hdrs = {};
          for (const h of (full.data.payload?.headers || [])) hdrs[h.name.toLowerCase()] = h.value;

          const subj = hdrs['subject'] || '';
          const from = hdrs['from'] || '';
          const to = hdrs['to'] || '';
          const cc = hdrs['cc'] || '';
          const date = hdrs['date'] || '';

          // Only include emails TO or CC rohan
          const allRecipients = (to + ' ' + cc).toLowerCase();
          if (!allRecipients.includes('rohan')) return null;

          // Extract body
          const body = extractPlainText(full.data.payload);
          const trimmedBody = stripQuotes(body);

          // Extract amount from subject
          const amtMatch = subj.match(/Rs\.?\s*([\d,]+)/i);
          let amount = '';
          if (amtMatch) {
            const n = parseInt(amtMatch[1].replace(/,/g,''));
            amount = fmtAmt(n);
          }

          // Extract total from body if not in subject
          if (!amount) {
            const bodyAmt = trimmedBody.match(/Total Amount\s*[:\s]+([\d,]+)/i) || trimmedBody.match(/Rs\.?\s*([\d,]+)\s*\/[-]?/i);
            if (bodyAmt) amount = fmtAmt(parseInt(bodyAmt[1].replace(/,/g,'')));
          }

          // Detect department
          const dept = (() => {
            const s = (subj + ' ' + from).toLowerCase();
            if (/\bhr\b|human resource/i.test(s)) return 'HR';
            if (/\bcrm\b|customer/i.test(s)) return 'CRM';
            if (/\btax\b/i.test(s)) return 'Tax';
            if (/\badmin\b/i.test(s)) return 'Admin';
            if (/\bops\b|operation/i.test(s)) return 'Operations';
            return 'Funds';
          })();

          // Check if Rohan already replied (submitted)
          const thread = await gmail.users.threads.get({ userId: 'me', id: msg.threadId, format: 'full' });
          const alreadySubmitted = (thread.data.messages || []).some(m => {
            const mFrom = (m.payload?.headers || []).find(h => h.name === 'From')?.value || '';
            const mBody = extractPlainText(m.payload);
            return mFrom.toLowerCase().includes('rohan') && mBody.includes('ROHAN-APPROVED');
          });

          return {
            mid: msg.id,
            tid: msg.threadId,
            subj,
            from: firstName(from) + ' <' + (from.match(/<([^>]+)>/)?.[1] || from) + '>',
            date: new Date(date).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' }),
            amount,
            dept,
            body: trimmedBody.slice(0, 1200),
            submitted: alreadySubmitted
          };
        } catch(e) { return null; }
      }));

      return res.json({ items: items.filter(Boolean), refreshedTokens: oauth2Client.credentials });
    } catch(e) { return res.status(500).json({ error: e.message }); }
  }

  return res.status(400).json({ error: 'Unknown action' });
}
