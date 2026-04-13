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
  if (n >= 10000000) return '₹' + Math.round(n / 10000000) + ' Cr';
  if (n >= 100000) return '₹' + Math.round(n / 100000) + ' L';
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

      rows.push({ name, amount: fmtAmt(amt), purpose: purpose || derivePurpose(name), category: cat });
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


// Derive human-readable purpose/nature from expense name
function derivePurpose(name) {
  const n = (name || '').toLowerCase();
  if (/fuel|petrol|diesel|fastag|toll|parking|local conveyance|driver/i.test(name)) return 'Travel';
  if (/flight|train|ticket|accommodation|hotel|stay/i.test(name)) return 'Travel & Stay';
  if (/food|meal|canteen|restaurant|catering|lunch|dinner|snack/i.test(name)) return 'Food';
  if (/legal|arbitration|court|advocate|solicitor|law firm|vesta|s&t/i.test(name)) return 'Legal';
  if (/aws|azure|google cloud|techmagify|sazs|slack|dropbox|figma|openai|anthropic|atlassian|rebrandly|gupshup|cursor|relic|creativeit|software|saas/i.test(name)) return 'Software';
  if (/salary|payroll|wages|stipend|staff|hr pay|reimbursement/i.test(name)) return 'Payroll';
  if (/ncd|interest|factoring|texterity|sunx.*fee|processing fee|bank charge/i.test(name)) return 'Finance';
  if (/insurance|tata aig|rubix/i.test(name)) return 'Insurance';
  if (/advisory|consultant|rbsa|thanawala|manish bohra|actuar/i.test(name)) return 'Advisory';
  if (/rent|utility|housekeeping|stationery|printing|courier|office|admin/i.test(name)) return 'Admin';
  if (/mobile|airtel|phone|broadband|recharge/i.test(name)) return 'Telecom';
  if (/repair|maintenance|it support/i.test(name)) return 'Maintenance';
  if (/corp cc|credit card/i.test(name)) return 'Corp Card';
  return 'Other';
}


// Robust amount extraction:
// 1. Try subject first (fastest)
// 2. Try body for explicit total
// 3. Cross-verify: if both found and close, use body (more precise)
// 4. If only one found, use that
// 5. If neither, sum amounts found in body
function extractAmount(subj, body) {
  // Parse any amount string → number
  const parseNum = (s) => {
    if (!s) return 0;
    s = s.replace(/[₹,\s]/g, '');
    const cr  = s.match(/(\d+\.?\d*)\s*cr/i);  if (cr)  return parseFloat(cr[1])  * 10000000;
    const lac = s.match(/(\d+\.?\d*)\s*lac/i); if (lac) return parseFloat(lac[1]) * 100000;
    const l   = s.match(/(\d+\.?\d*)\s*l$/i);  if (l)   return parseFloat(l[1])   * 100000;
    return parseFloat(s) || 0;
  };

  // Try subject: Rs. X,XX,XXX or X Cr or X L
  let subjAmt = 0;
  const sr = subj.match(/Rs\.?\s*([\d,]+)/i);
  if (sr) subjAmt = parseNum(sr[1]);
  if (!subjAmt) {
    const cr = subj.match(/(\d+\.?\d*)\s*Cr/i); if (cr) subjAmt = parseFloat(cr[1]) * 10000000;
  }
  if (!subjAmt) {
    const l = subj.match(/(\d+\.?\d*)\s*L(?:acs?|akhs?)?/i); if (l) subjAmt = parseFloat(l[1]) * 100000;
  }

  if (!body) return subjAmt > 0 ? fmtAmt(subjAmt) : '';

  // Try body for explicit totals — multiple patterns
  let bodyAmt = 0;
  const bodyPatterns = [
    /Total Amount\s*[:\s]+Rs\.?\s*([\d,]+)/i,
    /Total Amount\s*[:\s]+([\d,]+)\s*\/\-?/i,
    /total payable amount[^\d]*([\d,]+)/i,
    /Payment Amount\s*[:\s]+Rs\s*([\d,]+)/i,
    /Amount\s*[:\s]+Rs\.?\s*([\d,]+)/i,
    /Rs\.?\s*\*?([\d,]+)\*?\s*\/\-/i,
  ];
  for (const pat of bodyPatterns) {
    const m = body.match(pat);
    if (m) {
      const n = parseInt(m[1].replace(/,/g, ''));
      if (n > 10000) { bodyAmt = n; break; }
    }
  }

  // Cross-verify: if both found
  if (subjAmt > 0 && bodyAmt > 0) {
    // If they're within 10% of each other, use body (more precise)
    const diff = Math.abs(subjAmt - bodyAmt) / Math.max(subjAmt, bodyAmt);
    return fmtAmt(diff < 0.1 ? bodyAmt : Math.max(subjAmt, bodyAmt));
  }

  // Only one found
  if (bodyAmt > 0) return fmtAmt(bodyAmt);
  if (subjAmt > 0) return fmtAmt(subjAmt);

  // Last resort: sum all Rs. amounts found in body (for multi-tranche emails)
  let total = 0;
  const allAmts = [...body.matchAll(/Rs\.?\s*([\d,]+)/gi)];
  for (const m of allAmts) {
    const n = parseInt(m[1].replace(/,/g,''));
    if (n > 10000) total += n;
  }
  return total > 0 ? fmtAmt(total) : '';
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
    // Debug: log env vars (lengths only, not values) to help diagnose auth issues
    const clientId = (process.env.GMAIL_CLIENT_ID||'').trim();
    const clientSecret = (process.env.GMAIL_CLIENT_SECRET||'').trim();
    const redirectUri = (process.env.GMAIL_REDIRECT_URI||'').trim();
    if (!clientId || !clientSecret || !redirectUri) {
      return res.status(500).json({ error: `Missing env vars — CLIENT_ID:${clientId.length} CLIENT_SECRET:${clientSecret.length} REDIRECT_URI:${redirectUri}` });
    }
    // Recreate client with trimmed values to avoid whitespace issues
    const cleanClient = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
    const url = cleanClient.generateAuthUrl({
      access_type: 'offline', prompt: 'consent',
      scope: ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify']
    });
    return res.json({ url });
  }

  if (action === 'auth-callback') {
    const code = req.query.code;
    if (!code) return res.status(400).json({ error: 'No code in callback' });
    try {
      const cleanClient = new google.auth.OAuth2(
        (process.env.GMAIL_CLIENT_ID||'').trim(),
        (process.env.GMAIL_CLIENT_SECRET||'').trim(),
        (process.env.GMAIL_REDIRECT_URI||'').trim()
      );
      const { tokens } = await cleanClient.getToken(code);
      if (!tokens || !tokens.access_token) return res.status(500).json({ error: 'No tokens returned from Google' });
      const encoded = encodeURIComponent(JSON.stringify(tokens));
      res.setHeader('Location', `/?tokens=${encoded}`);
      res.setHeader('Cache-Control', 'no-store');
      return res.status(302).end();
    } catch (e) { return res.status(500).json({ error: 'Token exchange failed: ' + e.message }); }
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

      // Dedup by threadId — keep the OLDEST message (first in thread = original request)
      // Gmail returns messages in reverse chronological order, so last = oldest
      const seen = new Map();
      for (const m of [...all].reverse()) { seen.set(m.threadId, m); }
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

      let valid = metaList.filter(Boolean);

      // Deduplicate HR cards where Gmail broke the thread into separate threadIds
      // Strategy: normalize subject (strip Re:/RE: prefixes), keep oldest per subject,
      // and scan newer duplicate threads for approval signals to surface as pills
      {
        const hrBySubj = new Map(); // normSubj → meta (oldest/primary)
        const hrToRemove = new Set();

        for (const m of valid) {
          if (m.type !== 'HR') continue;
          const normSubj = (m.h.subject || '')
            .replace(/^(RE:|Re:|Fwd:|FWD:)\s*/gi, '')
            .trim().toLowerCase();
          const mDate = new Date(m.h.date || 0).getTime();

          if (!hrBySubj.has(normSubj)) {
            hrBySubj.set(normSubj, m);
          } else {
            const primary = hrBySubj.get(normSubj);
            const pDate = new Date(primary.h.date || 0).getTime();
            // Mark the newer one for removal, but check it for approvals
            const [older, newer] = mDate < pDate ? [m, primary] : [primary, m];
            // If the newer is just an approval reply (short snippet), absorb it
            const snip = (newer.snippet || '').toLowerCase();
            if (/^(please proceed|approved|okay|ok)/i.test(snip.trim())) {
              // Tag the older (primary) with approval info
              const fromName = firstName(newer.h.from || '');
              older._absorbedApproval = 'Amount Approved';
            }
            hrBySubj.set(normSubj, older);
            hrToRemove.add(newer.msg.threadId);
          }
        }
        valid = valid.filter(m => !hrToRemove.has(m.msg.threadId));
      }

      // EXP done items need full thread processing to get Rohan's category breakdown for KPIs
      // All other done items just get metadata (fast path)
      // Split: pending/fyi first (full processing), done items (lighter processing)
      const pending = valid.filter(m => m.status !== 'done');
      const doneAll = valid.filter(m => m.status === 'done');
      // EXP and TD done need full thread/body for amounts and rows
      const expDone = doneAll.filter(m => m.type === 'EXP');
      const tdDone = doneAll.filter(m => m.type === 'TD');
      // Other done items just need a quick body fetch for amount/note
      const doneLite = doneAll.filter(m => m.type !== 'EXP' && m.type !== 'TD');

      const processMsg = async (meta) => {
        const { msg, h, status, isUnread, isCc, type, snippet } = meta;
        const subj = h.subject || '';
        let amount = extractAmtFromSubject(subj); // pre-fill from subject, overridden by body
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
                isUnread, isCc, status: 'cancelled', rohanApproved: false, approvalPill: 'Cancelled'
              };
            }

            // Find Rohan's reply for expense breakdown
            const rohanMsg = messages.find(m =>
              (m.payload?.headers || []).find(hh => hh.name === 'From' && hh.value.toLowerCase().includes('rohan'))
            );
            // Check if Ronak has already approved
            const ronakApprovalMsg = messages.find(m => {
              const fromHdr = (m.payload?.headers || []).find(hh => hh.name === 'From');
              const from = (fromHdr?.value || '').toLowerCase();
              if (!from.includes('ronak')) return false;
              const body = extractPlainText(m.payload || {});
              return /^approved|^okay|^ok|^done/i.test(body.trim());
            });

            if (ronakApprovalMsg) approvalPill = 'Approved';

            if (rohanMsg) {
              const rohanBody = extractPlainText(rohanMsg.payload);
              rows = parseExpenseRows(rohanBody);
              rohanApproved = true;
              if (approvalPill === 'Approved') {
                note = `Approved by you. ${rows.length} expense items, pre-approved by Rohan.`;
              } else {
                note = `${rows.length} expense items pre-approved by Rohan. Review breakdown and approve.`;
              }
            } else {
              note = approvalPill === 'Approved'
                ? 'Approved by you. Full expense list in attached Excel.'
                : 'Pre-approved by Rohan. Full list in attached Excel. Awaiting your sign-off.';
              rohanApproved = true;
            }

            // Total from Sandesh's original message
            const sandeshMsg = messages.find(m =>
              (m.payload?.headers || []).find(hh => hh.name === 'From' && hh.value.toLowerCase().includes('sandesh'))
            );
            if (sandeshMsg) {
              const sb = extractPlainText(sandeshMsg.payload);
              amount = extractAmount(subj, sb);
            }

            fields = [
              { label: 'Company', value: 'Arisinfra Solutions Limited' },
              { label: 'Pre-Approved By', value: 'Rohan' },
              { label: 'Total Amount', value: amount },
            ];

          } else if (type === 'VPAY') {
            const full = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
            const body = extractPlainText(full.data.payload);

            // Set amount from body FIRST — so it's available even if Excel fetch fails
            amount = extractAmount(subj, body);

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

            // Only download Excel for pending VPAY (rows needed for display)
            // Done VPAY: amount+note sufficient, skip expensive Excel download
            if (atts.length > 0 && status === 'pending') {
              try {
                const attRes = await gmail.users.messages.attachments.get({ userId: 'me', messageId: msg.id, id: atts[0].id });
                rows = parseExcel(attRes.data.data, 'VPAY');
              } catch(attErr) { /* Excel fetch failed — amount already set from body */ }
            }

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
            amount = total > 0 ? fmtAmt(total) : extractAmount(subj, body);

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
            if (!amount) amount = extractAmount(subj, body);
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

            // Parse ALL transfer blocks — repeating From/To/Amount pattern
            const transfers = [];
            let totalAmt = 0;
            // Split into paragraphs and find each transfer block
            const lines = body.split(/\r?\n/);
            let cur = {};
            for (const line of lines) {
              const l = line.trim();
              const fromM = l.match(/^From:\s*(.+)/i);
              const toM   = l.match(/^To:\s*(.+)/i);
              const amtM  = l.match(/^Amount:\s*(.+)/i);
              if (fromM) { cur.from = fromM[1].trim(); }
              if (toM)   { cur.to   = toM[1].trim(); }
              if (amtM)  {
                cur.amtRaw = amtM[1].trim();
                // Parse amount value
                const cr  = cur.amtRaw.match(/(\d+\.?\d*)\s*cr/i);
                const lac = cur.amtRaw.match(/(\d+\.?\d*)\s*lac/i);
                const lak = cur.amtRaw.match(/(\d+\.?\d*)\s*l\b/i);
                const num = cr  ? parseFloat(cr[1])  * 10000000
                          : lac ? parseFloat(lac[1]) * 100000
                          : lak ? parseFloat(lak[1]) * 100000
                          : parseFloat(cur.amtRaw.replace(/[₹,]/g,'')) || 0;
                cur.amtNum = num;
                // Once we have all three, save transfer
                if (cur.from && cur.to) {
                  totalAmt += num;
                  transfers.push({
                    from: shortName(cur.from),
                    to:   shortName(cur.to),
                    amt:  fmtAmt(num),
                    amtRaw: cur.amtRaw
                  });
                  cur = {};
                }
              }
            }

            amount = totalAmt > 0 ? fmtAmt(totalAmt) : extractAmount(subj, body);

            // Fields — one row per transfer
            if (transfers.length > 0) {
              fields = transfers.map((t, i) => ({
                label: transfers.length > 1 ? `Transfer ${i + 1}` : 'Transfer',
                value: `${t.from} → ${t.to} · ${t.amt}`
              }));
            }
            fields.push({ label: 'Approved By', value: 'Nishita' });

            // Smart note
            if (transfers.length === 1) {
              note = `${transfers[0].from} → ${transfers[0].to} (${transfers[0].amt}). Approved by Nishita — for your awareness.`;
            } else if (transfers.length > 1) {
              const summary = transfers.map(t => `${t.to} ${t.amt}`).join(', ');
              note = `${transfers.length} transfers totalling ${amount} — ${summary}. Approved by Nishita.`;
            } else {
              note = 'Internal group transfer. Approved by Nishita — for your awareness.';
            }

          } else if (type === 'HR') {
            // Fetch the full thread — gets all messages including any broken-thread replies
            const thread = await gmail.users.threads.get({ userId: 'me', id: msg.threadId, format: 'full' });
            const allMsgs = thread.data.messages || [];

            // Get the earliest message body (original request)
            const originalMsg = allMsgs[0];
            const body = stripQuotes(extractPlainText(originalMsg?.payload || {}));

            // Check if anyone has approved in any message in this thread
            for (const tm of allMsgs) {
              const fromHdr = (tm.payload?.headers || []).find(hh => hh.name === 'From');
              const tmFrom = (fromHdr?.value || '').toLowerCase();
              const tmBody = extractPlainText(tm.payload || {}).trim();
              const firstLine = tmBody.split('\n')[0].trim();
              if (/^(please proceed|approved|okay|ok|done)/i.test(firstLine)) {
                approvalPill = 'Amount Approved';
              }
            }

            // Check if absorbed approval from broken thread should set pill now
            if (!approvalPill && meta._absorbedApproval) approvalPill = meta._absorbedApproval;

            // Extract amount
            amount = extractAmount(subj, body);

            // Smart title — use what the email is about, not subject
            const titleMatch = body.match(/Employee Expense\s+([A-Za-z]+-?\d*)/i)
              || body.match(/Expense report\s*[-–]?\s*([A-Za-z]+-?\d*)/i)
              || body.match(/approval for\s+(.{5,40}?)\s*\.?\s*(?:The|Please|Rs)/i);
            smartTitleStr = titleMatch
              ? 'Employee Expenses — ' + titleMatch[1].trim()
              : subj.replace(/^(RE:|Re:|HR-APP:|APP-HR:)\s*/gi, '').trim().slice(0, 50);

            // Parse expense category rows — mirrors EXP card style
            const hrRows = [];
            const expCats = [
              ['Fuel',                /^Fuel\s+([\d,]+)/m],
              ['Food Expense',        /^Food Expense\s+([\d,]+)/m],
              ['Local Conveyance',    /^Local Conveyance\s+([\d,]+)/m],
              ['Driver Allowance',    /^Driver Allowance\s+([\d,]+)/m],
              ['Repairs & Maintenance', /^Repairs & Maintenance\s+([\d,]+)/m],
              ['Printing & Stationary', /^Printing & Stationary\s+([\d,]+)/m],
              ['Courier',             /^Courier\s+([\d,]+)/m],
              ['Work Support',        /^Work Support Expenses\s+([\d,]+)/m],
              ['Flight & Train',      /^Flight & Train Ticket\s+([\d,]+)/m],
              ['Accommodation',       /^Accommodation\s+([\d,]+)/m],
              ['Mobile Reimb.',       /^Mobile Reimbursement\s+([\d,]+)/m],
              ['Fastag',              /^Fastag\s+([\d,]+)/m],
            ];
            for (const [name, re] of expCats) {
              const m2 = body.match(re);
              if (m2) {
                const amt = parseInt(m2[1].replace(/,/g,''));
                if (amt > 0) {
                  // Category is always one of the 5 accounting categories
                const cat = categorise(name, '', 'EXP');
                  hrRows.push({ name, amount: fmtAmt(amt), purpose: derivePurpose(name), category: cat });
                }
              }
            }

            // Department breakdown as fields
            const deptRows = [];
            const deptPairs = [
              ['Sales & Mktg',     /Sales and Marketing\s+([\d,]+)/],
              ['CRM',              /Customer Relation[^\n]+\s+([\d,]+)/],
              ['Admin',            /Administration\s+([\d,]+)/],
              ['Operations',       /^Operations\s+([\d,]+)/m],
              ['Accounts Ops',     /Accounts Operations\s+([\d,]+)/],
              ['Legal',            /Legal & Compliance\s+([\d,]+)/],
              ['HR',               /Human Resource\s+([\d,]+)/],
              ['Finance',          /Finance & Accounts\s+([\d,]+)/],
              ['Tech',             /Tech & Engineering\s+([\d,]+)/],
            ];
            for (const [label, re] of deptPairs) {
              const dm = body.match(re);
              if (dm) {
                const amt = parseInt(dm[1].replace(/,/g,''));
                if (amt > 0) deptRows.push({ label, value: fmtAmt(amt) });
              }
            }

            rows = hrRows; // show in expanded card like EXP

            fields = [
              { label: 'Requested By', value: firstName(h.from) },
              ...(amount ? [{ label: 'Total', value: amount }] : []),
              ...deptRows.slice(0, 5),
            ];

            const period = titleMatch?.[1] || 'the month';
            if (approvalPill === 'Amount Approved') {
              note = amount
                ? `Rohan has approved ${amount} for ${period}. Expense list pre-approved — awaiting your sign-off.`
                : `Rohan has approved expenses for ${period}. Awaiting your sign-off.`;
            } else {
              note = amount
                ? `Employee reimbursements for ${period} — ${amount} total. Rohan's approval pending before you can approve.`
                : `Employee expense reimbursement request for ${period}. Rohan's approval pending.`;
            }

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

        smartTitleStr = smartTitleStr || smartTitle(type, subj, h.from, '', rows);
        if (meta._absorbedApproval && !approvalPill) approvalPill = meta._absorbedApproval;

        // Final fallback — use snippet as body proxy
        if (!amount) amount = extractAmount(subj, snippet);

        // Ensure done items always have a note
        if (status === 'done' && !note) {
          note = amount ? `Approved — ${amount}.` : 'Approved.';
        }
        return {
          tid: msg.threadId, mid: msg.id, cat: msg.cat, subj, type,
          title: smartTitleStr, amount, risk, fields, rows, note,
          from: h.from || '', date: h.date || '', to: h.to || '', cc: h.cc || '',
          isUnread, isCc, status, rohanApproved, approvalPill
        };
      };

      // Process in parallel with priority — pending first, then done
      const [pendingItems, expDoneItems, tdDoneItems, liteItems] = await Promise.all([
        Promise.all(pending.slice(0, 20).map(processMsg)),
        Promise.all(expDone.slice(0, 8).map(processMsg)),
        Promise.all(tdDone.slice(0, 8).map(processMsg)),
        Promise.all(doneLite.slice(0, 15).map(processMsg)),
      ]);
      const items = [...pendingItems, ...expDoneItems, ...tdDoneItems, ...liteItems].filter(Boolean);
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

  // ── ROHAN INBOX ─────────────────────────────────────────────────────────────
  // Original email = what department sent TO Rohan (Ronak in CC).
  // approved = Rohan has replied on the thread = his approval forwarded to Ronak.
  // pending  = no reply from Rohan yet.
  if (action === 'rohan-inbox') {
    try {
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const afterDate = `${monthStart.getFullYear()}/${String(monthStart.getMonth()+1).padStart(2,'0')}/01`;

      const listRes = await gmail.users.messages.list({
        userId: 'me',
        q: `label:Funds APP-EXP after:${afterDate}`,
        maxResults: 50
      });
      const messages = listRes.data.messages || [];

      // One item per thread
      const seenThreads = new Set();
      const threads = messages.filter(m => {
        if (seenThreads.has(m.threadId)) return false;
        seenThreads.add(m.threadId);
        return true;
      });

      const items = await Promise.all(threads.map(async (msg) => {
        try {
          const thread = await gmail.users.threads.get({ userId: 'me', id: msg.threadId, format: 'full' });
          const threadMsgs = thread.data.messages || [];

          // Original = first message in thread (department → Rohan)
          const original = threadMsgs[0];
          if (!original) return null;

          const hdrs = {};
          for (const h of (original.payload?.headers || [])) hdrs[h.name.toLowerCase()] = h.value;

          const subj = hdrs['subject'] || '';
          const from = hdrs['from'] || '';
          const to   = hdrs['to']   || '';
          const cc   = hdrs['cc']   || '';
          const date = hdrs['date'] || '';

          // Only threads where Rohan is in To or CC of the original email
          if (!(to + ' ' + cc).toLowerCase().includes('rohan')) return null;

          // Body = original email (what department sent to Rohan)
          // Do NOT stripQuotes — original email content IS the body, not a quoted reply
          const body = extractPlainText(original.payload);

          // Amount from subject, fallback to body
          let amount = '';
          const amtSubj = subj.match(/Rs\.?\s*([\d,]+)/i);
          if (amtSubj) amount = fmtAmt(parseInt(amtSubj[1].replace(/,/g, '')));
          if (!amount) {
            const amtBody = body.match(/Total Amount\s*[:\s]+([\d,]+)/i)
                         || body.match(/Rs\.?\s*([\d,]+)\s*\/[-]?/i);
            if (amtBody) amount = fmtAmt(parseInt(amtBody[1].replace(/,/g, '')));
          }

          // Department
          const dept = (() => {
            const s = (subj + ' ' + from).toLowerCase();
            if (/salary|payroll|wages|\bhr\b|human resource/i.test(s)) return 'HR';
            if (/\bcrm\b|customer relation/i.test(s)) return 'CRM';
            if (/\btax\b/i.test(s)) return 'Tax';
            if (/\badmin\b/i.test(s)) return 'Admin';
            if (/tech|software|saas|cloud|aws|figma|cursor/i.test(s)) return 'Tech';
            if (/\bops\b|operation/i.test(s)) return 'Operations';
            return 'Funds';
          })();

          // approved = Rohan has replied on this thread after the original email
          const rohanReply = threadMsgs.slice(1).find(m => {
            const mFrom = (m.payload?.headers || []).find(h => h.name === 'From')?.value || '';
            return mFrom.toLowerCase().includes('rohan');
          });
          const approved = !!rohanReply;

          // Capture Rohan's reply as his approval note
          const rohanNote = rohanReply
            ? stripQuotes(extractPlainText(rohanReply.payload)).slice(0, 300)
            : '';

          return {
            mid: original.id,
            tid: msg.threadId,
            subj,
            from: firstName(from),
            date: new Date(date).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' }),
            amount,
            dept,
            body: body.slice(0, 2500),  // full original email for AI to read
            approved,                    // true = Rohan approved, false = pending
            rohanNote                    // Rohan's reply if already approved
          };
        } catch(e) { return null; }
      }));

      return res.json({ items: items.filter(Boolean), refreshedTokens: oauth2Client.credentials });
    } catch(e) { return res.status(500).json({ error: e.message }); }
  }

  return res.status(400).json({ error: 'Unknown action' });
}
