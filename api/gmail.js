import { google } from 'googleapis';
import * as XLSX from 'xlsx';

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

// ── Helpers ───────────────────────────────────────────────────────────────────

function cleanTitle(subj) {
  return (subj || '')
    .replace(/^(Re:|RE:|Fwd:|FWD:)\s*/gi, '')
    .replace(/^APP-[A-Z]+-?\s*[-–:]\s*/gi, '')
    .replace(/^APP-[A-Z]+\s+/gi, '')
    .replace(/^CRM-APP:\s*/gi, '')
    .replace(/^HR-APP:\s*/gi, '')
    .replace(/\s+[-–]\s+ASL\s+\d{2}-\d{2}-\d{4}/gi, '')
    .replace(/\s+[-–]\s+\d{2}-\d{2}-\d{4}/gi, '')
    .replace(/\s+of\s+Rs\.?\s*[\d,]+\s*\/-?/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
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
  if (rs) { const n = parseInt(rs[1].replace(/,/g,'')); if (n > 0) return fmtAmt(n); }
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
  return (from || '').replace(/<[^>]+>/g, '').replace(/"/g, '').trim().split(/\s+/)[0];
}

function decodeBody(data) {
  try { return Buffer.from((data||'').replace(/-/g,'+').replace(/_/g,'/'), 'base64').toString('utf-8'); }
  catch { return ''; }
}

function extractPlainText(payload) {
  if (!payload) return '';
  if (payload.mimeType === 'text/plain' && payload.body?.data) return decodeBody(payload.body.data);
  if (payload.parts) { for (const p of payload.parts) { const t = extractPlainText(p); if (t) return t; } }
  return '';
}

function stripQuotes(text) {
  const cuts = [/\r?\nOn .{10,120}wrote:\r?\n/i, /\r?\nFrom:.*\r?\nSent:/i];
  for (const re of cuts) { const i = text.search(re); if (i > 100) return text.slice(0, i); }
  return text;
}

function categorise(name, purpose) {
  const s = (name + ' ' + purpose).toLowerCase();
  if (/salary|payroll|wages|stipend|staff/i.test(s)) return 'Salary';
  if (/interest|ncd|bank charge|factoring|od |loan repay|finance|processing fee/i.test(s)) return 'Finance Cost';
  return 'Expenses';
}

// ── Parse Excel attachment ────────────────────────────────────────────────────
function parseExcel(base64Data) {
  try {
    const b64 = (base64Data || '').replace(/-/g,'+').replace(/_/g,'/');
    const buf = Buffer.from(b64, 'base64');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find header row — look for row containing name-like and amount-like columns
    let hIdx = -1, nameCol = -1, amtCol = -1, purposeCol = -1;
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i].map(c => String(c).toLowerCase().trim());
      const ni = row.findIndex(c => /vendor|party|name|payee|particular|expense head|beneficiary/i.test(c));
      const ai = row.findIndex(c => /^amount$|^amt$|payment amount|debit|₹/i.test(c));
      // Also try looser match
      const ai2 = ai >= 0 ? ai : row.findIndex(c => /amount|amt/i.test(c) && !c.includes('in words'));
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
      const amt = typeof amtRaw === 'number' ? amtRaw : parseFloat(String(amtRaw).replace(/[₹,\s]/g,'')) || 0;
      if (name && amt > 0 && !/total|grand total|sub.?total/i.test(name)) {
        const purpose = purposeCol >= 0 ? String(r[purposeCol] || '').trim() : '';
        dataRows.push({ name, amount: amt, purpose, category: categorise(name, purpose) });
      }
    }

    // Sort by amount desc, return top 5
    dataRows.sort((a, b) => b.amount - a.amount);
    return dataRows.slice(0, 5).map(r => ({ ...r, amount: fmtAmt(r.amount) }));
  } catch(e) {
    console.error('Excel parse error:', e.message);
    return [];
  }
}

// ── Parse EXP rows from Rohan's reply body ───────────────────────────────────
function parseExpenseRows(body) {
  const rows = [];
  const lines = body.split('\n');
  let inTable = false;
  for (const line of lines) {
    const l = line.trim();
    if (/top\s+5\s+expenses/i.test(l)) { inTable = true; continue; }
    if (!inTable) continue;
    if (/breakdown:|best,|regards/i.test(l)) break;
    const m = l.match(/^(.+?)\s{2,}([\d,]+)\s*(.*)?$/);
    if (m) {
      const name = m[1].trim();
      const amt = parseInt(m[2].replace(/,/g, ''));
      const purpose = (m[3] || '').trim();
      if (name && amt > 0 && !/expense\s+head|amount/i.test(name)) {
        rows.push({ name, amount: fmtAmt(amt), purpose, category: categorise(name, purpose) });
      }
    }
  }
  return rows.slice(0, 5);
}

// ── Extract CRM risk ──────────────────────────────────────────────────────────
function extractCrmRisk(body) {
  const m = (body || '').match(/Recommendation from Credit Team\s*:\s*(High|Medium|Low)/i);
  return m ? m[1] : '';
}

// ── Extract CRM fields ────────────────────────────────────────────────────────
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

// ── Auth ──────────────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action } = req.query;

  if (action === 'auth-url') {
    const url = oauth2Client.generateAuthUrl({
      access_type: 'offline', prompt: 'consent',
      scope: ['https://www.googleapis.com/auth/gmail.readonly','https://www.googleapis.com/auth/gmail.send','https://www.googleapis.com/auth/gmail.modify']
    });
    return res.json({ url });
  }

  if (action === 'auth-callback') {
    try {
      const { tokens } = await oauth2Client.getToken(req.query.code);
      res.setHeader('Location', `/?tokens=${encodeURIComponent(JSON.stringify(tokens))}`);
      return res.status(302).end();
    } catch(e) { return res.status(500).json({ error: e.message }); }
  }

  const { tokens } = req.method === 'POST' ? req.body : {};
  if (!tokens) return res.status(400).json({ error: 'No tokens' });
  oauth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

  // ── INBOX ─────────────────────────────────────────────────────────────────
  if (action === 'inbox') {
    try {
      const [fundsRes, crmRes, hrRes, taxRes] = await Promise.all([
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds in:inbox', maxResults: 50 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:CRM in:inbox', maxResults: 30 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:HR in:inbox', maxResults: 20 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:Tax in:inbox', maxResults: 20 }),
      ]);

      const all = [
        ...(fundsRes.data.messages||[]).map(m=>({...m,cat:'Funds'})),
        ...(crmRes.data.messages||[]).map(m=>({...m,cat:'CRM'})),
        ...(hrRes.data.messages||[]).map(m=>({...m,cat:'HR'})),
        ...(taxRes.data.messages||[]).map(m=>({...m,cat:'Tax'})),
      ];

      // Deduplicate by threadId
      const seen = new Map();
      for (const m of all) { if (!seen.has(m.threadId)) seen.set(m.threadId, m); }
      const msgList = [...seen.values()];

      // Step 1: fetch metadata for ALL in parallel (fast)
      const metaList = await Promise.all(msgList.map(async msg => {
        try {
          const d = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'metadata', metadataHeaders:['From','To','Cc','Subject','Date'] });
          const h = {};
          for (const hdr of (d.data.payload?.headers||[])) h[hdr.name.toLowerCase()] = hdr.value;
          const labelIds = d.data.labelIds || [];
          const isUnread = labelIds.includes('UNREAD');
          const to = (h.to||'').toLowerCase();
          const isCc = !to.includes('ronak@aris.in') && !to.includes('ronak@arisinfra.one');
          const status = !isUnread ? 'done' : isCc ? 'fyi' : 'pending';
          const type = detectType(h.subject||'');
          return { msg, h, status, isUnread, isCc, type, snippet: d.data.snippet||'' };
        } catch { return null; }
      }));

      const valid = metaList.filter(Boolean);
      const pending = valid.filter(m => m.status !== 'done');
      const done = valid.filter(m => m.status === 'done');

      // Step 2: for pending/fyi fetch full message + thread if EXP
      const processMsg = async (meta) => {
        const { msg, h, status, isUnread, isCc, type, snippet } = meta;
        const subj = h.subject || '';
        let amount = extractAmtFromSubject(subj);
        let fields = [];
        let rows = [];
        let note = '';
        let risk = '';
        let rohanApproved = false;
        let approvalPill = '';

        try {
          if (type === 'EXP') {
            // Fetch full thread to get Rohan's reply
            const thread = await gmail.users.threads.get({ userId:'me', id:msg.threadId, format:'full' });
            const messages = thread.data.messages || [];

            // Find Rohan's message in the thread
            const rohanMsg = messages.find(m =>
              (m.payload?.headers||[]).find(hh => hh.name==='From' && hh.value.toLowerCase().includes('rohan'))
            );

            if (rohanMsg) {
              const rohanBody = extractPlainText(rohanMsg.payload);
              rows = parseExpenseRows(rohanBody);
              rohanApproved = true;
              note = 'Expenses pre-approved by Rohan. Top 5 breakdown provided.';
            }

            // Get total from Sandesh's original
            const sandeshMsg = messages.find(m =>
              (m.payload?.headers||[]).find(hh => hh.name==='From' && hh.value.toLowerCase().includes('sandesh'))
            );
            if (sandeshMsg) {
              const sandeshBody = extractPlainText(sandeshMsg.payload);
              const tm = sandeshBody.match(/Total Amount\s*[:\s]+([\d,]+)\s*\/-?/i);
              if (tm) amount = fmtAmt(parseInt(tm[1].replace(/,/g,'')));
            }

            fields = [
              { label: 'Company', value: 'Arisinfra Solutions Limited' },
              { label: 'Pre-Approved By', value: 'Rohan' },
              { label: 'Total Amount', value: amount },
            ];

          } else if (type === 'VPAY') {
            // Fetch full message to get Excel attachment
            const full = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'full' });
            const body = extractPlainText(full.data.payload);

            // Find xlsx attachment
            const findAtts = (parts) => {
              const atts = [];
              if (!parts) return atts;
              for (const p of parts) {
                if (p.filename && p.body?.attachmentId &&
                    (p.mimeType?.includes('spreadsheet') || p.filename.endsWith('.xlsx'))) {
                  atts.push({ name: p.filename, id: p.body.attachmentId });
                }
                if (p.parts) atts.push(...findAtts(p.parts));
              }
              return atts;
            };
            const atts = findAtts([full.data.payload]);

            if (atts.length > 0) {
              const attRes = await gmail.users.messages.attachments.get({ userId:'me', messageId:msg.id, id:atts[0].id });
              rows = parseExcel(attRes.data.data);
            }

            // Total from body
            const tm = body.match(/Total Amount\s*[:\s]+Rs\.?\s*([\d,]+)/i);
            if (tm) amount = fmtAmt(parseInt(tm[1].replace(/,/g,'')));

            fields = [
              { label: 'Company', value: 'Arisinfra Solutions Limited' },
              { label: 'Accounts', value: subj.includes('OD') ? 'OD & CA' : 'CA' },
              { label: 'Total', value: amount },
            ];
            note = rows.length ? `Top ${rows.length} vendors shown. Full list in attached Excel.` : 'Approved sheet attached.';

          } else if (type === 'CRM') {
            const full = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'full' });
            const body = extractPlainText(full.data.payload);
            const trimmed = stripQuotes(body);
            risk = extractCrmRisk(trimmed);
            fields = extractCrmFields(trimmed);
            note = `Credit approval request from Divya. ${risk ? risk + ' Risk.' : ''}`;

          } else if (type === 'TD') {
            const full = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'full' });
            const body = extractPlainText(full.data.payload);
            // Extract party name from subject
            const party = subj.replace(/^APP-TD\s*-\s*Deposit\s*-\s*/i,'').replace(/-ASL$/i,'').trim();
            // Extract amount rows from body — look for date + amount pattern
            const amtRows = [];
            let total = 0;
            for (const line of body.split('\n')) {
              const m = line.match(/(\d{2}-\d{2}-\d{4})\s+([\d,]+)/);
              if (m) {
                const amt = parseInt(m[2].replace(/,/g,''));
                total += amt;
                amtRows.push({ name: m[1], amount: fmtAmt(amt), purpose: 'Trade Deposit', category: 'Payables' });
              }
            }
            if (total > 0) amount = fmtAmt(total);
            rows = amtRows;
            fields = [
              { label: 'Party', value: party },
              { label: 'Company', value: 'Arisinfra Solutions Limited' },
              { label: 'Account', value: 'HDFC Account-9899' },
              { label: 'Total', value: amount },
            ];
            note = `Trade deposit request from Trupti for ${party}.`;

          } else if (type === 'TRF') {
            // Check if Nishita approved
            const bodyLow = snippet.toLowerCase();
            if (bodyLow.includes('approved') && status === 'fyi') approvalPill = 'Transfer Approved';
            fields = [{ label: 'Approved By', value: 'Nishita' }];
            note = 'Internal transfer for group companies. Approved by Nishita, you are in CC.';

          } else if (type === 'HR') {
            const full = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'full' });
            const body = stripQuotes(extractPlainText(full.data.payload)).slice(0, 1500);
            note = body.slice(0, 200).replace(/\r\n/g,' ').trim();
            fields = [{ label: 'From', value: firstName(h.from) }];

          } else {
            note = snippet.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&#39;/g,"'");
            fields = [{ label: 'From', value: firstName(h.from) }];
          }

          // Detect fyi approval pill
          if (status === 'fyi' && !approvalPill) {
            const sl = snippet.toLowerCase();
            if (sl.includes('approved')) {
              if (type === 'TRF') approvalPill = 'Transfer Approved';
              else if (type === 'CRM') approvalPill = 'Limit Approved';
              else approvalPill = 'Approved';
            }
          }

        } catch(e) {
          note = snippet.replace(/&amp;/g,'&').replace(/&#39;/g,"'") || 'Email received.';
          if (!fields.length) fields = [{ label: 'From', value: firstName(h.from) }];
        }

        return {
          tid: msg.threadId, mid: msg.id, cat: msg.cat,
          subj, type,
          title: cleanTitle(subj),
          amount, risk, fields, rows, note,
          from: h.from||'', date: h.date||'', to: h.to||'', cc: h.cc||'',
          isUnread, isCc, status, rohanApproved, approvalPill
        };
      };

      // Process pending/fyi in parallel
      const pendingItems = await Promise.all(pending.slice(0, 20).map(processMsg));

      // Done items — no AI, just regex
      const doneItems = done.map(({ msg, h, status, isUnread, isCc }) => ({
        tid: msg.threadId, mid: msg.id, cat: msg.cat,
        subj: h.subject||'', type: detectType(h.subject||''),
        title: cleanTitle(h.subject||''),
        amount: extractAmtFromSubject(h.subject||''),
        risk:'', fields:[], rows:[], note:'',
        from: h.from||'', date: h.date||'', to: h.to||'', cc: h.cc||'',
        isUnread: false, isCc, status:'done', rohanApproved: false, approvalPill:''
      }));

      const items = [...pendingItems.filter(Boolean), ...doneItems];
      return res.json({ items, refreshedTokens: oauth2Client.credentials });

    } catch(e) { return res.status(500).json({ error: e.message }); }
  }

  // ── MARK READ ─────────────────────────────────────────────────────────────
  if (action === 'mark-read') {
    try {
      await gmail.users.threads.modify({ userId:'me', id:req.body.tid, requestBody:{ removeLabelIds:['UNREAD'] } });
      return res.json({ ok:true, refreshedTokens: oauth2Client.credentials });
    } catch(e) { return res.status(500).json({ error: e.message }); }
  }

  // ── SEND ──────────────────────────────────────────────────────────────────
  if (action === 'send') {
    const { tid, to, cc, subject, body } = req.body;
    try {
      const thread = await gmail.users.threads.get({ userId:'me', id:tid, format:'metadata' });
      const msgs = thread.data.messages || [];
      const last = msgs[msgs.length-1];
      const lh = {};
      for (const h of (last?.payload?.headers||[])) lh[h.name.toLowerCase()] = h.value;
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
      const raw = Buffer.from(lines.join('\r\n')).toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
      await gmail.users.messages.send({ userId:'me', requestBody:{ raw, threadId:tid } });
      await gmail.users.threads.modify({ userId:'me', id:tid, requestBody:{ removeLabelIds:['UNREAD'] } });
      return res.json({ ok:true, refreshedTokens: oauth2Client.credentials });
    } catch(e) { return res.status(500).json({ error: e.message }); }
  }

  return res.status(400).json({ error: 'Unknown action' });
}
