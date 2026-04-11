import { google } from 'googleapis';
import * as XLSX from 'xlsx';

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

function decodeBody(data) {
  try { return Buffer.from((data||'').replace(/-/g,'+').replace(/_/g,'/'), 'base64').toString('utf-8'); }
  catch { return ''; }
}

function parseExcelAllRows(base64Data) {
  try {
    const buf = Buffer.from((base64Data||'').replace(/-/g,'+').replace(/_/g,'/'), 'base64');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find header row
    let hIdx=-1, nameCol=-1, amtCol=-1;
    for (let i=0; i<Math.min(15,rows.length); i++) {
      const row = rows[i].map(c => String(c).toLowerCase().trim());
      const ni = row.findIndex(c => /vendor|party|name|payee|particular|beneficiary/i.test(c));
      const ai = row.findIndex(c => /^amount$|^amt$|payment amount|debit/i.test(c));
      const ai2 = ai>=0 ? ai : row.findIndex(c => /amount|amt/i.test(c) && !c.includes('words'));
      if (ni>=0 && ai2>=0) { hIdx=i; nameCol=ni; amtCol=ai2; break; }
    }
    if (hIdx<0) return [];

    const result = [];
    for (let i=hIdx+1; i<rows.length; i++) {
      const r = rows[i];
      const name = String(r[nameCol]||'').trim();
      const amtRaw = r[amtCol];
      const amt = typeof amtRaw==='number' ? amtRaw : parseFloat(String(amtRaw).replace(/[₹,\s]/g,''))||0;
      if (name && amt>0 && !/total|grand total|sub.?total/i.test(name)) {
        result.push({ name, amount: amt });
      }
    }
    return result;
  } catch(e) { return []; }
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { tokens } = req.method==='POST' ? req.body : {};
  if (!tokens) return res.status(400).json({ error: 'No tokens' });

  oauth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

  try {
    // Get current month range
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const afterDate = `${monthStart.getFullYear()}/${String(monthStart.getMonth()+1).padStart(2,'0')}/${String(monthStart.getDate()).padStart(2,'0')}`;

    // Fetch all VPAY emails this month
    const listRes = await gmail.users.messages.list({
      userId: 'me',
      q: `label:Funds in:inbox after:${afterDate} APP-VPAY`,
      maxResults: 50
    });

    const messages = listRes.data.messages || [];
    if (!messages.length) return res.json({ rows: [], totalEmails: 0 });

    // Process each VPAY email — download Excel and parse all rows
    const allRows = [];
    let totalEmails = 0;

    await Promise.all(messages.map(async (msg) => {
      try {
        const full = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'full' });
        const headers = {};
        for (const h of (full.data.payload?.headers||[])) headers[h.name.toLowerCase()] = h.value;

        const subj = headers.subject || '';
        const date = headers.date || '';

        // Find Excel attachment
        const findAtts = (parts) => {
          const atts = [];
          if (!parts) return atts;
          for (const p of parts) {
            if (p.filename && p.body?.attachmentId &&
                (p.mimeType?.includes('spreadsheet') || p.filename.endsWith('.xlsx')))
              atts.push({ name: p.filename, id: p.body.attachmentId });
            if (p.parts) atts.push(...findAtts(p.parts));
          }
          return atts;
        };
        const atts = findAtts([full.data.payload]);
        if (!atts.length) return;

        totalEmails++;

        const attRes = await gmail.users.messages.attachments.get({
          userId:'me', messageId:msg.id, id:atts[0].id
        });

        const rows = parseExcelAllRows(attRes.data.data);
        for (const row of rows) {
          allRows.push({ vendor: row.name, amount: row.amount, date, subj });
        }
      } catch(e) { /* skip failed */ }
    }));

    // Group by vendor name (normalised)
    const vendorMap = new Map();
    for (const row of allRows) {
      const key = row.vendor.trim().toLowerCase();
      if (!vendorMap.has(key)) {
        vendorMap.set(key, { vendor: row.vendor, total: 0, count: 0 });
      }
      const v = vendorMap.get(key);
      v.total += row.amount;
      v.count += 1;
    }

    // Sort by total descending
    const grouped = [...vendorMap.values()]
      .sort((a,b) => b.total - a.total);

    const grandTotal = grouped.reduce((s,r) => s+r.total, 0);

    return res.json({
      rows: grouped,
      grandTotal,
      totalEmails,
      month: monthStart.toLocaleString('en-IN', { month: 'long', year: 'numeric' })
    });

  } catch(e) {
    return res.status(500).json({ error: e.message });
  }
}
