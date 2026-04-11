import { google } from 'googleapis';
import * as XLSX from 'xlsx';

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

function parseExcelAllRows(base64Data) {
  try {
    const buf = Buffer.from((base64Data||'').replace(/-/g,'+').replace(/_/g,'/'), 'base64');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    let hIdx=-1, nameCol=-1, amtCol=-1;
    for (let i=0; i<Math.min(15,rows.length); i++) {
      const row = rows[i].map(c => String(c).toLowerCase().trim());
      const ni = row.findIndex(c => /vendor|party|name|payee|particular|beneficiary|expense head/i.test(c));
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
      if (name && amt>0 && !/total|grand total|sub.?total/i.test(name))
        result.push({ name, amount: amt });
    }
    return result;
  } catch(e) { return []; }
}

function categoriseRow(name, purpose) {
  const s = (name+' '+(purpose||'')).toLowerCase();
  if (/interest|ncd|bank charge|factoring|od |loan repay|finance|processing fee/i.test(s)) return 'finance';
  if (/\baws\b|\bazure\b|google cloud|techmagify|sazs apps|\bslack\b|\bdropbox\b|\bfigma\b|openai|anthropic|atlassian|rebrandly|gupshup|\bcursor\b|creativeit/i.test(s)) return 'technology';
  if (/salary|payroll|wages|stipend/i.test(s)) return 'salary';
  return 'expenses';
}

// Parse Rohan's EXP breakdown from body text — returns all rows with category
function parseExpBody(body) {
  const rows = [];
  const lines = body.split('\n');
  let section = null;
  for (const line of lines) {
    const l = line.trim();
    if (/interest\s+expenses/i.test(l)) { section='finance'; continue; }
    if (/top\s+5\s+expenses/i.test(l)) { section='expenses'; continue; }
    if (/^best,|^regards|^thanks/i.test(l)) break;
    if (!section) continue;
    if (/expense\s+head|amount\s*\(rs/i.test(l)) continue;
    const m = l.match(/^(.+?)\s{2,}(\d[\d,]+)\s*(.*)?$/);
    if (m) {
      const name = m[1].trim();
      const amt = parseInt(m[2].replace(/,/g,''));
      const purpose = (m[3]||'').trim();
      if (name && amt>0) {
        const cat = section==='finance' ? 'finance' : categoriseRow(name, purpose);
        rows.push({ name, amount: amt, category: cat });
      }
    }
  }
  return rows;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { tokens, type = 'payables' } = req.method==='POST' ? req.body : {};
  if (!tokens) return res.status(400).json({ error: 'No tokens' });

  oauth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

  try {
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const afterDate = `${monthStart.getFullYear()}/${String(monthStart.getMonth()+1).padStart(2,'0')}/01`;
    const month = monthStart.toLocaleString('en-IN', { month: 'long', year: 'numeric' });

    const allRows = [];
    let totalEmails = 0;

    if (type === 'payables') {
      // Fetch all VPAY emails — download Excel attachments
      const listRes = await gmail.users.messages.list({
        userId:'me', q:`label:Funds in:inbox after:${afterDate} APP-VPAY`, maxResults:50
      });
      const messages = listRes.data.messages || [];

      await Promise.all(messages.map(async (msg) => {
        try {
          const full = await gmail.users.messages.get({ userId:'me', id:msg.id, format:'full' });
          const findAtts = (parts) => {
            const atts = [];
            if (!parts) return atts;
            for (const p of parts) {
              if (p.filename && p.body?.attachmentId && (p.mimeType?.includes('spreadsheet') || p.filename.endsWith('.xlsx')))
                atts.push({ id: p.body.attachmentId });
              if (p.parts) atts.push(...findAtts(p.parts));
            }
            return atts;
          };
          const atts = findAtts([full.data.payload]);
          if (!atts.length) return;
          totalEmails++;
          const attRes = await gmail.users.messages.attachments.get({ userId:'me', messageId:msg.id, id:atts[0].id });
          const rows = parseExcelAllRows(attRes.data.data);
          for (const r of rows) allRows.push({ name: r.name, amount: r.amount });
        } catch(e) {}
      }));

    } else {
      // Fetch all EXP emails — parse Rohan's breakdown from thread
      const listRes = await gmail.users.messages.list({
        userId:'me', q:`label:Funds in:inbox after:${afterDate} APP-EXP`, maxResults:30
      });
      const messages = listRes.data.messages || [];

      await Promise.all(messages.map(async (msg) => {
        try {
          const thread = await gmail.users.threads.get({ userId:'me', id:msg.threadId, format:'full' });
          const rohanMsg = (thread.data.messages||[]).find(m =>
            (m.payload?.headers||[]).find(h => h.name==='From' && h.value.toLowerCase().includes('rohan'))
          );
          if (!rohanMsg) return;

          // Check not cancelled
          const latestMsg = thread.data.messages[thread.data.messages.length-1];
          const latestBody = latestMsg?.payload?.parts?.[0]?.body?.data
            ? Buffer.from(latestMsg.payload.parts[0].body.data.replace(/-/g,'+').replace(/_/g,'/'), 'base64').toString('utf-8')
            : '';
          if (/please ignore|kindly ignore/i.test(latestBody)) return;

          totalEmails++;
          const rohanBody = (() => {
            const extract = (part) => {
              if (!part) return '';
              if (part.mimeType==='text/plain' && part.body?.data)
                return Buffer.from(part.body.data.replace(/-/g,'+').replace(/_/g,'/'), 'base64').toString('utf-8');
              if (part.parts) { for (const p of part.parts) { const t=extract(p); if (t) return t; } }
              return '';
            };
            return extract(rohanMsg.payload);
          })();

          const rows = parseExpBody(rohanBody);

          // Get full total from Sandesh's email
          const sandeshMsg = (thread.data.messages||[]).find(m =>
            (m.payload?.headers||[]).find(h => h.name==='From' && h.value.toLowerCase().includes('sandesh'))
          );
          const sandeshBody = (() => {
            const extract = (part) => {
              if (!part) return '';
              if (part.mimeType==='text/plain' && part.body?.data)
                return Buffer.from(part.body.data.replace(/-/g,'+').replace(/_/g,'/'), 'base64').toString('utf-8');
              if (part.parts) { for (const p of part.parts) { const t=extract(p); if (t) return t; } }
              return '';
            };
            return sandeshMsg ? extract(sandeshMsg.payload) : '';
          })();
          const totalMatch = sandeshBody.match(/Total Amount\s*[:\s]+([\d,]+)\s*\/-?/i);
          const fullTotal = totalMatch ? parseInt(totalMatch[1].replace(/,/g,'')) : 0;

          // Scale rows to full total by ratio
          // For expenses report: include expenses + technology + salary (not finance)
          // For finance report: include only finance rows
          const filteredRows = type === 'finance'
            ? rows.filter(r => r.category === 'finance')
            : rows.filter(r => r.category !== 'finance');
          const rowTotal = rows.reduce((s,r) => s+r.amount, 0);

          if (rowTotal > 0 && fullTotal > 0) {
            for (const r of filteredRows) {
              const scaled = Math.round(r.amount * fullTotal / rowTotal);
              allRows.push({ name: r.name, amount: scaled, category: r.category });
            }
          } else {
            for (const r of filteredRows) allRows.push({ name: r.name, amount: r.amount, category: r.category });
          }
        } catch(e) {}
      }));
    }

    // Group by name
    const map = new Map();
    for (const r of allRows) {
      const key = r.name.trim().toLowerCase();
      if (!map.has(key)) map.set(key, { vendor: r.name, total: 0, count: 0 });
      const v = map.get(key);
      v.total += r.amount;
      v.count += 1;
    }

    const rows = [...map.values()].sort((a,b) => b.total-a.total);
    const grandTotal = rows.reduce((s,r) => s+r.total, 0);

    // For expenses report — also return rows grouped by category with subtotals
    let categories = null;
    if (type === 'expenses') {
      // Re-aggregate with category info from allRows
      const catMap = {};
      for (const r of allRows) {
        const cat = r.category || 'Expenses';
        const catLabel = cat === 'technology' ? 'Technology'
          : cat === 'salary' ? 'Salary'
          : 'Expenses';
        if (!catMap[catLabel]) catMap[catLabel] = new Map();
        const key = r.name.trim().toLowerCase();
        if (!catMap[catLabel].has(key)) catMap[catLabel].set(key, { vendor: r.name, total: 0, count: 0 });
        const v = catMap[catLabel].get(key);
        v.total += r.amount; v.count++;
      }
      // Sort each category by amount, sort categories by total
      categories = {};
      const catOrder = Object.entries(catMap)
        .map(([cat, m]) => ({ cat, total: [...m.values()].reduce((s,r)=>s+r.total,0) }))
        .sort((a,b) => b.total-a.total);
      for (const { cat } of catOrder) {
        categories[cat] = [...catMap[cat].values()].sort((a,b) => b.total-a.total);
      }
    }

    return res.json({ rows, grandTotal, totalEmails, month, type, categories });

  } catch(e) {
    return res.status(500).json({ error: e.message });
  }
}
