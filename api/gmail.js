import { google } from 'googleapis';

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

// Extract amount from subject line
function extractAmt(s) {
  s = s || '';
  const cr = s.match(/(\d+\.?\d*)\s*Cr/i);
  if (cr) return '₹' + cr[1] + ' Cr';
  const rs = s.match(/Rs\.?\s*([\d,]+)/i);
  if (rs) {
    const num = parseInt(rs[1].replace(/,/g, ''));
    if (num >= 10000000) return '₹' + (num/10000000).toFixed(2) + ' Cr';
    if (num >= 100000) return '₹' + (num/100000).toFixed(2) + ' L';
    return '₹' + rs[1];
  }
  const l = s.match(/(\d+\.?\d*)\s*L(acs?|akhs?)?/i);
  if (l) return '₹' + l[1] + ' L';
  return '';
}

// Clean subject line into readable title
function cleanTitle(s) {
  return (s || '')
    .replace(/^(Re:|RE:|Fwd:|FWD:)\s*/gi, '')
    .replace(/^APP-[A-Z]+-?\s*/gi, '')
    .replace(/^CRM-APP:\s*/gi, '')
    .replace(/^HR-APP:\s*/gi, '')
    .replace(/Rs\.?\s*[\d,]+\s*\/-?\s*/gi, '')
    .replace(/\s*-?\s*ASL\s*\d{2}-\d{2}-\d{4}\s*/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action } = req.query;

  // ── Auth URL ──────────────────────────────────────────────────────────────────
  if (action === 'auth-url') {
    const url = oauth2Client.generateAuthUrl({
      access_type: 'offline', prompt: 'consent',
      scope: [
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/gmail.modify'
      ]
    });
    return res.json({ url });
  }

  // ── Auth callback ─────────────────────────────────────────────────────────────
  if (action === 'auth-callback') {
    const { code } = req.query;
    try {
      const { tokens } = await oauth2Client.getToken(code);
      const tokenData = encodeURIComponent(JSON.stringify(tokens));
      res.setHeader('Location', `/?tokens=${tokenData}`);
      return res.status(302).end();
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── Requires tokens ───────────────────────────────────────────────────────────
  const { tokens } = req.method === 'POST' ? req.body : {};
  if (!tokens) return res.status(400).json({ error: 'No tokens provided' });

  oauth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

  // ── Fetch inbox ───────────────────────────────────────────────────────────────
  if (action === 'inbox') {
    try {
      // Fetch all label lists in parallel
      const [fundsRes, crmRes, hrRes, taxRes] = await Promise.all([
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds in:inbox', maxResults: 50 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:CRM in:inbox', maxResults: 30 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:HR in:inbox', maxResults: 20 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:Tax in:inbox', maxResults: 20 })
      ]);

      const allMessages = [
        ...(fundsRes.data.messages || []).map(m => ({ ...m, cat: 'Funds' })),
        ...(crmRes.data.messages || []).map(m => ({ ...m, cat: 'CRM' })),
        ...(hrRes.data.messages || []).map(m => ({ ...m, cat: 'HR' })),
        ...(taxRes.data.messages || []).map(m => ({ ...m, cat: 'Tax' }))
      ];

      // Deduplicate by threadId
      const threadMap = new Map();
      for (const msg of allMessages) {
        if (!threadMap.has(msg.threadId)) threadMap.set(msg.threadId, msg);
      }
      const msgList = [...threadMap.values()];

      // ── Step 1: Fetch metadata for ALL messages (fast, parallel) ─────────────
      const metaResults = await Promise.all(msgList.map(async (msg) => {
        try {
          const detail = await gmail.users.messages.get({
            userId: 'me', id: msg.id, format: 'metadata',
            metadataHeaders: ['From', 'To', 'Cc', 'Subject', 'Date']
          });
          const headers = {};
          for (const h of (detail.data.payload?.headers || [])) {
            headers[h.name.toLowerCase()] = h.value;
          }
          const labelIds = detail.data.labelIds || [];
          const isUnread = labelIds.includes('UNREAD');
          const to = (headers.to || '').toLowerCase();
          const isCc = !to.includes('ronak@aris.in') && !to.includes('ronak@arisinfra.one');
          let status;
          if (!isUnread) status = 'done';
          else if (isCc) status = 'fyi';
          else status = 'pending';

          return {
            msg,
            headers,
            status,
            isUnread,
            isCc,
            snippet: detail.data.snippet || ''
          };
        } catch(e) { return null; }
      }));

      const validMeta = metaResults.filter(Boolean);

      // ── Step 2: AI parse ONLY pending + fyi (max 15) ─────────────────────────
      const needsAI = validMeta.filter(m => m.status !== 'done').slice(0, 15);
      const doneItems = validMeta.filter(m => m.status === 'done');

      const aiParse = async (meta) => {
        const { msg, headers, status, isUnread, isCc, snippet } = meta;
        let title = cleanTitle(headers.subject || '');
        let amount = extractAmt(headers.subject || '');
        let risk = '';
        let fields = [];
        let rows = [];
        let note = '';

        try {
          // Fetch full body for AI
          const detail = await gmail.users.messages.get({
            userId: 'me', id: msg.id, format: 'full'
          });
          let bodyText = snippet;
          try {
            const extractBody = (part) => {
              if (!part) return '';
              if (part.mimeType === 'text/plain' && part.body?.data) {
                return Buffer.from(part.body.data, 'base64').toString('utf-8');
              }
              if (part.parts) {
                for (const p of part.parts) {
                  const t = extractBody(p);
                  if (t) return t;
                }
              }
              return '';
            };
            const extracted = extractBody(detail.data.payload);
            if (extracted) bodyText = extracted.slice(0, 3000);
          } catch(e) {}

          const aiRes = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'x-api-key': process.env.ANTHROPIC_API_KEY,
              'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
              model: 'claude-haiku-4-5-20251001',
              max_tokens: 500,
              system: `You process internal approval emails for ARIS. Extract structured data for the CMD approval dashboard.
Always use first names only. Return ONLY valid JSON:
{
  "title": "Short plain English title, max 8 words, no APP- prefix",
  "amount": "₹X Cr or ₹X L or empty string",
  "risk": "High or Medium or Low or empty (CRM only)",
  "fields": [{"label": "Short label max 3 words", "value": "value"}],
  "rows": [{"name": "vendor/person name", "category": "Payables or Salary or Expenses or Finance Cost", "amount": "₹X L"}],
  "note": "1-2 sentence context"
}
Fields: key facts (from, amount, account, party, credit limit etc). Rows: top 5 for expense/payment emails only.
Categories: Payables=vendor/material, Finance Cost=interest/bank/NCD/factoring, Salary=payroll/wages, Expenses=professional fees/operational.`,
              messages: [{
                role: 'user',
                content: `Label: ${msg.cat}\nStatus: ${status}\nFrom: ${headers.from}\nSubject: ${headers.subject}\nBody:\n${bodyText}`
              }]
            })
          });

          const aiData = await aiRes.json();
          const aiText = aiData.content?.[0]?.text || '';
          const s = aiText.indexOf('{');
          const e = aiText.lastIndexOf('}');
          if (s !== -1 && e !== -1) {
            const parsed = JSON.parse(aiText.slice(s, e + 1));
            title = parsed.title || title;
            amount = parsed.amount || amount;
            risk = parsed.risk || '';
            fields = parsed.fields || [];
            rows = parsed.rows || [];
            note = parsed.note || '';
          }

          // Detect approval pills
          const bodyLower = bodyText.toLowerCase();
          const rohanApproved = bodyLower.includes('rohan') && bodyLower.includes('approv');
          let approvalPill = '';
          if (status === 'fyi') {
            if (bodyLower.includes('transfer') && bodyLower.includes('approv')) approvalPill = 'Transfer Approved';
            else if ((bodyLower.includes('limit') || bodyLower.includes('credit')) && bodyLower.includes('approv')) approvalPill = 'Limit Approved';
            else if (bodyLower.includes('amount') && bodyLower.includes('approv')) approvalPill = 'Amount Approved';
            else if ((bodyLower.includes('nishita') || bodyLower.includes('divya')) && bodyLower.includes('approv')) approvalPill = 'Approved';
          }

          return { tid: msg.threadId, mid: msg.id, cat: msg.cat, subj: headers.subject || '', title, amount, risk, fields, rows, note, from: headers.from || '', date: headers.date || '', to: headers.to || '', cc: headers.cc || '', isUnread, isCc, status, rohanApproved, approvalPill };

        } catch(e) {
          // AI failed — fallback to snippet
          const snip = snippet.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&#39;/g,"'").replace(/&quot;/g,'"');
          note = snip || 'Email received.';
          fields = [{ label: 'From', value: (headers.from || '').replace(/<[^>]+>/g,'').trim() }];
          return { tid: msg.threadId, mid: msg.id, cat: msg.cat, subj: headers.subject || '', title, amount, risk, fields, rows, note, from: headers.from || '', date: headers.date || '', to: headers.to || '', cc: headers.cc || '', isUnread, isCc, status, rohanApproved: false, approvalPill: '' };
        }
      };

      // ── Step 3: Build done items without AI (just subject + amount) ───────────
      const doneFormatted = doneItems.map(({ msg, headers, status, isUnread, isCc }) => ({
        tid: msg.threadId,
        mid: msg.id,
        cat: msg.cat,
        subj: headers.subject || '',
        title: cleanTitle(headers.subject || ''),
        amount: extractAmt(headers.subject || ''),
        risk: '',
        fields: [],
        rows: [],
        note: '',
        from: headers.from || '',
        date: headers.date || '',
        to: headers.to || '',
        cc: headers.cc || '',
        isUnread,
        isCc,
        status,
        rohanApproved: false,
        approvalPill: ''
      }));

      // Run AI on pending/fyi in parallel
      const aiResults = await Promise.all(needsAI.map(aiParse));
      const items = [...aiResults.filter(Boolean), ...doneFormatted];

      return res.json({ items, refreshedTokens: oauth2Client.credentials });
    } catch(e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── Mark as read ──────────────────────────────────────────────────────────────
  if (action === 'mark-read') {
    const { tid } = req.body;
    try {
      await gmail.users.threads.modify({
        userId: 'me', id: tid,
        requestBody: { removeLabelIds: ['UNREAD'] }
      });
      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch(e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── Send reply ────────────────────────────────────────────────────────────────
  if (action === 'send') {
    const { tid, to, cc, subject, body } = req.body;
    try {
      const thread = await gmail.users.threads.get({ userId: 'me', id: tid, format: 'metadata' });
      const msgs = thread.data.messages || [];
      const lastMsg = msgs[msgs.length - 1];
      const lastHeaders = {};
      for (const h of (lastMsg?.payload?.headers || [])) {
        lastHeaders[h.name.toLowerCase()] = h.value;
      }
      const messageId = lastHeaders['message-id'] || '';
      const replySubject = subject.startsWith('Re:') ? subject : `Re: ${subject}`;
      const emailLines = [
        `From: Ronak Morbia <ronak@aris.in>`,
        `To: ${to}`,
        cc ? `Cc: ${cc}` : '',
        `Subject: ${replySubject}`,
        messageId ? `In-Reply-To: ${messageId}` : '',
        messageId ? `References: ${messageId}` : '',
        `Content-Type: text/plain; charset=utf-8`,
        '',
        body
      ].filter(l => l !== '');

      const raw = Buffer.from(emailLines.join('\r\n'))
        .toString('base64')
        .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

      await gmail.users.messages.send({
        userId: 'me',
        requestBody: { raw, threadId: tid }
      });
      await gmail.users.threads.modify({
        userId: 'me', id: tid,
        requestBody: { removeLabelIds: ['UNREAD'] }
      });

      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch(e) {
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(400).json({ error: 'Unknown action' });
}
