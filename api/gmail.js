import { google } from 'googleapis';

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action } = req.query;

  // ── Auth URL ─────────────────────────────────────────────────────────────────
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
      const [fundsRes, crmRes, hrRes, taxRes] = await Promise.all([
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds in:inbox', maxResults: 20 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:CRM in:inbox', maxResults: 15 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:HR in:inbox', maxResults: 15 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:Tax in:inbox', maxResults: 10 })
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

      // Cap at 20 most recent to stay within timeout
      const msgList = [...threadMap.values()].slice(0, 20);

      const processMsg = async (msg) => {
        try {
          const detail = await gmail.users.messages.get({
            userId: 'me', id: msg.id, format: 'full'
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

          // Extract plain text body
          let bodyText = detail.data.snippet || '';
          try {
            const extractBody = (part) => {
              if (!part) return '';
              if (part.mimeType === 'text/plain' && part.body?.data) {
                return Buffer.from(part.body.data, 'base64').toString('utf-8');
              }
              if (part.parts) {
                for (const p of part.parts) {
                  const text = extractBody(p);
                  if (text) return text;
                }
              }
              return '';
            };
            const extracted = extractBody(detail.data.payload);
            if (extracted) bodyText = extracted.slice(0, 3000);
          } catch(e) {}

          // AI smart card parsing
          let title = (headers.subject || '').replace(/^(Re:|RE:|APP-[A-Z]+-?\s*)/gi, '').trim();
          let amount = '';
          let risk = '';
          let fields = [];
          let rows = [];
          let note = '';

          try {
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
Always use first names only. Return ONLY valid JSON in this exact format:
{
  "title": "Short plain English title, max 8 words, no APP- prefix",
  "amount": "₹X Cr or ₹X L or empty string",
  "risk": "High or Medium or Low or empty string (CRM only)",
  "fields": [{"label": "Short label", "value": "value"}],
  "rows": [{"name": "vendor/person", "category": "Payables or Salary or Expenses or Finance Cost", "amount": "₹X L"}],
  "note": "1-2 sentence context"
}
Fields: key facts for the card (from, amount, account, party etc).
Rows: only for expense/payment emails, top 5 by amount.
Categories: Payables=vendor/material payments, Finance Cost=interest/bank charges/NCD/factoring, Salary=payroll/wages, Expenses=professional fees/operational.`,
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
              amount = parsed.amount || '';
              risk = parsed.risk || '';
              fields = parsed.fields || [];
              rows = parsed.rows || [];
              note = parsed.note || '';
            }
          } catch(e) {
            // AI failed — use snippet as fallback so card is never blank
            const snip = (detail.data.snippet || '')
              .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&#39;/g,"'").replace(/&quot;/g,'"');
            note = snip || 'Email received.';
            fields = [
              { label: 'From', value: (headers.from || '').replace(/<[^>]+>/g,'').trim() },
              { label: 'Date', value: headers.date || '' }
            ];
          }

          // Detect Rohan pre-approval and approvalPill
          const bodyLower = bodyText.toLowerCase();
          const rohanApproved = bodyLower.includes('rohan') && (bodyLower.includes('approved') || bodyLower.includes('approve'));
          let approvalPill = '';
          if (status === 'fyi') {
            if (bodyLower.includes('transfer') && bodyLower.includes('approv')) approvalPill = 'Transfer Approved';
            else if ((bodyLower.includes('limit') || bodyLower.includes('credit')) && bodyLower.includes('approv')) approvalPill = 'Limit Approved';
            else if (bodyLower.includes('amount') && bodyLower.includes('approv')) approvalPill = 'Amount Approved';
            else if ((bodyLower.includes('nishita') || bodyLower.includes('divya')) && bodyLower.includes('approv')) approvalPill = 'Approved';
          }

          return {
            tid: msg.threadId,
            mid: msg.id,
            cat: msg.cat,
            subj: headers.subject || '',
            title,
            amount,
            risk,
            fields,
            rows,
            note,
            from: headers.from || '',
            date: headers.date || '',
            to: headers.to || '',
            cc: headers.cc || '',
            isUnread,
            isCc,
            status,
            rohanApproved,
            approvalPill
          };
        } catch(e) {
          return null;
        }
      };

      // Process all messages in parallel
      const results = await Promise.all(msgList.map(processMsg));
      const items = results.filter(Boolean);

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
      const lastMsgId = lastMsg?.id || '';
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
        ``,
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
