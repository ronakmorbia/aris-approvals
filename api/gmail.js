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

  // ── Auth: get Google OAuth URL ──────────────────────────────────────────
  if (action === 'auth-url') {
    const url = oauth2Client.generateAuthUrl({
      access_type: 'offline',
      prompt: 'consent',
      scope: [
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/gmail.modify'
      ]
    });
    return res.json({ url });
  }

  // ── Auth: exchange code for tokens ──────────────────────────────────────
  if (action === 'auth-callback') {
    const { code } = req.query;
    try {
      const { tokens } = await oauth2Client.getToken(code);
      // Redirect to homepage with tokens as URL fragment (never sent to server)
      const tokenData = encodeURIComponent(JSON.stringify(tokens));
      res.setHeader('Location', `/?tokens=${tokenData}`);
      return res.status(302).end();
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── Requires tokens from here ────────────────────────────────────────────
  const { tokens } = req.method === 'POST' ? req.body : {};
  if (!tokens) return res.status(400).json({ error: 'No tokens provided' });

  oauth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

  // ── Fetch inbox ──────────────────────────────────────────────────────────
  if (action === 'inbox') {
    try {
      const [fundsRes, crmRes, hrRes, taxRes] = await Promise.all([
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds in:inbox', maxResults: 50 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:CRM in:inbox', maxResults: 30 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:HR in:inbox', maxResults: 30 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:Tax in:inbox', maxResults: 30 })
      ]);

      const allMessages = [
        ...(fundsRes.data.messages || []).map(m => ({ ...m, cat: 'Funds' })),
        ...(crmRes.data.messages || []).map(m => ({ ...m, cat: 'CRM' })),
        ...(hrRes.data.messages || []).map(m => ({ ...m, cat: 'HR' })),
        ...(taxRes.data.messages || []).map(m => ({ ...m, cat: 'Tax' }))
      ];

      // Deduplicate by threadId — keep latest message per thread
      const threadMap = new Map();
      for (const msg of allMessages) {
        if (!threadMap.has(msg.threadId)) {
          threadMap.set(msg.threadId, msg);
        }
      }

      const items = [];
      for (const msg of threadMap.values()) {
        try {
          // Fetch full message to get body text
          const detail = await gmail.users.messages.get({
            userId: 'me',
            id: msg.id,
            format: 'full',
            metadataHeaders: ['From', 'To', 'Cc', 'Subject', 'Date']
          });

          const headers = {};
          for (const h of detail.data.payload.headers) {
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
            if (extracted) bodyText = extracted.slice(0, 2000);
          } catch(e) {}

          // Generate smart card content via Anthropic API
          let title = headers.subject || '';
          let amount = '';
          let summary = '';

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
                max_tokens: 300,
                system: `You are an assistant that processes approval request emails for ARIS, a construction materials company. 
Extract key information and return ONLY valid JSON with no other text:
{
  "title": "Short action-oriented title (max 8 words, no APP- prefix, e.g. 'Vendor payment — Arisinfra' or 'Customer creation — Casa Grande Axiom')",
  "amount": "Amount if mentioned (e.g. '₹2.06 Cr' or '₹30 L') or empty string",
  "summary": "2-3 sentence summary. For pending: what is being requested, by whom, key details. For FYI/done: what happened, who requested it, who approved it."
}`,
                messages: [{
                  role: 'user',
                  content: `Category: ${msg.cat}\nStatus: ${status}\nSubject: ${headers.subject}\nFrom: ${headers.from}\nBody:\n${bodyText}`
                }]
              })
            });
            const aiData = await aiRes.json();
            const aiText = aiData.content?.[0]?.text || '';
            const s = aiText.indexOf('{'), e = aiText.lastIndexOf('}');
            if (s !== -1 && e !== -1) {
              const parsed = JSON.parse(aiText.slice(s, e + 1));
              title = parsed.title || title;
              amount = parsed.amount || '';
              summary = parsed.summary || '';
            }
          } catch(e) {
            // Fall back to subject and snippet
            summary = detail.data.snippet || '';
          }

          items.push({
            tid: msg.threadId,
            mid: msg.id,
            cat: msg.cat,
            subj: headers.subject || '',
            title,
            amount,
            summary,
            from: headers.from || '',
            date: headers.date || '',
            to: headers.to || '',
            cc: headers.cc || '',
            isUnread,
            isCc,
            status
          });
        } catch (e) { /* skip */ }
      }

      return res.json({ items, refreshedTokens: oauth2Client.credentials });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── Mark thread as read ──────────────────────────────────────────────────
  if (action === 'mark-read') {
    const { tid } = req.body;
    try {
      await gmail.users.threads.modify({
        userId: 'me',
        id: tid,
        requestBody: { removeLabelIds: ['UNREAD'] }
      });
      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // ── Send reply ───────────────────────────────────────────────────────────
  if (action === 'send') {
    const { tid, to, cc, subject, body } = req.body;
    try {
      // Get thread to find last message ID for In-Reply-To header
      const thread = await gmail.users.threads.get({ userId: 'me', id: tid, format: 'metadata' });
      const msgs = thread.data.messages;
      const lastMsg = msgs[msgs.length - 1];
      const lastMsgId = lastMsg.payload.headers.find(h => h.name === 'Message-Id')?.value || '';

      const replySubject = subject.startsWith('Re:') ? subject : 'Re: ' + subject;
      const emailLines = [
        `From: Ronak Morbia <ronak@aris.in>`,
        `To: ${to}`,
        cc ? `Cc: ${cc}` : '',
        `Subject: ${replySubject}`,
        `In-Reply-To: ${lastMsgId}`,
        `References: ${lastMsgId}`,
        `Content-Type: text/plain; charset=utf-8`,
        ``,
        body
      ].filter(l => l !== null && l !== undefined);

      const raw = Buffer.from(emailLines.join('\r\n'))
        .toString('base64')
        .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

      await gmail.users.messages.send({
        userId: 'me',
        requestBody: { raw, threadId: tid }
      });

      // Mark thread as read after sending
      await gmail.users.threads.modify({
        userId: 'me',
        id: tid,
        requestBody: { removeLabelIds: ['UNREAD'] }
      });

      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(400).json({ error: 'Unknown action' });
}
