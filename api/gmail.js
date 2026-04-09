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
      const [fundsRes, crmRes, sentRes] = await Promise.all([
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds in:inbox', maxResults: 50 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:CRM in:inbox', maxResults: 30 }),
        gmail.users.messages.list({ userId: 'me', q: 'label:Funds from:ronak@aris.in', maxResults: 50 })
      ]);

      const approvedThreadIds = new Set(
        (sentRes.data.messages || []).map(m => m.threadId)
      );

      const allMessages = [
        ...(fundsRes.data.messages || []).map(m => ({ ...m, cat: 'Funds' })),
        ...(crmRes.data.messages || []).map(m => ({ ...m, cat: 'CRM' }))
      ];

      // Deduplicate by threadId, fetch headers
      const seen = new Set();
      const items = [];
      for (const msg of allMessages) {
        if (seen.has(msg.threadId)) continue;
        seen.add(msg.threadId);
        try {
          const detail = await gmail.users.messages.get({
            userId: 'me',
            id: msg.id,
            format: 'metadata',
            metadataHeaders: ['From', 'To', 'Cc', 'Subject', 'Date']
          });
          const headers = {};
          for (const h of detail.data.payload.headers) {
            headers[h.name.toLowerCase()] = h.value;
          }
          const to = (headers.to || '').toLowerCase();
          const isCc = !to.includes('ronak@aris.in') && !to.includes('ronak@arisinfra.one');
          const approved = approvedThreadIds.has(msg.threadId);
          items.push({
            tid: msg.threadId,
            mid: msg.id,
            cat: msg.cat,
            subj: headers.subject || '',
            from: headers.from || '',
            date: headers.date || '',
            to: headers.to || '',
            cc: headers.cc || '',
            snippet: detail.data.snippet || '',
            isCc,
            status: approved ? 'done' : isCc ? 'fyi' : 'pending'
          });
        } catch (e) { /* skip */ }
      }

      return res.json({ items, refreshedTokens: oauth2Client.credentials });
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

      return res.json({ ok: true, refreshedTokens: oauth2Client.credentials });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(400).json({ error: 'Unknown action' });
}
