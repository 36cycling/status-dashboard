import { Router } from 'express';
import { getAll, getOne, runQuery } from '../db';
import { syncMails } from '../services/outlook';
import { syncTeamleaderForCustomers } from '../services/teamleader';

const router = Router();

router.get('/customers', (_req, res) => {
  // Get all customers that have events after their dismissed_at (or were never dismissed)
  const customers = getAll(`
    SELECT c.*, MAX(e.date) as last_activity
    FROM customers c
    INNER JOIN timeline_events e ON e.customer_id = c.id
    WHERE (c.dismissed_at IS NULL OR e.date > c.dismissed_at)
    GROUP BY c.id
    ORDER BY last_activity DESC NULLS LAST
  `);

  const result = customers.map((c) => ({
    ...c,
    archived: Boolean(c.archived),
    // Only show events after dismissed_at
    events: getAll(
      'SELECT * FROM timeline_events WHERE customer_id = ? AND (? IS NULL OR date > ?) ORDER BY date DESC',
      [c.id, c.dismissed_at, c.dismissed_at]
    ).map((e) => ({
      ...e,
      is_replied: Boolean(e.is_replied),
      metadata: JSON.parse((e.metadata as string) || '{}'),
    })),
  }));

  res.json(result);
});

router.get('/customers/archived', (_req, res) => {
  const customers = getAll('SELECT * FROM customers WHERE archived = 1 ORDER BY created_at DESC');

  const result = customers.map((c) => ({
    ...c,
    archived: Boolean(c.archived),
    events: getAll('SELECT * FROM timeline_events WHERE customer_id = ? ORDER BY date DESC', [c.id]).map((e) => ({
      ...e,
      is_replied: Boolean(e.is_replied),
      metadata: JSON.parse((e.metadata as string) || '{}'),
    })),
  }));

  res.json(result);
});

router.post('/customers/:id/archive', (req, res) => {
  // Use ISO format so it's comparable with event dates from Graph API
  runQuery('UPDATE customers SET dismissed_at = ? WHERE id = ?', [new Date().toISOString(), Number(req.params.id)]);
  res.json({ success: true });
});

router.post('/customers/:id/unarchive', (req, res) => {
  runQuery('UPDATE customers SET dismissed_at = NULL WHERE id = ?', [Number(req.params.id)]);
  res.json({ success: true });
});

router.post('/sync', async (_req, res) => {
  try {
    const mailResult = await syncMails();
    await syncTeamleaderForCustomers();

    const lastSync = getOne("SELECT value FROM sync_state WHERE key = 'last_outlook_sync'");

    res.json({
      success: true,
      mails_synced: mailResult.synced,
      last_sync: lastSync?.value || null,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/debug/folders', async (_req, res) => {
  try {
    const msal = await import('@azure/msal-node');
    const clientId = process.env.AZURE_CLIENT_ID!;
    const clientSecret = process.env.AZURE_CLIENT_SECRET!;
    const tenantId = process.env.AZURE_TENANT_ID!;
    const mailbox = process.env.OUTLOOK_MAILBOX || 'NOT SET';

    const client = new msal.ConfidentialClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret,
      },
    });

    const tokenResult = await client.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });

    if (!tokenResult) {
      return res.status(500).json({ error: 'Failed to get token', mailbox });
    }

    // Get top-level folders only
    const foldersRes = await fetch(`https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders?$top=100`, {
      headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
    });

    if (!foldersRes.ok) {
      const body = await foldersRes.text();
      return res.status(500).json({ error: `Graph: ${foldersRes.status}`, body, mailbox });
    }

    const foldersData: any = await foldersRes.json();
    const folderNames = foldersData.value.map((f: any) => ({ name: f.displayName, id: f.id, childCount: f.childFolderCount }));

    res.json({ mailbox, folders: folderNames });
  } catch (err: any) {
    res.status(500).json({ error: err.message, stack: err.stack?.substring(0, 500) });
  }
});

// Reset database — clears events and sync state, but preserves dismissed_at
router.post('/reset', (_req, res) => {
  try {
    // Save dismissed customers before reset
    const dismissed = getAll('SELECT email, dismissed_at FROM customers WHERE dismissed_at IS NOT NULL');
    runQuery('DELETE FROM timeline_events', []);
    runQuery('DELETE FROM customers', []);
    runQuery("DELETE FROM sync_state WHERE key IN ('last_outlook_sync', 'last_teamleader_sync')", []);
    // Restore dismissed_at for known emails
    for (const d of dismissed) {
      runQuery('INSERT OR IGNORE INTO customers (name, email, dismissed_at) VALUES (?, ?, ?)', ['', d.email, d.dismissed_at]);
    }
    res.json({ success: true, message: 'Database reset. Click Sync to re-import.' });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/sync/status', (_req, res) => {
  const lastOutlook = getOne("SELECT value FROM sync_state WHERE key = 'last_outlook_sync'");
  const lastTeamleader = getOne("SELECT value FROM sync_state WHERE key = 'last_teamleader_sync'");

  res.json({
    last_outlook_sync: lastOutlook?.value || null,
    last_teamleader_sync: lastTeamleader?.value || null,
  });
});

// Debug: show dismissed customers and their dates
router.get('/debug/dismissed', (_req, res) => {
  const dismissed = getAll('SELECT id, name, email, dismissed_at FROM customers WHERE dismissed_at IS NOT NULL');
  const sampleEvent = getOne("SELECT date FROM timeline_events ORDER BY date DESC LIMIT 1");
  res.json({
    count: dismissed.length,
    dismissed,
    sample_event_date_format: sampleEvent?.date || 'no events',
    note: 'dismissed_at and event dates must be in same format for comparison to work',
  });
});

router.get('/debug/reply-check', async (req, res) => {
  try {
    const email = req.query.email as string;
    if (!email) return res.json({ error: 'Pass ?email=someone@example.com' });

    const customer = getOne('SELECT * FROM customers WHERE email = ?', [email]);
    if (!customer) return res.json({ error: 'Customer not found' });

    // Get customer's email_in events with conversationId
    const events = getAll(
      "SELECT id, subject, date, is_replied, outlook_message_id FROM timeline_events WHERE customer_id = ? AND type = 'email_in' ORDER BY date DESC",
      [customer.id]
    );

    // Check sent items from all mailboxes
    const msal = await import('@azure/msal-node');
    const client = new msal.ConfidentialClientApplication({
      auth: {
        clientId: process.env.AZURE_CLIENT_ID!,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET!,
      },
    });
    const tokenResult = await client.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });
    if (!tokenResult) return res.status(500).json({ error: 'No token' });
    const token = tokenResult.accessToken;

    const REPLY_MAILBOXES = ['info@36cycling.com', 'jeroen@36cycling.com', 'lisette@36cycling.com', 'michael@36cycling.com', 'lars@36cycling.com'];
    const mailbox = process.env.OUTLOOK_MAILBOX || 'info@36cycling.com';

    // Get the original message details (conversationId)
    const results: any[] = [];
    for (const ev of events) {
      let msgDetail: any = null;
      try {
        const msgRes = await fetch(`https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${ev.outlook_message_id}?$select=conversationId,subject,from`, {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (msgRes.ok) msgDetail = await msgRes.json();
      } catch {}

      // Search for replies in all mailboxes
      const repliesFound: any[] = [];
      if (msgDetail?.conversationId) {
        for (const mb of REPLY_MAILBOXES) {
          try {
            const filter = `conversationId eq '${msgDetail.conversationId}'`;
            const sentRes = await fetch(`https://graph.microsoft.com/v1.0/users/${mb}/mailFolders('SentItems')/messages?$filter=${encodeURIComponent(filter)}&$select=subject,sentDateTime,from&$top=5`, {
              headers: { Authorization: `Bearer ${token}` },
            });
            if (sentRes.ok) {
              const sentData: any = await sentRes.json();
              for (const s of sentData.value) {
                repliesFound.push({ mailbox: mb, subject: s.subject, sentAt: s.sentDateTime, from: s.from?.emailAddress?.name });
              }
            }
          } catch {}
        }
      }

      results.push({
        eventId: ev.id,
        subject: ev.subject,
        date: ev.date,
        is_replied: Boolean(ev.is_replied),
        conversationId: msgDetail?.conversationId || 'NOT_FOUND',
        repliesFound,
      });
    }

    res.json({ customer: { name: customer.name, email: customer.email }, results });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/debug/events', (_req, res) => {
  const events = getAll("SELECT id, type, metadata FROM timeline_events WHERE type IN ('tl_contact', 'tl_deal', 'email_out') ORDER BY id DESC LIMIT 20");
  res.json(events.map(e => ({ ...e, metadata: JSON.parse((e.metadata as string) || '{}') })));
});

router.get('/debug/teamleader-raw', async (_req, res) => {
  try {
    const { findContact, findDeals } = await import('../services/teamleader');
    // Pick a customer that has a tl_contact event
    const event = getOne("SELECT customer_id, metadata FROM timeline_events WHERE type = 'tl_contact' LIMIT 1");
    if (!event) return res.json({ error: 'No tl_contact events found' });
    const customer = getOne('SELECT * FROM customers WHERE id = ?', [event.customer_id]);
    if (!customer) return res.json({ error: 'Customer not found' });

    const contact = await findContact(customer.email as string);
    let deals: any[] = [];
    if (contact) {
      deals = await findDeals(contact.id);
    }
    res.json({ deals });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

export default router;
