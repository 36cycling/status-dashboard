import { Router } from 'express';
import { getAll, getOne, runQuery } from '../db';
import { syncMails } from '../services/outlook';
import { syncTeamleaderForCustomers } from '../services/teamleader';

const router = Router();

router.get('/customers', (_req, res) => {
  const customers = getAll(`
    SELECT c.*, MAX(e.date) as last_activity
    FROM customers c
    LEFT JOIN timeline_events e ON e.customer_id = c.id
    WHERE c.archived = 0
    GROUP BY c.id
    ORDER BY last_activity DESC NULLS LAST
  `);

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
  runQuery('UPDATE customers SET archived = 1 WHERE id = ?', [Number(req.params.id)]);
  res.json({ success: true });
});

router.post('/customers/:id/unarchive', (req, res) => {
  runQuery('UPDATE customers SET archived = 0 WHERE id = ?', [Number(req.params.id)]);
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

// Reset database — clears all customers and events for a fresh sync
router.post('/reset', (_req, res) => {
  try {
    runQuery('DELETE FROM timeline_events', []);
    runQuery('DELETE FROM customers', []);
    runQuery("DELETE FROM sync_state WHERE key IN ('last_outlook_sync', 'last_teamleader_sync')", []);
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

export default router;
