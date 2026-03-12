import { Router } from 'express';
import { getAll, getOne, runQuery } from '../db';
import { syncMails } from '../services/outlook';
import { syncTeamleaderForCustomers } from '../services/teamleader';

const router = Router();

router.get('/customers', (_req, res) => {
  const customers = getAll('SELECT * FROM customers WHERE archived = 0 ORDER BY created_at DESC');

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

router.get('/sync/status', (_req, res) => {
  const lastOutlook = getOne("SELECT value FROM sync_state WHERE key = 'last_outlook_sync'");
  const lastTeamleader = getOne("SELECT value FROM sync_state WHERE key = 'last_teamleader_sync'");

  res.json({
    last_outlook_sync: lastOutlook?.value || null,
    last_teamleader_sync: lastTeamleader?.value || null,
  });
});

export default router;
