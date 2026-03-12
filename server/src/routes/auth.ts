import { Router } from 'express';
import * as outlook from '../services/outlook';
import * as teamleader from '../services/teamleader';

const router = Router();

const getBaseUrl = (req: any) => process.env.APP_URL || `${req.protocol}://${req.get('host')}`;

// Outlook OAuth
router.get('/outlook', async (req, res) => {
  try {
    const redirectUri = `${getBaseUrl(req)}/api/auth/outlook/callback`;
    const url = await outlook.getAuthUrl(redirectUri);
    res.redirect(url);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/outlook/callback', async (req, res) => {
  try {
    const redirectUri = `${getBaseUrl(req)}/api/auth/outlook/callback`;
    await outlook.handleCallback(req.query.code as string, redirectUri);
    res.redirect('/?connected=outlook');
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Teamleader OAuth
router.get('/teamleader', (req, res) => {
  const redirectUri = `${getBaseUrl(req)}/api/auth/teamleader/callback`;
  const url = teamleader.getAuthUrl(redirectUri);
  res.redirect(url);
});

router.get('/teamleader/callback', async (req, res) => {
  try {
    const redirectUri = `${getBaseUrl(req)}/api/auth/teamleader/callback`;
    await teamleader.handleCallback(req.query.code as string, redirectUri);
    res.redirect('/?connected=teamleader');
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Auth status
router.get('/status', async (_req, res) => {
  const outlookConnected = await outlook.isConnected();
  const teamleaderConnected = await teamleader.isConnected();
  res.json({
    outlook_connected: outlookConnected,
    teamleader_connected: teamleaderConnected,
  });
});

export default router;
