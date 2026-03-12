import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import session from 'express-session';
import path from 'path';
import { initDb } from './db';
import apiRouter from './routes/api';
import authRouter from './routes/auth';

async function start() {
  await initDb();

  const app = express();
  const PORT = parseInt(process.env.PORT || '3001', 10);

  app.use(cors({ origin: true, credentials: true }));
  app.use(express.json());
  app.use(
    session({
      secret: process.env.SESSION_SECRET || 'dev-secret-change-me',
      resave: false,
      saveUninitialized: false,
      cookie: { secure: false, maxAge: 24 * 60 * 60 * 1000 },
    })
  );

  // Simple password auth middleware
  const appPassword = process.env.APP_PASSWORD;
  if (appPassword) {
    app.use((req, res, next) => {
      if (req.path.startsWith('/api/auth/') && req.path.includes('/callback')) {
        return next();
      }
      if ((req.session as any).authenticated) {
        return next();
      }
      if (req.path === '/api/login' && req.method === 'POST') {
        if (req.body.password === appPassword) {
          (req.session as any).authenticated = true;
          return res.json({ success: true });
        }
        return res.status(401).json({ error: 'Onjuist wachtwoord' });
      }
      if (req.path === '/api/login' || !req.path.startsWith('/api/')) {
        return next();
      }
      res.status(401).json({ error: 'Niet ingelogd' });
    });
  }

  app.use('/api', apiRouter);
  app.use('/api/auth', authRouter);

  // Serve React frontend in production
  const clientDist = path.join(__dirname, '..', 'public');
  app.use(express.static(clientDist));
  app.get('*', (_req, res) => {
    res.sendFile(path.join(clientDist, 'index.html'));
  });

  app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

start().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
