import { useState, useEffect } from 'react';
import Dashboard from './components/Dashboard';
import LoginPage from './components/LoginPage';
import { getAuthStatus } from './api';

export default function App() {
  const [authenticated, setAuthenticated] = useState<boolean | null>(null);

  useEffect(() => {
    // Check if we're already authenticated by trying an API call
    getAuthStatus()
      .then(() => setAuthenticated(true))
      .catch((err) => {
        if (err.message === 'UNAUTHORIZED') {
          setAuthenticated(false);
        } else {
          // No password protection or other error
          setAuthenticated(true);
        }
      });
  }, []);

  if (authenticated === null) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-100">
        <div className="text-slate-400">Laden...</div>
      </div>
    );
  }

  if (!authenticated) {
    return <LoginPage onLogin={() => setAuthenticated(true)} />;
  }

  return <Dashboard />;
}
