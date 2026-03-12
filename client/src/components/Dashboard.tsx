import { useState, useEffect, useCallback } from 'react';
import type { Customer } from '../../../shared/types';
import type { AuthStatus } from '../../../shared/types';
import { getCustomers, archiveCustomer, syncNow, getAuthStatus, getSyncStatus } from '../api';
import CustomerRow from './CustomerRow';

export default function Dashboard() {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [authStatus, setAuthStatus] = useState<AuthStatus | null>(null);
  const [lastSync, setLastSync] = useState<string | null>(null);
  const [syncing, setSyncing] = useState(false);
  const [error, setError] = useState('');

  const loadData = useCallback(async () => {
    try {
      const [custs, auth, sync] = await Promise.all([
        getCustomers(),
        getAuthStatus(),
        getSyncStatus(),
      ]);
      setCustomers(custs);
      setAuthStatus(auth);
      setLastSync(sync.last_outlook_sync);
      setError('');
    } catch (err: any) {
      if (err.message === 'UNAUTHORIZED') throw err;
      setError(err.message);
    }
  }, []);

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 60000); // refresh every minute
    return () => clearInterval(interval);
  }, [loadData]);

  const handleSync = async () => {
    setSyncing(true);
    try {
      await syncNow();
      await loadData();
    } catch (err: any) {
      setError(err.message);
    } finally {
      setSyncing(false);
    }
  };

  const handleArchive = async (id: number) => {
    try {
      await archiveCustomer(id);
      setCustomers((prev) => prev.filter((c) => c.id !== id));
    } catch (err: any) {
      setError(err.message);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4">
        <div className="flex items-center justify-between">
          <h1 className="text-2xl font-bold text-slate-800">Status dashboard</h1>
          <div className="flex items-center gap-4">
            {/* Connection status indicators */}
            <div className="flex items-center gap-2 text-xs">
              <span
                className={`inline-block w-2 h-2 rounded-full ${
                  authStatus?.outlook_connected ? 'bg-green-500' : 'bg-red-400'
                }`}
              />
              <span className="text-slate-600">Outlook</span>
              {!authStatus?.outlook_connected && (
                <a href="/api/auth/outlook" className="text-blue-600 hover:underline">
                  Verbinden
                </a>
              )}
            </div>
            <div className="flex items-center gap-2 text-xs">
              <span
                className={`inline-block w-2 h-2 rounded-full ${
                  authStatus?.teamleader_connected ? 'bg-green-500' : 'bg-red-400'
                }`}
              />
              <span className="text-slate-600">Teamleader</span>
              {!authStatus?.teamleader_connected && (
                <a href="/api/auth/teamleader" className="text-blue-600 hover:underline">
                  Verbinden
                </a>
              )}
            </div>

            {/* Sync button */}
            <button
              onClick={handleSync}
              disabled={syncing}
              className="px-4 py-1.5 bg-blue-600 text-white text-sm rounded-md hover:bg-blue-700 disabled:opacity-50"
            >
              {syncing ? 'Synchroniseren...' : 'Sync nu'}
            </button>

            {lastSync && (
              <span className="text-xs text-slate-400">
                Laatste sync: {new Date(lastSync).toLocaleString('nl-NL')}
              </span>
            )}
          </div>
        </div>
      </header>

      {/* Error banner */}
      {error && (
        <div className="mx-6 mt-4 p-3 bg-red-50 border border-red-200 rounded-md text-red-700 text-sm">
          {error}
          <button onClick={() => setError('')} className="ml-2 text-red-500 hover:text-red-700">
            Sluiten
          </button>
        </div>
      )}

      {/* Legend */}
      <div className="px-6 py-3 flex items-center gap-4 text-xs text-slate-500">
        <div className="flex items-center gap-1.5">
          <div className="w-4 h-3 bg-white border-2 border-blue-500 rounded-sm" />
          <span>Ontvangen (beantwoord)</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-4 h-3 bg-white border-2 border-red-500 rounded-sm" />
          <span>Ontvangen (onbeantwoord)</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-4 h-3 bg-blue-500 rounded-sm" />
          <span>Verstuurd antwoord</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-4 h-3 bg-teal-500 rounded-sm" />
          <span>Teamleader (contact/deal)</span>
        </div>
      </div>

      {/* Customer list */}
      <main className="px-6 py-4">
        {customers.length === 0 && !error && (
          <div className="text-center text-slate-400 py-12">
            <p className="text-lg mb-2">Nog geen klantaanvragen</p>
            <p className="text-sm">
              Verbind je Outlook account en synchroniseer om te beginnen.
            </p>
          </div>
        )}

        {customers.map((customer) => (
          <CustomerRow
            key={customer.id}
            customer={customer}
            onArchive={handleArchive}
          />
        ))}
      </main>
    </div>
  );
}
