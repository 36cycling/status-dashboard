import type { Customer, AuthStatus } from '../../shared/types';

const BASE = '/api';

async function request<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    credentials: 'include',
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...options?.headers,
    },
  });
  if (res.status === 401) {
    throw new Error('UNAUTHORIZED');
  }
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.error || `Request failed: ${res.status}`);
  }
  return res.json();
}

export function getCustomers(): Promise<Customer[]> {
  return request('/customers');
}

export function getArchivedCustomers(): Promise<Customer[]> {
  return request('/customers/archived');
}

export function archiveCustomer(id: number): Promise<{ success: boolean }> {
  return request(`/customers/${id}/archive`, { method: 'POST' });
}

export function unarchiveCustomer(id: number): Promise<{ success: boolean }> {
  return request(`/customers/${id}/unarchive`, { method: 'POST' });
}

export function syncNow(): Promise<{ success: boolean; mails_synced: number; last_sync: string }> {
  return request('/sync', { method: 'POST' });
}

export function getSyncStatus(): Promise<{ last_outlook_sync: string | null; last_teamleader_sync: string | null }> {
  return request('/sync/status');
}

export function getAuthStatus(): Promise<AuthStatus> {
  return request('/auth/status');
}

export function login(password: string): Promise<{ success: boolean }> {
  return request('/login', {
    method: 'POST',
    body: JSON.stringify({ password }),
  });
}
