import { runQuery, getOne, getAll, saveDb, getDb } from '../db';

const TL_AUTH_URL = 'https://focus.teamleader.eu/oauth2/authorize';
const TL_TOKEN_URL = 'https://focus.teamleader.eu/oauth2/access_token';
const TL_API_BASE = 'https://api.focus.teamleader.eu';

export function getAuthUrl(redirectUri: string): string {
  const params = new URLSearchParams({
    client_id: process.env.TL_CLIENT_ID || '',
    redirect_uri: redirectUri,
    response_type: 'code',
  });
  return `${TL_AUTH_URL}?${params.toString()}`;
}

export async function handleCallback(code: string, redirectUri: string) {
  const res = await fetch(TL_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      client_id: process.env.TL_CLIENT_ID,
      client_secret: process.env.TL_CLIENT_SECRET,
      code,
      grant_type: 'authorization_code',
      redirect_uri: redirectUri,
    }),
  });

  if (!res.ok) throw new Error(`Teamleader token error: ${res.status}`);

  const data: any = await res.json();
  const expiresAt = new Date(Date.now() + data.expires_in * 1000).toISOString();

  runQuery(
    `INSERT OR REPLACE INTO auth_tokens (service, access_token, refresh_token, expires_at) VALUES (?, ?, ?, ?)`,
    ['teamleader', data.access_token, data.refresh_token, expiresAt]
  );

  return data;
}

async function getAccessToken(): Promise<string | null> {
  const row = getOne('SELECT * FROM auth_tokens WHERE service = ?', ['teamleader']);
  if (!row) return null;

  if (row.expires_at && new Date(row.expires_at as string) < new Date()) {
    if (!row.refresh_token) return null;

    const res = await fetch(TL_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        client_id: process.env.TL_CLIENT_ID,
        client_secret: process.env.TL_CLIENT_SECRET,
        refresh_token: row.refresh_token,
        grant_type: 'refresh_token',
      }),
    });

    if (!res.ok) return null;

    const data: any = await res.json();
    const expiresAt = new Date(Date.now() + data.expires_in * 1000).toISOString();

    runQuery(
      `INSERT OR REPLACE INTO auth_tokens (service, access_token, refresh_token, expires_at) VALUES (?, ?, ?, ?)`,
      ['teamleader', data.access_token, data.refresh_token, expiresAt]
    );

    return data.access_token;
  }

  return row.access_token as string;
}

async function tlRequest(path: string, body: Record<string, unknown>): Promise<any> {
  const token = await getAccessToken();
  if (!token) throw new Error('Teamleader not connected');

  const res = await fetch(`${TL_API_BASE}${path}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    throw new Error(`Teamleader API error: ${res.status} ${res.statusText}`);
  }
  return res.json();
}

export async function isConnected(): Promise<boolean> {
  const token = await getAccessToken();
  return token !== null;
}

export async function debugContactRaw(email: string): Promise<any> {
  try {
    return await tlRequest('/contacts.list', {
      filter: {
        email: {
          type: 'primary',
          email: email,
        },
      },
      include: 'responsible_user',
    });
  } catch (err: any) {
    return { error: err.message };
  }
}

export async function debugDealsRaw(contactId: string): Promise<any> {
  try {
    const list = await tlRequest('/deals.list', {
      filter: {
        customer: {
          type: 'contact',
          id: contactId,
        },
      },
      include: 'responsible_user',
    });

    // Also try deals.info for the first deal to see all fields
    let dealInfo = null;
    if (list.data && list.data.length > 0) {
      try {
        dealInfo = await tlRequest('/deals.info', { id: list.data[0].id });
      } catch (e: any) {
        dealInfo = { error: e.message };
      }
    }

    return { list, dealInfo };
  } catch (err: any) {
    return { error: err.message };
  }
}

export async function findContact(email: string): Promise<{ id: string; name: string; createdAt: string; responsibleUser: string | null } | null> {
  try {
    const result = await tlRequest('/contacts.list', {
      filter: {
        email: {
          type: 'primary',
          email: email,
        },
      },
      include: 'responsible_user',
    });

    if (result.data && result.data.length > 0) {
      const contact = result.data[0];

      // Build user lookup from sideloaded data
      const userMap = new Map<string, string>();
      if (result.included?.users) {
        for (const user of result.included.users) {
          userMap.set(user.id, `${user.first_name || ''} ${user.last_name || ''}`.trim());
        }
      }

      return {
        id: contact.id,
        name: `${contact.first_name} ${contact.last_name}`.trim(),
        createdAt: contact.added_at || contact.created_at || new Date().toISOString(),
        responsibleUser: contact.responsible_user?.id ? (userMap.get(contact.responsible_user.id) || null) : null,
      };
    }
    return null;
  } catch {
    return null;
  }
}

export async function findDeals(contactId: string): Promise<Array<{ id: string; title: string; status: string; createdAt: string; responsibleUser: string | null }>> {
  try {
    const listResult = await tlRequest('/deals.list', {
      filter: {
        customer: {
          type: 'contact',
          id: contactId,
        },
      },
    });

    if (!listResult.data) return [];

    // Build user lookup from list response (included.user, not included.users)
    const userMap = new Map<string, string>();
    const includedUsers = listResult.included?.user || listResult.included?.users || [];
    for (const user of includedUsers) {
      userMap.set(user.id, `${user.first_name || ''} ${user.last_name || ''}`.trim());
    }

    // If list didn't include users, fetch each deal individually
    const needsInfo = userMap.size === 0;

    const deals: Array<{ id: string; title: string; status: string; createdAt: string; responsibleUser: string | null }> = [];
    for (const deal of listResult.data) {
      let responsibleUser: string | null = null;

      if (deal.responsible_user?.id) {
        responsibleUser = userMap.get(deal.responsible_user.id) || null;
      }

      // Fallback to deals.info if we still don't have the user
      if (!responsibleUser && deal.responsible_user?.id && needsInfo) {
        try {
          const info = await tlRequest('/deals.info', { id: deal.id });
          const infoUsers = info.included?.user || info.included?.users || [];
          for (const u of infoUsers) {
            if (u.id === deal.responsible_user.id) {
              responsibleUser = `${u.first_name || ''} ${u.last_name || ''}`.trim();
              break;
            }
          }
        } catch {}
      }

      deals.push({
        id: deal.id,
        title: deal.title,
        status: deal.status,
        createdAt: deal.created_at || new Date().toISOString(),
        responsibleUser,
      });
    }
    return deals;
  } catch {
    return [];
  }
}

export async function syncTeamleaderForCustomers() {
  const token = await getAccessToken();
  if (!token) return;

  const customers = getAll('SELECT * FROM customers WHERE archived = 0');
  const d = getDb();

  for (const customer of customers) {
    const contact = await findContact(customer.email as string);
    if (!contact) continue;

    // Check if we already have a tl_contact event
    const existingContact = getOne(
      "SELECT id, metadata FROM timeline_events WHERE customer_id = ? AND type = 'tl_contact' AND metadata LIKE ?",
      [customer.id, `%"tl_id":"${contact.id}"%`]
    );

    if (!existingContact) {
      d.run(
        "INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, 'tl_contact', ?, ?, ?, 0, NULL, ?)",
        [customer.id, 'Contact in Teamleader', `Contact aangemaakt: ${contact.name}`, contact.createdAt, JSON.stringify({ tl_id: contact.id, actor: contact.responsibleUser || null })]
      );
    } else {
      // Backfill actor for existing tl_contact events
      if (contact.responsibleUser) {
        const meta = JSON.parse((existingContact.metadata as string) || '{}');
        if (!meta.actor) {
          meta.actor = contact.responsibleUser;
          d.run('UPDATE timeline_events SET metadata = ? WHERE id = ?', [JSON.stringify(meta), existingContact.id]);
        }
      }
    }

    // Check for deals
    const deals = await findDeals(contact.id);
    for (const deal of deals) {
      const existingDeal = getOne(
        "SELECT id FROM timeline_events WHERE customer_id = ? AND type = 'tl_deal' AND metadata LIKE ?",
        [customer.id, `%"tl_id":"${deal.id}"%`]
      );

      if (!existingDeal) {
        d.run(
          "INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, 'tl_deal', ?, ?, ?, 0, NULL, ?)",
          [customer.id, `Deal: ${deal.title}`, `Deal status: ${deal.status}`, deal.createdAt, JSON.stringify({ tl_id: deal.id, status: deal.status, actor: deal.responsibleUser || null })]
        );
      }
    }

    // Backfill actor for existing tl_deal events that don't have one yet
    for (const deal of deals) {
      if (deal.responsibleUser) {
        const existingDealEvent = getOne(
          "SELECT id, metadata FROM timeline_events WHERE customer_id = ? AND type = 'tl_deal' AND metadata LIKE ?",
          [customer.id, `%"tl_id":"${deal.id}"%`]
        );
        if (existingDealEvent) {
          const meta = JSON.parse((existingDealEvent.metadata as string) || '{}');
          if (!meta.actor) {
            meta.actor = deal.responsibleUser;
            d.run('UPDATE timeline_events SET metadata = ? WHERE id = ?', [JSON.stringify(meta), existingDealEvent.id]);
          }
        }
      }
    }

    // Update customer name from Teamleader if better
    if (contact.name && (!customer.name || customer.name === customer.email)) {
      d.run('UPDATE customers SET name = ? WHERE id = ?', [contact.name, customer.id]);
    }
  }

  saveDb();

  runQuery(
    `INSERT OR REPLACE INTO sync_state (key, value) VALUES ('last_teamleader_sync', datetime('now'))`,
    []
  );
}
