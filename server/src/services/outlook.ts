import * as msal from '@azure/msal-node';
import { runQuery, getOne, getAll, saveDb, getDb } from '../db';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['https://graph.microsoft.com/Mail.ReadWrite.Shared'];

let _msalClient: msal.ConfidentialClientApplication | null = null;

function getMsalClient(): msal.ConfidentialClientApplication {
  if (!_msalClient) {
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    if (!clientId || !clientSecret) {
      throw new Error('Azure credentials not configured. Set AZURE_CLIENT_ID and AZURE_CLIENT_SECRET.');
    }
    _msalClient = new msal.ConfidentialClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID || 'common'}`,
        clientSecret,
      },
    });
  }
  return _msalClient;
}

export function getAuthUrl(redirectUri: string): Promise<string> {
  return getMsalClient().getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri,
  });
}

export async function handleCallback(code: string, redirectUri: string) {
  const result = await getMsalClient().acquireTokenByCode({
    code,
    scopes: SCOPES,
    redirectUri,
  });

  const expiresAt = result.expiresOn?.toISOString() || new Date(Date.now() + 3600000).toISOString();

  runQuery(
    `INSERT OR REPLACE INTO auth_tokens (service, access_token, refresh_token, expires_at) VALUES (?, ?, ?, ?)`,
    ['outlook', result.accessToken, '', expiresAt]
  );

  return result;
}

async function getAccessToken(): Promise<string | null> {
  const row = getOne('SELECT * FROM auth_tokens WHERE service = ?', ['outlook']);
  if (!row) return null;

  if (row.expires_at && new Date(row.expires_at as string) < new Date()) {
    try {
      const accounts = await getMsalClient().getTokenCache().getAllAccounts();
      if (accounts.length > 0) {
        const result = await getMsalClient().acquireTokenSilent({
          account: accounts[0],
          scopes: SCOPES,
        });
        if (result) {
          runQuery(
            `INSERT OR REPLACE INTO auth_tokens (service, access_token, refresh_token, expires_at) VALUES (?, ?, ?, ?)`,
            ['outlook', result.accessToken, '', result.expiresOn?.toISOString() || '']
          );
          return result.accessToken;
        }
      }
    } catch {
      return null;
    }
  }

  return row.access_token as string;
}

async function graphRequest(path: string, token: string): Promise<any> {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(`Graph API error: ${res.status} ${res.statusText}`);
  }
  return res.json();
}

export async function isConnected(): Promise<boolean> {
  try {
    const token = await getAccessToken();
    return token !== null;
  } catch {
    return false;
  }
}

interface GraphMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  from: { emailAddress: { name: string; address: string } };
  isRead: boolean;
  conversationId: string;
}

interface GraphMailFolder {
  id: string;
  displayName: string;
}

export async function syncMails() {
  const token = await getAccessToken();
  if (!token) throw new Error('Outlook not connected');

  const folderName = process.env.OUTLOOK_FOLDER_NAME || 'Klantaanvragen';
  const targetMailbox = process.env.OUTLOOK_MAILBOX || '';
  const mailboxPath = targetMailbox ? `/users/${targetMailbox}` : '/me';

  // Find the target folder (also search in child folders of Inbox)
  const folders = await graphRequest(`${mailboxPath}/mailFolders?$top=100`, token);
  let folder = folders.value.find((f: GraphMailFolder) => f.displayName === folderName);

  if (!folder) {
    // Search in child folders of each top-level folder (e.g. Inbox subfolders)
    for (const parentFolder of folders.value) {
      try {
        const children = await graphRequest(`${mailboxPath}/mailFolders/${parentFolder.id}/childFolders?$top=100`, token);
        folder = children.value.find((f: GraphMailFolder) => f.displayName === folderName);
        if (folder) break;
      } catch {
        // No child folders, continue
      }
    }
  }

  if (!folder) throw new Error(`Mail folder "${folderName}" not found`);

  // Get messages from the folder
  const messages = await graphRequest(
    `${mailboxPath}/mailFolders/${folder.id}/messages?$top=200&$orderby=receivedDateTime desc&$select=id,subject,bodyPreview,receivedDateTime,from,isRead,conversationId`,
    token
  );

  // Get sent items to match replies
  const sentItems = await graphRequest(
    `${mailboxPath}/mailFolders('SentItems')/messages?$top=200&$orderby=sentDateTime desc&$select=id,subject,bodyPreview,sentDateTime,toRecipients,conversationId`,
    token
  );

  // Build a set of conversation IDs that have sent replies
  const repliedConversations = new Set<string>();
  for (const sent of sentItems.value) {
    repliedConversations.add(sent.conversationId);
  }

  const d = getDb();

  for (const msg of messages.value as GraphMessage[]) {
    const senderEmail = msg.from.emailAddress.address.toLowerCase();
    const senderName = msg.from.emailAddress.name;

    // Skip if event already exists
    const existing = getOne('SELECT id FROM timeline_events WHERE outlook_message_id = ?', [msg.id]);
    if (existing) continue;

    // Create or find customer
    d.run('INSERT OR IGNORE INTO customers (name, email) VALUES (?, ?)', [senderName, senderEmail]);
    const customer = getOne('SELECT id FROM customers WHERE email = ?', [senderEmail]);
    if (!customer) continue;

    const isReplied = repliedConversations.has(msg.conversationId);

    d.run(
      'INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
      [customer.id, 'email_in', msg.subject, msg.bodyPreview.substring(0, 200), msg.receivedDateTime, isReplied ? 1 : 0, msg.id, '{}']
    );

    // If replied, also add the sent reply as an event
    if (isReplied) {
      const reply = sentItems.value.find((s: any) => s.conversationId === msg.conversationId);
      if (reply) {
        const existingReply = getOne('SELECT id FROM timeline_events WHERE outlook_message_id = ?', [reply.id]);
        if (!existingReply) {
          d.run(
            'INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            [customer.id, 'email_out', reply.subject, reply.bodyPreview.substring(0, 200), reply.sentDateTime, 0, reply.id, '{}']
          );
        }
      }
    }
  }

  saveDb();

  runQuery(
    `INSERT OR REPLACE INTO sync_state (key, value) VALUES ('last_outlook_sync', datetime('now'))`,
    []
  );

  return { synced: messages.value.length };
}
