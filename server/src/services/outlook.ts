import * as msal from '@azure/msal-node';
import { runQuery, getOne, getAll, saveDb, getDb } from '../db';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

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

// Client credentials flow - no user login needed
async function getAppToken(): Promise<string> {
  const result = await getMsalClient().acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  if (!result || !result.accessToken) {
    throw new Error('Failed to acquire app token');
  }
  return result.accessToken;
}

function getMailboxPath(): string {
  const targetMailbox = process.env.OUTLOOK_MAILBOX || '';
  if (!targetMailbox) {
    throw new Error('OUTLOOK_MAILBOX environment variable not set. Set it to the email address to read (e.g. info@36cycling.com).');
  }
  return `/users/${targetMailbox}`;
}

async function graphRequest(path: string, token: string): Promise<any> {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API error: ${res.status} ${res.statusText} - ${body}`);
  }
  return res.json();
}

export async function isConnected(): Promise<boolean> {
  try {
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const tenantId = process.env.AZURE_TENANT_ID;
    const mailbox = process.env.OUTLOOK_MAILBOX;
    return !!(clientId && clientSecret && tenantId && mailbox);
  } catch {
    return false;
  }
}

export async function listAllFolders(): Promise<any[]> {
  const token = await getAppToken();
  const mailboxPath = getMailboxPath();

  const folders = await graphRequest(`${mailboxPath}/mailFolders?$top=100`, token);
  const result: any[] = [];

  for (const f of folders.value) {
    const entry: any = { name: f.displayName, id: f.id, children: [] };
    try {
      const children = await graphRequest(`${mailboxPath}/mailFolders/${f.id}/childFolders?$top=100`, token);
      entry.children = children.value.map((c: any) => ({ name: c.displayName, id: c.id }));
    } catch {
      // no children
    }
    result.push(entry);
  }

  return result;
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
  const token = await getAppToken();
  const mailboxPath = getMailboxPath();
  const folderName = process.env.OUTLOOK_FOLDER_NAME || 'Klantaanvragen';

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
