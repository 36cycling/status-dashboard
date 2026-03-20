import * as msal from '@azure/msal-node';
import { runQuery, getOne, getAll, saveDb, getDb } from '../db';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// Internal email addresses to skip as customers
const INTERNAL_EMAILS = [
  'info@36cycling.com',
  'jeroen@36cycling.com',
  'lisette@36cycling.com',
  'michael@36cycling.com',
  'lars@36cycling.com',
  'info@teamleader.eu',
  'noreply@',
  'no-reply@',
];

// Internal mailboxes to check for sent replies
const REPLY_MAILBOXES = [
  'info@36cycling.com',
  'jeroen@36cycling.com',
  'lisette@36cycling.com',
  'michael@36cycling.com',
  'lars@36cycling.com',
];

function isInternalEmail(email: string): boolean {
  const lower = email.toLowerCase();
  return INTERNAL_EMAILS.some(internal =>
    internal.includes('@')
      ? lower === internal
      : lower.startsWith(internal)
  );
}

// Domains that are never customer inquiries
const BULK_DOMAINS = [
  'mailchimp.com', 'sendinblue.com', 'brevo.com', 'mailgun.com',
  'sendgrid.net', 'constantcontact.com', 'hubspot.com', 'klaviyo.com',
  'mailerlite.com', 'campaignmonitor.com', 'exactonline.nl', 'exact.nl',
  'mollie.com', 'stripe.com', 'paypal.com', 'tikkie.me',
  'google.com', 'microsoft.com', 'linkedin.com', 'facebook.com',
  'instagram.com', 'twitter.com', 'tiktok.com',
  'teamleader.eu', 'teamleader.be',
  'lightspeedhq.com', 'lightspeed.com', 'shopify.com',
  'postnl.nl', 'dhl.com', 'dpd.nl', 'ups.com', 'fedex.com',
];

// Generic sender prefixes that are usually automated/system emails
const GENERIC_PREFIXES = [
  'noreply', 'no-reply', 'no_reply', 'donotreply', 'do-not-reply',
  'mailer-daemon', 'postmaster', 'bounce', 'notifications',
  'newsletter', 'nieuwsbrief', 'marketing', 'billing', 'invoice',
  'support', 'helpdesk', 'service', 'admin', 'system',
];

// Subject keywords that indicate non-inquiry emails
const EXCLUDE_SUBJECT_KEYWORDS = [
  'factuur', 'invoice', 'betaling', 'payment', 'creditnota',
  'newsletter', 'nieuwsbrief', 'unsubscribe', 'afmelden', 'uitschrijven',
  'order bevestiging', 'orderbevestiging', 'order confirmation',
  'verzending', 'tracking', 'shipment', 'delivered', 'afgeleverd',
  'wachtwoord', 'password reset', 'verificatie', 'verification',
  'out of office', 'automatisch antwoord', 'auto-reply', 'afwezigheid',
];

function isLikelyInquiry(msg: GraphMessage): boolean {
  const fromEmail = msg.from.emailAddress.address.toLowerCase();
  const fromName = msg.from.emailAddress.name.toLowerCase();
  const subject = msg.subject.toLowerCase();

  // Skip internal emails
  if (isInternalEmail(fromEmail)) return false;

  // Skip bulk/system domains
  const domain = fromEmail.split('@')[1] || '';
  if (BULK_DOMAINS.some(d => domain === d || domain.endsWith('.' + d))) return false;

  // Skip generic sender prefixes
  const localPart = fromEmail.split('@')[0] || '';
  if (GENERIC_PREFIXES.some(p => localPart === p || localPart.startsWith(p + '.'))) return false;

  // Skip based on subject keywords
  if (EXCLUDE_SUBJECT_KEYWORDS.some(kw => subject.includes(kw))) return false;

  // Skip if sender name looks automated
  if (fromName.includes('mailer') || fromName.includes('daemon') || fromName.includes('notification')) return false;

  return true;
}

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

// Parse contact form emails to extract the actual sender's name and email
// Handles \r\n, \n, and various form formats
function parseContactForm(body: string): { name: string; email: string } | null {
  // Normalize line endings
  const text = body.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  let name = '';
  let email = '';

  // Try to find email address in form data
  const emailPatterns = [
    // "E-mail\nkees@example.com" or "E-mail\n kees@example.com"
    /E-?[Mm]ail\s*\n\s*([^\s\n]+@[^\s\n]+)/i,
    // "E-mail: kees@example.com" or "E-Mail: kees@example.com"
    /E-?[Mm]ail\s*[:]\s*([^\s\n]+@[^\s\n]+)/i,
    // "Customer: Name\n\nE-Mail: email" (3D Designer format)
    /Customer\s*:\s*[^\n]+[\s\S]*?E-?[Mm]ail\s*:\s*([^\s\n]+@[^\s\n]+)/i,
  ];

  for (const pattern of emailPatterns) {
    const match = text.match(pattern);
    if (match) {
      email = match[1].trim().toLowerCase();
      break;
    }
  }

  if (!email) return null;

  // Don't extract if the extracted email is also internal
  if (isInternalEmail(email)) return null;

  // Try to find name
  const namePatterns = [
    // "Voornaam: X\nAchternaam: Y" (Teamleader form)
    /Voornaam\s*[:]\s*([^\n]+)[\s\S]*?Achternaam\s*[:]\s*([^\n]+)/i,
    // "Customer: Name" (3D Designer English)
    /Customer\s*[:]\s*([^\n]+)/i,
    // "Naam\nValue" (Dutch contact form - value on next line)
    /(?:^|\n)Naam\s*\n\s*([^\n]+)/i,
    // "Name\nValue" (English contact form - value on next line)
    /(?:^|\n)Name\s*\n\s*([^\n]+)/i,
    // "Naam: Value" or "Naam : Value"
    /(?:^|\n)Naam\s*[:]\s*([^\n]+)/i,
    // "Name: Value"
    /(?:^|\n)Name\s*[:]\s*([^\n]+)/i,
  ];

  for (const pattern of namePatterns) {
    const match = text.match(pattern);
    if (match) {
      if (match[2]) {
        // Voornaam + Achternaam
        name = `${match[1].trim()} ${match[2].trim()}`;
      } else {
        name = match[1].trim();
      }
      // Don't use the email as name if pattern matched email field
      if (name.includes('@')) {
        name = '';
      }
      break;
    }
  }

  return { name: name || email, email };
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

  // Only fetch mails from the last 2 months
  const twoMonthsAgo = new Date();
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
  const dateFilter = `receivedDateTime ge ${twoMonthsAgo.toISOString()}`;
  const selectFields = '$select=id,subject,bodyPreview,receivedDateTime,from,isRead,conversationId';

  // Collect all messages from multiple sources
  const allMessages: GraphMessage[] = [];
  const seenMessageIds = new Set<string>();

  // Source 1: Klantaanvragen folder (trusted — all emails are inquiries)
  if (folder) {
    const folderMessages = await graphRequest(
      `${mailboxPath}/mailFolders/${folder.id}/messages?$top=200&$orderby=receivedDateTime desc&$filter=${encodeURIComponent(dateFilter)}&${selectFields}`,
      token
    );
    for (const msg of folderMessages.value) {
      if (!seenMessageIds.has(msg.id)) {
        seenMessageIds.add(msg.id);
        allMessages.push(msg);
      }
    }
  }

  // Source 2: Inbox — heuristically filtered for likely inquiries
  try {
    const inboxMessages = await graphRequest(
      `${mailboxPath}/mailFolders('Inbox')/messages?$top=200&$orderby=receivedDateTime desc&$filter=${encodeURIComponent(dateFilter)}&${selectFields}`,
      token
    );
    for (const msg of inboxMessages.value as GraphMessage[]) {
      if (!seenMessageIds.has(msg.id) && isLikelyInquiry(msg)) {
        seenMessageIds.add(msg.id);
        allMessages.push(msg);
      }
    }
  } catch {
    // Inbox not accessible, continue with folder only
  }

  const messages = { value: allMessages };

  // Get sent items from ALL internal mailboxes to match replies (also last 2 months)
  const sentDateFilter = `sentDateTime ge ${twoMonthsAgo.toISOString()}`;
  const allSentItems: any[] = [];

  for (const mailbox of REPLY_MAILBOXES) {
    try {
      const sentResult = await graphRequest(
        `/users/${mailbox}/mailFolders('SentItems')/messages?$top=200&$orderby=sentDateTime desc&$filter=${encodeURIComponent(sentDateFilter)}&$select=id,subject,bodyPreview,sentDateTime,toRecipients,conversationId,from`,
        token
      );
      allSentItems.push(...sentResult.value);
    } catch {
      // Mailbox might not be accessible, skip
    }
  }

  // Build a set of conversation IDs that have sent replies
  const repliedConversations = new Set<string>();
  for (const sent of allSentItems) {
    repliedConversations.add(sent.conversationId);
  }

  const d = getDb();

  for (const msg of messages.value as GraphMessage[]) {
    // Skip if event already exists
    const existing = getOne('SELECT id FROM timeline_events WHERE outlook_message_id = ?', [msg.id]);
    if (existing) continue;

    // Determine the real sender
    const fromEmail = msg.from.emailAddress.address.toLowerCase();
    const fromName = msg.from.emailAddress.name;

    // If the email is from an internal/system address, try to parse the contact form
    let senderEmail = fromEmail;
    let senderName = fromName;

    if (isInternalEmail(fromEmail)) {
      const contactForm = parseContactForm(msg.bodyPreview);
      if (contactForm) {
        senderEmail = contactForm.email;
        senderName = contactForm.name;
      } else {
        // Internal email without parseable contact form data — skip
        continue;
      }
    }

    // Double-check: skip if sender is still internal after parsing
    if (isInternalEmail(senderEmail)) continue;

    // Create or find customer — reactivate if archived
    d.run('INSERT OR IGNORE INTO customers (name, email) VALUES (?, ?)', [senderName, senderEmail]);
    const customer = getOne('SELECT id, archived FROM customers WHERE email = ?', [senderEmail]);
    if (!customer) continue;

    // If customer was archived but has a new mail, reactivate
    if (customer.archived) {
      d.run('UPDATE customers SET archived = 0 WHERE id = ?', [customer.id]);
    }

    const isReplied = repliedConversations.has(msg.conversationId);

    d.run(
      'INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
      [customer.id, 'email_in', msg.subject, msg.bodyPreview.substring(0, 200), msg.receivedDateTime, isReplied ? 1 : 0, msg.id, '{}']
    );

    // If replied, also add the sent reply as an event
    if (isReplied) {
      const reply = allSentItems.find((s: any) => s.conversationId === msg.conversationId);
      if (reply) {
        const existingReply = getOne('SELECT id FROM timeline_events WHERE outlook_message_id = ?', [reply.id]);
        if (!existingReply) {
          const replyFrom = reply.from?.emailAddress?.name || reply.from?.emailAddress?.address || '';
          d.run(
            'INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            [customer.id, 'email_out', reply.subject, reply.bodyPreview.substring(0, 200), reply.sentDateTime, 0, reply.id, JSON.stringify({ actor: replyFrom })]
          );
        }
      }
    }
  }

  // Backfill actor for existing email_out events that don't have one yet
  const outEventsWithoutActor = getAll(
    "SELECT id, outlook_message_id FROM timeline_events WHERE type = 'email_out' AND (metadata = '{}' OR metadata NOT LIKE '%\"actor\"%')"
  );
  for (const ev of outEventsWithoutActor) {
    const sent = allSentItems.find((s: any) => s.id === ev.outlook_message_id);
    if (sent) {
      const actor = sent.from?.emailAddress?.name || sent.from?.emailAddress?.address || '';
      if (actor) {
        const existing = JSON.parse((ev.metadata as string) || '{}');
        existing.actor = actor;
        d.run('UPDATE timeline_events SET metadata = ? WHERE id = ?', [JSON.stringify(existing), ev.id]);
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
