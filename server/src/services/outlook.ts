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

// ── Inbox heuristic filter ──
// Approach: ONLY allow emails that look like a real person asking a question.
// Everything else is excluded.

// Personal email providers — these are almost always real people
const PERSONAL_DOMAINS = [
  'gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'hotmail.nl',
  'live.com', 'live.nl', 'msn.com', 'yahoo.com', 'yahoo.nl',
  'icloud.com', 'me.com', 'mac.com', 'protonmail.com', 'proton.me',
  'ziggo.nl', 'kpnmail.nl', 'xs4all.nl', 'hetnet.nl', 'home.nl',
  'planet.nl', 'upcmail.nl', 'casema.nl', 'chello.nl', 'quicknet.nl',
  'tele2.nl', 'telfort.nl', 'online.nl', 'solcon.nl', 'wxs.nl',
];

// Domains that are NEVER customer inquiries (companies/systems)
const BLOCKED_DOMAINS = [
  // Bulk mail / marketing
  'mailchimp.com', 'sendinblue.com', 'brevo.com', 'mailgun.com', 'mailgun.org',
  'sendgrid.net', 'constantcontact.com', 'hubspot.com', 'hubspotmail.com', 'klaviyo.com',
  'mailerlite.com', 'campaignmonitor.com', 'activecampaign.com',
  // Financial / payment
  'exactonline.nl', 'exact.nl', 'mollie.com', 'stripe.com', 'paypal.com',
  'paypal.nl', 'tikkie.me', 'adyen.com',
  // Big tech / social
  'google.com', 'microsoft.com', 'linkedin.com', 'facebook.com', 'facebookmail.com',
  'instagram.com', 'twitter.com', 'tiktok.com', 'apple.com',
  // CRM / tools
  'teamleader.eu', 'teamleader.be', 'salesforce.com',
  // E-commerce / platforms
  'lightspeedhq.com', 'lightspeed.com', 'shopify.com', 'bol.com',
  'amazon.com', 'amazon.nl', 'coolblue.nl', 'zalando.nl',
  // Shipping / logistics
  'postnl.nl', 'postnl.com', 'dhl.com', 'dpd.nl', 'ups.com', 'fedex.com',
  'gls-group.eu', 'budbee.com', 'trunkrs.nl',
  // Telecom / utilities
  'kpn.com', 'kpn.nl', 'vodafone.nl', 'vodafone.com', 't-mobile.nl',
  'ziggo.com', 'odido.nl',
  // Software / IT
  'sap.com', 'oracle.com', 'atlassian.com', 'jira.com', 'slack.com',
  'zendesk.com', 'freshdesk.com', 'intercom.com', 'github.com',
  // Government / organizations
  'belastingdienst.nl', 'kvk.nl', 'uwv.nl', 'rijksoverheid.nl',
];

// Sender prefixes that are always automated
const BLOCKED_PREFIXES = [
  'noreply', 'no-reply', 'no_reply', 'donotreply', 'do-not-reply',
  'mailer-daemon', 'postmaster', 'bounce', 'notifications', 'notification',
  'newsletter', 'nieuwsbrief', 'marketing', 'billing', 'invoice',
  'helpdesk', 'admin', 'system', 'alert', 'alerts', 'news',
  'updates', 'info', 'klantenservice', 'customerservice', 'webmaster',
  'orders', 'order', 'shipping', 'delivery', 'tracking',
];

// Subject keywords that indicate non-inquiry emails
const BLOCKED_SUBJECTS = [
  'factuur', 'invoice', 'betaling', 'payment', 'creditnota', 'credit nota',
  'newsletter', 'nieuwsbrief', 'unsubscribe', 'afmelden', 'uitschrijven',
  'order bevestiging', 'orderbevestiging', 'order confirmation', 'bestelling',
  'verzending', 'tracking', 'shipment', 'delivered', 'afgeleverd', 'pakket',
  'wachtwoord', 'password', 'verificatie', 'verification', 'verify',
  'out of office', 'automatisch antwoord', 'auto-reply', 'afwezigheid',
  'welkom bij', 'welcome to', 'bevestig je', 'confirm your',
  'your account', 'je account', 'inloggen', 'login', 'sign in',
  'abonnement', 'subscription', 'renewal', 'verlenging',
];

function isLikelyInquiry(msg: GraphMessage): boolean {
  const fromEmail = msg.from.emailAddress.address.toLowerCase();
  const subject = (msg.subject || '').toLowerCase();
  const domain = fromEmail.split('@')[1] || '';
  const localPart = fromEmail.split('@')[0] || '';

  // Always skip internal emails
  if (isInternalEmail(fromEmail)) return false;

  // Always skip blocked domains
  if (BLOCKED_DOMAINS.some(d => domain === d || domain.endsWith('.' + d))) return false;

  // Always skip blocked prefixes
  if (BLOCKED_PREFIXES.some(p => localPart === p || localPart.startsWith(p + '.') || localPart.startsWith(p + '-') || localPart.startsWith(p + '_'))) return false;

  // Always skip blocked subjects
  if (BLOCKED_SUBJECTS.some(kw => subject.includes(kw))) return false;

  // Personal email domains are always OK (real people)
  if (PERSONAL_DOMAINS.includes(domain)) return true;

  // For business domains: only allow if sender looks like a person
  // (has a first.last or firstname pattern, not a generic prefix)
  const looksLikePerson = localPart.includes('.') || localPart.includes('_') || /^[a-z]{2,15}$/.test(localPart);
  if (looksLikePerson) return true;

  // Everything else: skip
  return false;
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

    // Check reply: first by conversationId, then fallback by recipient email
    let isReplied = repliedConversations.has(msg.conversationId);
    let reply = isReplied
      ? allSentItems.find((s: any) => s.conversationId === msg.conversationId)
      : null;

    // Fallback: find a sent item TO this customer's email address
    if (!isReplied) {
      reply = allSentItems.find((s: any) =>
        s.toRecipients?.some((r: any) => r.emailAddress?.address?.toLowerCase() === senderEmail)
      );
      if (reply) isReplied = true;
    }

    d.run(
      'INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
      [customer.id, 'email_in', msg.subject, msg.bodyPreview.substring(0, 200), msg.receivedDateTime, isReplied ? 1 : 0, msg.id, JSON.stringify({ conversationId: msg.conversationId })]
    );

    // If replied, also add the sent reply as an event
    if (isReplied && reply) {
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

  // Backfill: re-check unreplied email_in events against all mailbox sent items
  const unrepliedEvents = getAll(
    "SELECT id, outlook_message_id, customer_id, metadata FROM timeline_events WHERE type = 'email_in' AND is_replied = 0"
  );

  // Build a map of conversationId -> sent reply for quick lookup
  const sentByConversation = new Map<string, any>();
  for (const sent of allSentItems) {
    if (sent.conversationId && !sentByConversation.has(sent.conversationId)) {
      sentByConversation.set(sent.conversationId, sent);
    }
  }

  for (const ev of unrepliedEvents) {
    const meta = JSON.parse((ev.metadata as string) || '{}');
    let convId = meta.conversationId;

    // If no conversationId stored, try to fetch it from Graph
    if (!convId && ev.outlook_message_id) {
      try {
        const msgDetail = await graphRequest(
          `${mailboxPath}/messages/${ev.outlook_message_id}?$select=conversationId`,
          token
        );
        convId = msgDetail?.conversationId;
        // Store it for next time
        if (convId) {
          meta.conversationId = convId;
          d.run('UPDATE timeline_events SET metadata = ? WHERE id = ?', [JSON.stringify(meta), ev.id]);
        }
      } catch {
        // Message might have been deleted
      }
    }

    // Try conversationId match first
    let reply = convId ? sentByConversation.get(convId) : undefined;

    // Fallback: match by recipient email
    if (!reply) {
      const customer = getOne('SELECT email FROM customers WHERE id = ?', [ev.customer_id]);
      if (customer) {
        const custEmail = (customer.email as string).toLowerCase();
        reply = allSentItems.find((s: any) =>
          s.toRecipients?.some((r: any) => r.emailAddress?.address?.toLowerCase() === custEmail)
        );
      }
    }

    if (reply) {
      // Mark as replied
      d.run('UPDATE timeline_events SET is_replied = 1 WHERE id = ?', [ev.id]);

      // Add the reply as email_out event if not already there
      const existingReply = getOne('SELECT id FROM timeline_events WHERE outlook_message_id = ?', [reply.id]);
      if (!existingReply) {
        const replyFrom = reply.from?.emailAddress?.name || reply.from?.emailAddress?.address || '';
        d.run(
          'INSERT INTO timeline_events (customer_id, type, subject, summary, date, is_replied, outlook_message_id, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
          [ev.customer_id, 'email_out', reply.subject, reply.bodyPreview?.substring(0, 200) || '', reply.sentDateTime, 0, reply.id, JSON.stringify({ actor: replyFrom })]
        );
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
