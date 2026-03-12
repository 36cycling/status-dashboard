export interface Customer {
  id: number;
  name: string;
  company: string;
  email: string;
  archived: boolean;
  created_at: string;
  events: TimelineEvent[];
}

export type EventType = 'email_in' | 'email_out' | 'tl_contact' | 'tl_deal';

export interface TimelineEvent {
  id: number;
  customer_id: number;
  type: EventType;
  subject: string;
  summary: string;
  date: string;
  is_replied: boolean;
  outlook_message_id: string | null;
  metadata: Record<string, unknown>;
}

export interface SyncStatus {
  last_sync: string | null;
  is_syncing: boolean;
}

export interface AuthStatus {
  outlook_connected: boolean;
  teamleader_connected: boolean;
}
