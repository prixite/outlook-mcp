export interface OutlookMessage {
  id: number;
  subject: string;
  senderName: string;
  senderEmail: string;
  dateReceived: string;
  isRead: boolean;
  hasAttachments: boolean;
  folder?: string;
  preview?: string;
  body?: string;
}

export interface OutlookEvent {
  id: number;
  subject: string;
  startTime: string;
  endTime: string;
  location: string;
  isAllDay: boolean;
  body?: string;
}

export interface OutlookFolder {
  name: string;
  unreadCount: number;
}
