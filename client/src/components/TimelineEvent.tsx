import { useState } from 'react';
import type { TimelineEvent as TEvent } from '../../../shared/types';

interface Props {
  event: TEvent;
}

function formatDate(dateStr: string): string {
  const d = new Date(dateStr);
  const day = d.getDate();
  const months = ['jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  const hours = d.getHours().toString().padStart(2, '0');
  const mins = d.getMinutes().toString().padStart(2, '0');
  return `${day} ${month} ${year} - ${hours}:${mins}`;
}

function getEventStyle(event: TEvent): string {
  switch (event.type) {
    case 'email_out':
      return 'bg-blue-500 border-blue-500 text-white';
    case 'tl_contact':
    case 'tl_deal':
      return 'bg-teal-500 border-teal-500 text-white';
    case 'email_in':
      if (!event.is_replied) {
        return 'bg-white border-2 border-red-500 text-red-600';
      }
      return 'bg-white border-2 border-blue-500 text-blue-800';
    default:
      return 'bg-white border-2 border-slate-300 text-slate-700';
  }
}

export default function TimelineEvent({ event }: Props) {
  const [showTooltip, setShowTooltip] = useState(false);

  return (
    <div
      className="relative flex-shrink-0"
      onMouseEnter={() => setShowTooltip(true)}
      onMouseLeave={() => setShowTooltip(false)}
    >
      <div className={`w-44 p-3 rounded-md cursor-default ${getEventStyle(event)}`}>
        <div className="text-xs font-medium mb-1">{formatDate(event.date)}</div>
        <div className="text-xs leading-relaxed line-clamp-3">
          {event.type === 'tl_contact' || event.type === 'tl_deal'
            ? event.summary
            : `${event.subject}`}
        </div>
      </div>

      {showTooltip && (
        <div className="absolute z-50 bottom-full left-0 mb-2 w-72 bg-slate-900 text-white text-xs rounded-lg p-3 shadow-xl pointer-events-none">
          <div className="font-semibold mb-1">{event.subject}</div>
          <div className="text-slate-300 mb-2">{formatDate(event.date)}</div>
          <div className="text-slate-200 leading-relaxed">{event.summary || 'Geen samenvatting beschikbaar'}</div>
          <div className="absolute bottom-0 left-4 transform translate-y-1/2 rotate-45 w-2 h-2 bg-slate-900" />
        </div>
      )}
    </div>
  );
}
