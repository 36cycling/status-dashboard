import type { Customer } from '../../../shared/types';
import TimelineEvent from './TimelineEvent';

interface Props {
  customer: Customer;
  onArchive: (id: number) => void;
}

export default function CustomerRow({ customer, onArchive }: Props) {
  const handleArchive = () => {
    if (confirm(`Wil je de tracking voor ${customer.name || customer.email} sluiten?`)) {
      onArchive(customer.id);
    }
  };

  return (
    <div className="flex items-start gap-0 mb-4">
      {/* Customer info card */}
      <div className="flex-shrink-0 w-52 bg-blue-800 text-white p-3 rounded-l-md relative">
        <button
          onClick={handleArchive}
          className="absolute -right-3 -top-3 w-6 h-6 bg-white border-2 border-slate-300 rounded-full text-slate-500 text-xs font-bold flex items-center justify-center hover:bg-red-50 hover:border-red-400 hover:text-red-600 z-10"
          title="Tracking sluiten"
        >
          X
        </button>
        <div className="font-bold text-sm truncate">{customer.name || 'Onbekend'}</div>
        <div className="text-xs text-blue-200 truncate">{customer.company || ''}</div>
        <div className="text-xs text-blue-200 truncate">{customer.email}</div>
      </div>

      {/* Timeline */}
      <div className="flex gap-2 overflow-x-auto pb-2 px-2 flex-1 min-w-0">
        {customer.events.map((event) => (
          <TimelineEvent key={event.id} event={event} />
        ))}
        {customer.events.length === 0 && (
          <div className="text-sm text-slate-400 italic py-3 px-4">Geen events</div>
        )}
      </div>
    </div>
  );
}
