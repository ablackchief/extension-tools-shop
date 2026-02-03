import { AlertTriangle, Lightbulb, Info, XCircle } from 'lucide-react';
import { cn } from '@/lib/utils';

type CalloutType = 'warning' | 'tip' | 'important' | 'danger';

interface CalloutProps {
  type?: CalloutType;
  children: React.ReactNode;
}

const calloutConfig: Record<CalloutType, {
  icon: React.ElementType;
  bgColor: string;
  borderColor: string;
  iconColor: string;
  title: string;
}> = {
  warning: {
    icon: AlertTriangle,
    bgColor: 'bg-amber-50',
    borderColor: 'border-amber-500',
    iconColor: 'text-amber-600',
    title: 'Warning',
  },
  tip: {
    icon: Lightbulb,
    bgColor: 'bg-green-50',
    borderColor: 'border-green-500',
    iconColor: 'text-green-600',
    title: 'Tip',
  },
  important: {
    icon: Info,
    bgColor: 'bg-blue-50',
    borderColor: 'border-blue-500',
    iconColor: 'text-blue-600',
    title: 'Important',
  },
  danger: {
    icon: XCircle,
    bgColor: 'bg-red-50',
    borderColor: 'border-red-500',
    iconColor: 'text-red-600',
    title: 'Danger',
  },
};

export function Callout({ type = 'important', children }: CalloutProps) {
  const config = calloutConfig[type];
  const Icon = config.icon;

  return (
    <div
      className={cn(
        'my-6 rounded-r-xl border-l-4 p-5',
        config.bgColor,
        config.borderColor
      )}
    >
      <div className="flex items-start gap-3">
        <Icon className={cn('h-5 w-5 mt-0.5 flex-shrink-0', config.iconColor)} />
        <div className="flex-1">
          <p className={cn('font-semibold text-sm mb-1', config.iconColor)}>
            {config.title}
          </p>
          <div className="text-slate-700 text-sm [&>p]:m-0">{children}</div>
        </div>
      </div>
    </div>
  );
}
