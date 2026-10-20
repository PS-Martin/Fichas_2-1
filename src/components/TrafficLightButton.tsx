import React from 'react';
import { TrafficColor } from '../types';

interface TrafficLightGroupProps {
  label: string;
  icon?: React.ReactNode;
  value: TrafficColor;
  onChange: (newValue: TrafficColor) => void;
  disabled?: boolean;
}

export const TrafficLightGroup: React.FC<TrafficLightGroupProps> = ({
  label,
  icon,
  value,
  onChange,
  disabled = false,
}) => {
  return (
    <div className="flex items-center justify-between p-2.5 bg-white rounded-lg border border-slate-200 shadow-xs hover:border-slate-300 transition-colors">
      <div className="flex items-center gap-2">
        {icon && <span className="text-slate-500 text-sm">{icon}</span>}
        <span className="font-semibold text-slate-800 text-sm tracking-tight">{label}</span>
      </div>

      {/* Traffic light 3-button row: Green, Yellow, Red (excluyentes) */}
      <div className="flex items-center gap-2 bg-slate-900/90 p-1.5 rounded-full border border-slate-700/80 shadow-inner">
        {/* VERDE */}
        <button
          type="button"
          disabled={disabled}
          onClick={() => onChange('green')}
          title={`${label}: Verde`}
          aria-label={`${label}: Verde`}
          className={`relative w-7 h-7 rounded-full flex items-center justify-center transition-all duration-200 focus:outline-hidden ${
            value === 'green'
              ? 'bg-emerald-500 shadow-[0_0_14px_rgba(16,185,129,0.9)] ring-2 ring-emerald-300 scale-105'
              : 'bg-emerald-950/80 opacity-45 hover:opacity-75 hover:bg-emerald-900 border border-emerald-900/40'
          }`}
        >
          {value === 'green' && (
            <span className="w-2 h-2 rounded-full bg-white/90 shadow-xs"></span>
          )}
        </button>

        {/* AMARILLO */}
        <button
          type="button"
          disabled={disabled}
          onClick={() => onChange('yellow')}
          title={`${label}: Amarillo`}
          aria-label={`${label}: Amarillo`}
          className={`relative w-7 h-7 rounded-full flex items-center justify-center transition-all duration-200 focus:outline-hidden ${
            value === 'yellow'
              ? 'bg-amber-400 shadow-[0_0_14px_rgba(245,158,11,0.9)] ring-2 ring-amber-200 scale-105'
              : 'bg-amber-950/80 opacity-45 hover:opacity-75 hover:bg-amber-900 border border-amber-900/40'
          }`}
        >
          {value === 'yellow' && (
            <span className="w-2 h-2 rounded-full bg-white/90 shadow-xs"></span>
          )}
        </button>

        {/* ROJO */}
        <button
          type="button"
          disabled={disabled}
          onClick={() => onChange('red')}
          title={`${label}: Rojo`}
          aria-label={`${label}: Rojo`}
          className={`relative w-7 h-7 rounded-full flex items-center justify-center transition-all duration-200 focus:outline-hidden ${
            value === 'red'
              ? 'bg-rose-600 shadow-[0_0_14px_rgba(225,29,72,0.9)] ring-2 ring-rose-300 scale-105'
              : 'bg-rose-950/80 opacity-45 hover:opacity-75 hover:bg-rose-900 border border-rose-900/40'
          }`}
        >
          {value === 'red' && (
            <span className="w-2 h-2 rounded-full bg-white/90 shadow-xs"></span>
          )}
        </button>
      </div>
    </div>
  );
};

/**
 * Non-clickable traffic light badge indicator for the section overview list:
 * "con un único 'botón' no clicable que estará iluminado en verde si todas las etiquetas correspondientes
 * a su numero de bloque estan en verde, rojo si hay una sola etiqueta marcada en rojo, y amarillo en cualquier otro caso"
 */
interface SummaryStatusBadgeProps {
  color: TrafficColor;
  size?: 'sm' | 'md';
}

export const SummaryStatusBadge: React.FC<SummaryStatusBadgeProps> = ({ color, size = 'md' }) => {
  const sizeClasses = size === 'sm' ? 'w-5 h-5 text-[10px]' : 'w-6 h-6 text-xs';

  let colorClasses = '';
  let title = '';

  switch (color) {
    case 'green':
      colorClasses = 'bg-emerald-500 shadow-[0_0_10px_rgba(16,185,129,0.85)] border-emerald-300 text-white';
      title = 'Estado: Verde (Todos en verde)';
      break;
    case 'yellow':
      colorClasses = 'bg-amber-400 shadow-[0_0_10px_rgba(245,158,11,0.85)] border-amber-200 text-amber-950';
      title = 'Estado: Amarillo (En progreso / Mixto)';
      break;
    case 'red':
    default:
      colorClasses = 'bg-rose-600 shadow-[0_0_10px_rgba(225,29,72,0.85)] border-rose-300 text-white';
      title = 'Estado: Rojo (Al menos una etiqueta en rojo)';
      break;
  }

  return (
    <div
      title={title}
      aria-label={title}
      className={`rounded-full border-2 flex items-center justify-center font-bold pointer-events-none select-none transition-all duration-300 ${sizeClasses} ${colorClasses}`}
    >
      <span className="w-1.5 h-1.5 rounded-full bg-white/80"></span>
    </div>
  );
};
