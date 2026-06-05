import React from 'react';
import { ChevronRight } from 'lucide-react';

/* ─── Badge ──────────────────────────────────────────────────────────────── */
type BadgeVariant = 'ok' | 'warning' | 'error' | 'gray' | 'blue' | 'navy';
const badgeStyles: Record<BadgeVariant, { bg: string; color: string; border: string }> = {
  ok:      { bg: '#f0fdf4', color: '#15803d', border: '#bbf7d0' },
  warning: { bg: '#fffbeb', color: '#b45309', border: '#fde68a' },
  error:   { bg: '#fef2f2', color: '#b91c1c', border: '#fecaca' },
  gray:    { bg: '#f8fafc', color: '#475569', border: '#e2e8f0' },
  blue:    { bg: '#eff6ff', color: '#1d4ed8', border: '#bfdbfe' },
  navy:    { bg: '#f0f4ff', color: '#1e2a4a', border: '#c7d2fe' },
};
export function Badge({ variant, children }: { variant: BadgeVariant; children: React.ReactNode }) {
  const s = badgeStyles[variant];
  return (
    <span
      className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold"
      style={{ background: s.bg, color: s.color, border: `1px solid ${s.border}` }}
    >
      {children}
    </span>
  );
}

/* ─── Card ───────────────────────────────────────────────────────────────── */
export function Card({
  children,
  className = '',
  noPad = false,
}: {
  children: React.ReactNode;
  className?: string;
  noPad?: boolean;
}) {
  return (
    <div
      className={`bg-white rounded-xl ${noPad ? '' : ''} ${className}`}
      style={{ border: '1px solid #e8edf2', boxShadow: '0 1px 3px rgba(30,42,74,0.06)' }}
    >
      {children}
    </div>
  );
}

/* ─── CardHeader ─────────────────────────────────────────────────────────── */
export function CardHeader({
  title,
  subtitle,
  icon: Icon,
  iconColor,
  action,
  step,
}: {
  title: string;
  subtitle?: string;
  icon?: React.ElementType;
  iconColor?: string;
  action?: React.ReactNode;
  step?: number;
}) {
  return (
    <div
      className="flex items-center justify-between px-5 py-3.5"
      style={{ borderBottom: '1px solid #f1f5f9' }}
    >
      <div className="flex items-center gap-3">
        {step !== undefined && (
          <span
            className="w-8 h-8 rounded-full flex items-center justify-center text-sm font-black text-white flex-shrink-0"
            style={{ background: '#f97316' }}
          >
            {step}
          </span>
        )}
        {Icon && !step && (
          <span
            className="w-7 h-7 rounded-lg flex items-center justify-center flex-shrink-0"
            style={{ background: iconColor ? `${iconColor}18` : '#fff7ed' }}
          >
            <Icon size={14} style={{ color: iconColor ?? '#f97316' }} strokeWidth={2} />
          </span>
        )}
        <div>
          <h3 className="text-base font-extrabold leading-tight" style={{ color: '#1e293b' }}>{title}</h3>
          {subtitle && <p className="text-xs mt-0.5 leading-tight" style={{ color: '#94a3b8' }}>{subtitle}</p>}
        </div>
      </div>
      {action && <div>{action}</div>}
    </div>
  );
}

/* ─── SectionLabel ───────────────────────────────────────────────────────── */
export function SectionLabel({ letter, title }: { letter: string; title: string }) {
  return (
    <div className="flex items-center gap-2.5 mb-4">
      <span
        className="w-8 h-8 rounded-lg flex items-center justify-center text-sm font-black text-white flex-shrink-0"
        style={{ background: '#1e2a4a' }}
      >
        {letter}
      </span>
      <h3 className="text-base font-extrabold" style={{ color: '#1e293b' }}>
        {title}
      </h3>
      <div className="flex-1 h-px" style={{ background: '#f1f5f9' }} />
    </div>
  );
}

/* ─── Field ──────────────────────────────────────────────────────────────── */
export function Field({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <div>
      <label className="block text-xs font-semibold mb-1.5 uppercase tracking-wider" style={{ color: '#64748b' }}>
        {label}
      </label>
      {children}
    </div>
  );
}

/* ─── Input ──────────────────────────────────────────────────────────────── */
export function Input({
  value,
  onChange,
  placeholder,
  readOnly,
  mono,
}: {
  value?: string;
  onChange?: (v: string) => void;
  placeholder?: string;
  readOnly?: boolean;
  mono?: boolean;
}) {
  return (
    <input
      value={value}
      onChange={onChange ? e => onChange(e.target.value) : undefined}
      placeholder={placeholder}
      readOnly={readOnly}
      className={`w-full px-3 py-2 rounded-lg text-sm outline-none transition-all ${mono ? 'font-mono' : ''}`}
      style={{
        border: '1px solid #e2e8f0',
        color: readOnly ? '#94a3b8' : '#1e293b',
        background: readOnly ? '#f8fafc' : '#fff',
      }}
      onFocus={e => { if (!readOnly) e.target.style.boxShadow = '0 0 0 3px rgba(249,115,22,0.14)'; }}
      onBlur={e => { e.target.style.boxShadow = 'none'; }}
    />
  );
}

/* ─── Select ─────────────────────────────────────────────────────────────── */
export function Select({
  value,
  onChange,
  options,
}: {
  value: string;
  onChange: (v: string) => void;
  options: string[];
}) {
  return (
    <select
      value={value}
      onChange={e => onChange(e.target.value)}
      className="w-full px-3 py-2 rounded-lg text-sm outline-none"
      style={{ border: '1px solid #e2e8f0', color: '#1e293b', background: '#fff' }}
    >
      {options.map(o => <option key={o}>{o}</option>)}
    </select>
  );
}

/* ─── Btn ─────────────────────────────────────────────────────────────────── */
type BtnVariant = 'primary' | 'secondary' | 'ghost' | 'danger';
const btnBase = 'inline-flex items-center justify-center gap-2 rounded-lg text-sm font-semibold transition-all cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed';
const btnSizes = { sm: 'px-3 py-1.5 text-xs', md: 'px-4 py-2', lg: 'px-5 py-2.5', xl: 'px-6 py-3' };
const btnStyles: Record<BtnVariant, string> = {
  primary:   'text-white',
  secondary: 'border text-slate-700 bg-white hover:bg-slate-50',
  ghost:     'text-slate-500 hover:bg-slate-100',
  danger:    'border text-red-600 bg-white hover:bg-red-50',
};

export function Btn({
  variant = 'primary',
  size = 'md',
  children,
  onClick,
  disabled,
  fullWidth,
  type = 'button',
}: {
  variant?: BtnVariant;
  size?: keyof typeof btnSizes;
  children: React.ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  fullWidth?: boolean;
  type?: 'button' | 'submit';
}) {
  const inlineStyle =
    variant === 'primary'
      ? { background: 'linear-gradient(135deg,#f97316,#ea580c)', boxShadow: '0 2px 8px rgba(249,115,22,0.35)' }
      : variant === 'secondary'
      ? { borderColor: '#e2e8f0' }
      : variant === 'danger'
      ? { borderColor: '#fecaca' }
      : {};
  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      className={`${btnBase} ${btnStyles[variant]} ${btnSizes[size]} ${fullWidth ? 'w-full' : ''}`}
      style={inlineStyle}
      onMouseEnter={e => { if (variant === 'primary' && !disabled) (e.currentTarget as HTMLElement).style.opacity = '0.9'; }}
      onMouseLeave={e => { if (variant === 'primary') (e.currentTarget as HTMLElement).style.opacity = '1'; }}
    >
      {children}
    </button>
  );
}

/* ─── PageHeader ─────────────────────────────────────────────────────────── */
export function PageHeader({
  title,
  description,
  crumbs,
}: {
  title: string;
  description?: string;
  crumbs?: string[];
}) {
  return (
    <div className="mb-6">
      {crumbs && (
        <div className="flex items-center gap-1 text-xs mb-1.5" style={{ color: '#94a3b8' }}>
          {crumbs.map((c, i) => (
            <React.Fragment key={i}>
              {i > 0 && <ChevronRight size={11} />}
              <span className={i === crumbs.length - 1 ? 'font-medium' : ''} style={i === crumbs.length - 1 ? { color: '#475569' } : {}}>{c}</span>
            </React.Fragment>
          ))}
        </div>
      )}
      <h2 className="text-xl font-bold leading-tight" style={{ color: '#1e293b' }}>{title}</h2>
      {description && <p className="text-sm mt-1" style={{ color: '#64748b' }}>{description}</p>}
    </div>
  );
}

/* ─── THead ──────────────────────────────────────────────────────────────── */
export function THead({ cols }: { cols: string[] }) {
  return (
    <thead>
      <tr style={{ background: '#f8fafc', borderBottom: '1px solid #f1f5f9' }}>
        {cols.map(c => (
          <th
            key={c}
            className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest"
            style={{ color: '#94a3b8', letterSpacing: '0.07em' }}
          >
            {c}
          </th>
        ))}
      </tr>
    </thead>
  );
}

/* ─── TRow ───────────────────────────────────────────────────────────────── */
export function TRow({ children }: { children: React.ReactNode }) {
  return (
    <tr
      className="transition-colors"
      style={{ borderBottom: '1px solid #f8fafc' }}
      onMouseEnter={e => (e.currentTarget.style.background = '#fafbfc')}
      onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
    >
      {children}
    </tr>
  );
}

export function TD({
  children,
  mono,
  muted,
}: {
  children: React.ReactNode;
  mono?: boolean;
  muted?: boolean;
}) {
  return (
    <td
      className={`px-4 py-3 text-sm ${mono ? 'font-mono text-xs' : ''}`}
      style={{ color: muted ? '#64748b' : '#1e293b' }}
    >
      {children}
    </td>
  );
}

/* ─── UploadZone ─────────────────────────────────────────────────────────── */
export function UploadZone({
  label,
  hint,
  file,
  dragging,
  onDragOver,
  onDragLeave,
  onDrop,
  onClick,
  icon: Icon,
}: {
  label: string;
  hint?: string;
  file?: File | null;
  dragging?: boolean;
  onDragOver?: (e: React.DragEvent) => void;
  onDragLeave?: () => void;
  onDrop?: (e: React.DragEvent) => void;
  onClick?: () => void;
  icon?: React.ElementType;
}) {
  const active = dragging || !!file;
  return (
    <div
      className="relative rounded-xl p-6 text-center cursor-pointer transition-all select-none"
      style={{
        border: `2px dashed ${active ? '#f97316' : '#e2e8f0'}`,
        background: active ? '#fff7ed' : '#fafbfc',
      }}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
      onClick={onClick}
    >
      {Icon && (
        <div
          className="w-10 h-10 rounded-xl flex items-center justify-center mx-auto mb-3"
          style={{ background: active ? '#fed7aa' : '#f1f5f9' }}
        >
          <Icon size={20} style={{ color: active ? '#f97316' : '#94a3b8' }} />
        </div>
      )}
      <p className="text-sm font-semibold" style={{ color: active ? '#c2410c' : '#475569' }}>{file ? file.name : label}</p>
      {hint && !file && <p className="text-xs mt-1" style={{ color: '#94a3b8' }}>{hint}</p>}
    </div>
  );
}
