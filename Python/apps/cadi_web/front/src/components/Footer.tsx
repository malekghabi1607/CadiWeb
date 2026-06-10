/**
 * CADI Web - Footer dashboard compact.
 *
 * Pas de navigation metier ici : la sidebar gere deja le dashboard.
 * Le footer donne seulement les liens d'aide et les informations internes.
 */

import { FileText, HelpCircle, Mail, Scale, ShieldCheck } from 'lucide-react';
import type { InfoPage } from '../pages/InfoPages';

interface FooterProps {
  sidebarWidth: number;
  onNavigateInfo: (page: InfoPage) => void;
}

const links: { page: InfoPage; label: string; Icon: React.ElementType }[] = [
  { page: 'terms', label: 'Conditions', Icon: FileText },
  { page: 'privacy', label: 'Confidentialite', Icon: ShieldCheck },
  { page: 'legal', label: 'Mentions legales', Icon: Scale },
  { page: 'faq', label: 'FAQ', Icon: HelpCircle },
  { page: 'contact', label: 'Contact', Icon: Mail },
];

export default function Footer({ sidebarWidth, onNavigateInfo }: FooterProps) {
  return (
    <footer
      className="flex items-center justify-between gap-4 px-5 transition-all duration-200"
      style={{
        marginLeft: sidebarWidth,
        borderTop: '1px solid #e8edf2',
        background: '#ffffff',
        minHeight: 48,
        flexShrink: 0,
      }}
    >
      <div className="flex min-w-0 items-center gap-2">
        <span className="h-2 w-2 rounded-full" style={{ background: '#f97316' }} />
        <span className="truncate text-xs font-bold" style={{ color: '#64748b' }}>
          CADI Web - INSTN / CEA
        </span>
      </div>

      <nav className="flex items-center gap-1.5" aria-label="Liens d'aide et mentions">
        {links.map(({ page, label, Icon }) => (
          <button
            key={page}
            type="button"
            onClick={() => onNavigateInfo(page)}
            className="flex h-8 w-8 items-center justify-center rounded-md transition-colors hover:bg-orange-50 hover:text-orange-500"
            style={{ color: '#94a3b8' }}
            title={label}
            aria-label={label}
          >
            <Icon size={16} strokeWidth={1.9} />
          </button>
        ))}
      </nav>

      <div className="hidden items-center gap-1.5 text-xs font-semibold sm:flex" style={{ color: '#94a3b8' }}>
        <ShieldCheck size={14} />
        Usage interne
      </div>
    </footer>
  );
}
