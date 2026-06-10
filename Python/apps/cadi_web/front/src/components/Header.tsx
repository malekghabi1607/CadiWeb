/**
 * CADI Web — COMPOSANT HEADER
 * 
 * Version simplifiée et épurée :
 * - Affiche le titre de l'écran courant à gauche.
 * - Propose les boutons d'actions rapides et raccourcis à droite.
 * - Supprime les badges textuels institutionnels doublons pour aérer l'interface.
 */

import type { ElementType } from 'react';
import { BarChart2, BookOpen, Database, FileText, HelpCircle, Home, Mail, Scale, Settings, ShieldCheck } from 'lucide-react';
import { Screen } from './Sidebar';

interface HeaderProps {
  sidebarWidth: number;
  title?: string;
  onNavigate?: (screen: Screen) => void;
}

export default function Header({ sidebarWidth, title }: HeaderProps) {
  const titleIcons: Record<string, ElementType> = {
    'IRIS': Database,
    'EvalStat': BarChart2,
    'Bilans': FileText,
    'Formations': BookOpen,
    'Paramètres': Settings,
    'Messagerie': Mail,
    'Tableau de bord': Home,
    'Conditions': FileText,
    'Confidentialite': ShieldCheck,
    'Mentions legales': Scale,
    'FAQ': HelpCircle,
    'Contact': Mail,
  };
  const TitleIcon = title ? titleIcons[title] : undefined;

  return (
    <header
      className="fixed top-0 right-0 z-20 flex items-center justify-between px-6 transition-all duration-200"
      style={{
        left: sidebarWidth,
        height: 76,
        background: '#ffffff',
        borderBottom: '1px solid #e8edf2',
        boxShadow: '0 1px 4px rgba(30,42,74,0.07)',
      }}
    >
      {/* Zone Gauche : Titre dynamique */}
      <div className="flex items-center gap-4 animate-fadeIn min-w-0">
        {TitleIcon && (
          <span className="w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0" style={{ background: '#fff7ed' }}>
            <TitleIcon size={21} style={{ color: '#f97316' }} strokeWidth={2} />
          </span>
        )}
        {title && (
          <span className="text-xl font-black tracking-wide uppercase text-slate-800">
            {title}
          </span>
        )}
      </div>

      {/* Zone Droite : Logo CADI */}
      <div className="flex items-center justify-end flex-shrink-0">
        <img
          src="/image.png"
          alt="CADI"
          className="flex-shrink-0"
          style={{ height: 72, width: 230, objectFit: 'contain' }}
        />
      </div>
    </header>
  );
}
