/**
 * CADI Web — COMPOSANT HEADER
 * 
 * Version simplifiée et épurée :
 * - Affiche le titre de l'écran courant à gauche.
 * - Propose les boutons d'actions rapides et raccourcis à droite.
 * - Supprime les badges textuels institutionnels doublons pour aérer l'interface.
 */

import { Bell, Home, Mail, Settings } from 'lucide-react';
import { Screen } from './Sidebar';

interface HeaderProps {
  sidebarWidth: number;
  title?: string;
  onNavigate?: (screen: Screen) => void;
}

export default function Header({ sidebarWidth, title, onNavigate }: HeaderProps) {
  return (
    <header
      className="fixed top-0 right-0 z-20 flex items-center justify-between px-6 transition-all duration-200"
      style={{
        left: sidebarWidth,
        height: 64,
        background: '#ffffff',
        borderBottom: '1px solid #e8edf2',
        boxShadow: '0 1px 4px rgba(30,42,74,0.07)',
      }}
    >
      {/* Zone Gauche : Titre dynamique */}
      <div className="flex items-center gap-3 animate-fadeIn">
        <div className="w-1.5 h-4.5 rounded-full" style={{ background: '#f97316' }} />
        {title && (
          <span className="text-sm font-extrabold tracking-wide uppercase text-slate-800">
            {title}
          </span>
        )}
      </div>

      {/* Zone Droite : Boutons Raccourcis Directs */}
      <div className="flex items-center gap-2">
        {/* Bouton Raccourci : Accueil */}
        <button
          onClick={() => onNavigate?.('accueil')}
          className="w-9 h-9 rounded-xl flex items-center justify-center transition-colors text-slate-400 hover:text-slate-800 hover:bg-slate-50"
          title="Tableau de bord"
        >
          <Home size={17} strokeWidth={1.8} />
        </button>

        {/* Bouton Raccourci : Messagerie */}
        <button
          onClick={() => onNavigate?.('messagerie')}
          className="w-9 h-9 rounded-xl flex items-center justify-center transition-colors text-slate-400 hover:text-slate-800 hover:bg-slate-50"
          title="Messagerie (Historique)"
        >
          <Mail size={17} strokeWidth={1.8} />
        </button>

        {/* Bouton Raccourci : Paramètres */}
        <button
          onClick={() => onNavigate?.('parametres')}
          className="w-9 h-9 rounded-xl flex items-center justify-center transition-colors text-slate-400 hover:text-slate-800 hover:bg-slate-50"
          title="Paramètres"
        >
          <Settings size={17} strokeWidth={1.8} />
        </button>

        {/* Séparateur visuel vertical */}
        <div className="w-px h-5 mx-1" style={{ background: '#e2e8f0' }} />

        {/* Bouton Notifications */}
        <button
          className="relative w-9 h-9 rounded-xl flex items-center justify-center transition-colors text-slate-400 hover:text-slate-800 hover:bg-slate-50"
          title="Notifications"
        >
          <Bell size={17} strokeWidth={1.8} />
          <span
            className="absolute top-1.5 right-1.5 w-2 h-2 rounded-full border-2 border-white bg-orange-500"
          />
        </button>
      </div>
    </header>
  );
}
