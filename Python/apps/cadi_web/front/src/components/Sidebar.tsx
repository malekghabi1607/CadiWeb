/**
 * CADI Web — COMPOSANT SIDEBAR
 * 
 * Version nettoyée et optimisée :
 * - Logo agrandi (hauteur 52px).
 * - Retrait de l'onglet autonome "Fiches de coûts" (géré via des popups contextuelles dans l'écran Bilans).
 * - Navigation ultra-claire et épurée.
 */

import {
  Home,
  Database,
  BarChart2,
  FileText,
  BookOpen,
  Settings,
  Mail,
  ChevronRight,
  LogOut,
  User,
} from 'lucide-react';

export type Screen =
  | 'accueil'
  | 'iris'
  | 'evalstat'
  | 'bilans'
  | 'formations'
  | 'parametres'
  | 'messagerie';

interface SidebarProps {
  active: Screen;
  onNavigate: (screen: Screen) => void;
  collapsed: boolean;
  onToggle: () => void;
  onLogout: () => void;
  currentUser?: { prenom: string; nom: string; email: string; unite?: string } | null;
}

const navItems: { id: Screen; label: string; Icon: React.ElementType; badge?: string }[] = [
  { id: 'accueil',    label: 'Accueil',         Icon: Home },
  { id: 'iris',       label: 'IRIS',            Icon: Database,  badge: '3' },
  { id: 'evalstat',   label: 'EvalStat',        Icon: BarChart2 },
  { id: 'bilans',     label: 'Bilans',          Icon: FileText },
  { id: 'formations', label: 'Formations',      Icon: BookOpen },
  { id: 'parametres', label: 'Paramètres',      Icon: Settings },
  { id: 'messagerie', label: 'Messagerie',      Icon: Mail,      badge: '1' },
];

export const SIDEBAR_W = 220;
export const SIDEBAR_W_COLLAPSED = 64;

export default function Sidebar({ active, onNavigate, collapsed, onToggle, onLogout, currentUser }: SidebarProps) {
  const w = collapsed ? SIDEBAR_W_COLLAPSED : SIDEBAR_W;

  return (
    <aside
      className="fixed left-0 top-0 h-full flex flex-col z-30 select-none transition-all duration-200"
      style={{ width: w, background: '#1e2a4a', overflow: 'hidden' }}
    >
      {/* ── Section supérieure : Hamburger + Logo agrandi ── */}
      <div
        className="flex items-center px-3 flex-shrink-0"
        style={{ borderBottom: '1px solid rgba(255,255,255,0.08)', minHeight: 64, gap: 10 }}
      >
        <button
          onClick={onToggle}
          className="flex-shrink-0 w-8 h-8 rounded-xl flex items-center justify-center transition-colors"
          style={{ background: 'rgba(255,255,255,0.07)', color: 'rgba(255,255,255,0.7)' }}
          onMouseEnter={e => {
            (e.currentTarget as HTMLElement).style.background = 'rgba(255,255,255,0.15)';
            (e.currentTarget as HTMLElement).style.color = '#fff';
          }}
          onMouseLeave={e => {
            (e.currentTarget as HTMLElement).style.background = 'rgba(255,255,255,0.07)';
            (e.currentTarget as HTMLElement).style.color = 'rgba(255,255,255,0.7)';
          }}
          title={collapsed ? 'Déplier' : 'Réduire'}
        >
          <svg width="16" height="12" viewBox="0 0 16 12" fill="none">
            <rect y="0"  width="16" height="2" rx="1" fill="currentColor"/>
            <rect y="5"  width="11" height="2" rx="1" fill="currentColor"/>
            <rect y="10" width="16" height="2" rx="1" fill="currentColor"/>
          </svg>
        </button>

        {!collapsed && (
          <div
            className="flex items-center gap-3 cursor-pointer flex-1 min-w-0"
            onClick={() => onNavigate('accueil')}
          >
            <img
              src="/image.png"
              alt="CADI"
              style={{ height: 48, objectFit: 'contain', filter: 'brightness(0) invert(1)', flexShrink: 0 }}
            />
            <div className="min-w-0">
              <div className="text-white font-extrabold text-sm leading-tight tracking-wide whitespace-nowrap">CADI Web</div>
              <div className="text-xs font-semibold whitespace-nowrap" style={{ color: 'rgba(255,255,255,0.45)' }}>INSTN / CEA</div>
            </div>
          </div>
        )}
      </div>

      {/* ── Section centrale : Navigation ── */}
      <nav className="flex-1 overflow-y-auto py-4 px-2">
        {!collapsed && (
          <p
            className="text-xs font-semibold uppercase tracking-widest px-2 mb-2"
            style={{ color: 'rgba(255,255,255,0.28)', letterSpacing: '0.1em' }}
          >
            Navigation
          </p>
        )}
        <ul className="space-y-0.5">
          {navItems.map(({ id, label, Icon, badge }) => {
            const isActive = active === id;
            return (
              <li key={id}>
                <button
                  onClick={() => onNavigate(id)}
                  title={collapsed ? label : undefined}
                  className="w-full flex items-center rounded-xl text-sm font-medium transition-all duration-150 relative group"
                  style={{
                    gap: collapsed ? 0 : 10,
                    padding: collapsed ? '9px 0' : '9px 10px',
                    justifyContent: collapsed ? 'center' : 'flex-start',
                    background: isActive ? 'rgba(249,115,22,0.18)' : 'transparent',
                    color: isActive ? '#fdba74' : 'rgba(255,255,255,0.58)',
                  }}
                  onMouseEnter={e => {
                    if (!isActive) {
                      (e.currentTarget as HTMLElement).style.background = 'rgba(255,255,255,0.06)';
                      (e.currentTarget as HTMLElement).style.color = 'rgba(255,255,255,0.9)';
                    }
                  }}
                  onMouseLeave={e => {
                    if (!isActive) {
                      (e.currentTarget as HTMLElement).style.background = 'transparent';
                      (e.currentTarget as HTMLElement).style.color = 'rgba(255,255,255,0.58)';
                    }
                  }}
                >
                  {isActive && (
                    <span
                      className="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-5 rounded-r-full"
                      style={{ background: '#f97316' }}
                    />
                  )}
                  <Icon
                    size={17}
                    strokeWidth={isActive ? 2.2 : 1.7}
                    style={{ color: isActive ? '#f97316' : 'inherit', flexShrink: 0 }}
                  />
                  {!collapsed && (
                    <>
                      <span className="flex-1 text-left">{label}</span>
                      {badge && !isActive && (
                        <span
                          className="flex items-center justify-center rounded-full font-bold flex-shrink-0"
                          style={{ background: '#f97316', color: '#fff', fontSize: 10, width: 18, height: 18 }}
                        >
                          {badge}
                        </span>
                      )}
                      {isActive && <ChevronRight size={12} style={{ color: '#f97316', flexShrink: 0 }} />}
                    </>
                  )}
                  {collapsed && badge && (
                    <span
                      className="absolute top-1.5 right-1.5 w-2 h-2 rounded-full border border-slate-800"
                      style={{ background: '#f97316' }}
                    />
                  )}
                </button>
              </li>
            );
          })}
        </ul>
      </nav>

      {/* ── Section basse : Utilisateur ── */}
      <div
        className="flex-shrink-0 px-2 py-3"
        style={{ borderTop: '1px solid rgba(255,255,255,0.08)' }}
      >
        {!collapsed && (
          <div className="flex items-center gap-2.5 px-2 py-2 mb-1">
            <div
              className="w-7 h-7 rounded-full flex items-center justify-center flex-shrink-0"
              style={{ background: 'rgba(255,255,255,0.12)' }}
            >
              <User size={14} style={{ color: 'rgba(255,255,255,0.8)' }} />
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-semibold truncate" style={{ color: 'rgba(255,255,255,0.85)' }}>
                {currentUser ? `${currentUser.prenom} ${currentUser.nom}` : 'Martin Dupont'}
              </p>
              <p className="text-xs truncate" style={{ color: 'rgba(255,255,255,0.38)' }}>
                {currentUser?.email === 'admin@cea.fr' ? 'Administrateur' : (currentUser?.unite || 'Utilisateur')}
              </p>
            </div>
          </div>
        )}

        <button
          onClick={onLogout}
          title="Se déconnecter"
          className="w-full flex items-center rounded-xl text-sm font-medium transition-all duration-150"
          style={{
            gap: collapsed ? 0 : 10,
            padding: collapsed ? '9px 0' : '9px 10px',
            justifyContent: collapsed ? 'center' : 'flex-start',
            color: 'rgba(255,100,100,0.7)',
          }}
          onMouseEnter={e => {
            (e.currentTarget as HTMLElement).style.background = 'rgba(220,38,38,0.15)';
            (e.currentTarget as HTMLElement).style.color = '#fca5a5';
          }}
          onMouseLeave={e => {
            (e.currentTarget as HTMLElement).style.background = 'transparent';
            (e.currentTarget as HTMLElement).style.color = 'rgba(255,100,100,0.7)';
          }}
        >
          <LogOut size={16} strokeWidth={1.8} style={{ flexShrink: 0 }} />
          {!collapsed && <span>Se déconnecter</span>}
        </button>
      </div>
    </aside>
  );
}
