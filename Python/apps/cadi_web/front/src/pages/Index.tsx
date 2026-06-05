import { Database, BarChart2, FileText, BookOpen, ArrowRight, Shield, Zap, Users } from 'lucide-react';

const modules = [
  {
    Icon: Database,
    color: '#3b82f6',
    bg: '#eff6ff',
    label: 'IRIS',
    desc: 'Mise à jour automatique des exports GED et traitement des dumps IRIS en quelques clics.',
  },
  {
    Icon: BarChart2,
    color: '#16a34a',
    bg: '#f0fdf4',
    label: 'EvalStat',
    desc: 'Consolidation des évaluations stagiaires et génération des fichiers Excel de résultats.',
  },
  {
    Icon: FileText,
    color: '#f97316',
    bg: '#fff7ed',
    label: 'Bilans',
    desc: 'Génération automatique des bilans de sessions et de formations au format Word.',
  },
  {
    Icon: BookOpen,
    color: '#7c3aed',
    bg: '#f5f3ff',
    label: 'Formations',
    desc: 'Consultation et gestion du catalogue complet des formations INSTN synchronisé.',
  },
];

const features = [
  { Icon: Zap,    label: 'Rapide',    desc: 'Traitements automatisés en quelques secondes' },
  { Icon: Shield, label: 'Sécurisé', desc: 'Accès contrôlé par identifiants CEA' },
  { Icon: Users,  label: 'Multi-utilisateurs', desc: 'Gestion des rôles et des droits d\'accès' },
];

interface IndexProps {
  onLogin: () => void;
  onRegister: () => void;
}

export default function Index({ onLogin, onRegister }: IndexProps) {
  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#f0f3f8' }}>

      {/* ── Nav bar ── */}
      <nav
        className="flex items-center justify-between px-8 py-4 sticky top-0 z-10"
        style={{ background: 'rgba(255,255,255,0.92)', backdropFilter: 'blur(12px)', borderBottom: '1px solid #e8edf2', boxShadow: '0 1px 4px rgba(30,42,74,0.06)' }}
      >
        <div className="flex items-center gap-4">
          <img src="/image.png" alt="CADI" style={{ height: 64, width: 160, objectFit: 'contain' }} />
          <div>
            <span className="font-bold text-sm" style={{ color: '#1e2a4a' }}>CADI Web</span>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={onLogin}
            className="px-4 py-2 rounded-xl text-sm font-semibold transition-colors"
            style={{ color: '#1e2a4a', background: 'transparent' }}
            onMouseEnter={e => (e.currentTarget.style.background = '#f1f5f9')}
            onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
          >
            Se connecter
          </button>
          <button
            onClick={onRegister}
            className="px-4 py-2 rounded-xl text-sm font-bold transition-all"
            style={{ background: '#1e2a4a', color: '#fff', boxShadow: '0 2px 8px rgba(30,42,74,0.25)' }}
            onMouseEnter={e => (e.currentTarget.style.background = '#162036')}
            onMouseLeave={e => (e.currentTarget.style.background = '#1e2a4a')}
          >
            S'inscrire
          </button>
        </div>
      </nav>

      {/* ── Hero ── */}
      <section className="flex flex-col items-center text-center px-8 pt-20 pb-16">
        <img
          src="/image.png"
          alt="CADI"
          className="mb-10"
          style={{ height: 150, width: 'min(520px, 90vw)', objectFit: 'contain' }}
        />
        <div
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-semibold mb-6"
          style={{ background: '#fff7ed', color: '#f97316', border: '1px solid #fed7aa' }}
        >
          <span className="w-1.5 h-1.5 rounded-full bg-orange-400" />
          INSTN / CEA — Outil interne de gestion
        </div>
        <h1 className="text-5xl font-black leading-tight mb-5 max-w-2xl" style={{ color: '#1e293b' }}>
          Simplifiez la gestion<br />
          de vos <span style={{ color: '#f97316' }}>formations INSTN</span>
        </h1>
        <p className="text-lg max-w-xl leading-relaxed mb-10" style={{ color: '#64748b' }}>
          CADI Web centralise et automatise le traitement des données IRIS, évaluations stagiaires, bilans de formations et fiches de coûts.
        </p>
        <div className="flex items-center gap-4">
          <button
            onClick={onLogin}
            className="flex items-center gap-2 px-6 py-3.5 rounded-xl text-sm font-bold transition-all"
            style={{ background: '#1e2a4a', color: '#fff', boxShadow: '0 4px 16px rgba(30,42,74,0.3)' }}
            onMouseEnter={e => (e.currentTarget.style.background = '#162036')}
            onMouseLeave={e => (e.currentTarget.style.background = '#1e2a4a')}
          >
            Accéder à l'application <ArrowRight size={16} />
          </button>
          <button
            onClick={onRegister}
            className="px-6 py-3.5 rounded-xl text-sm font-semibold transition-colors"
            style={{ color: '#475569', background: '#fff', border: '1px solid #e2e8f0' }}
            onMouseEnter={e => { (e.currentTarget.style.background = '#f8fafc'); }}
            onMouseLeave={e => { (e.currentTarget.style.background = '#fff'); }}
          >
            Créer un compte
          </button>
        </div>

        {/* Feature pills */}
        <div className="flex items-center gap-6 mt-10">
          {features.map(({ Icon, label, desc }) => (
            <div key={label} className="flex items-center gap-2.5">
              <div className="w-7 h-7 rounded-lg flex items-center justify-center flex-shrink-0" style={{ background: '#f1f5f9' }}>
                <Icon size={14} style={{ color: '#475569' }} />
              </div>
              <div className="text-left">
                <p className="text-xs font-bold" style={{ color: '#1e293b' }}>{label}</p>
                <p className="text-xs" style={{ color: '#94a3b8' }}>{desc}</p>
              </div>
            </div>
          ))}
        </div>
      </section>

      {/* ── Modules ── */}
      <section className="px-8 pb-20 max-w-5xl mx-auto w-full">
        <div className="text-center mb-10">
          <h2 className="text-2xl font-black mb-2" style={{ color: '#1e293b' }}>4 modules intégrés</h2>
          <p className="text-sm" style={{ color: '#64748b' }}>Tout ce dont vous avez besoin pour gérer le cycle complet d'une formation</p>
        </div>
        <div className="grid grid-cols-2 gap-4">
          {modules.map(({ Icon, color, bg, label, desc }) => (
            <div
              key={label}
              className="flex items-start gap-4 p-6 rounded-2xl transition-all"
              style={{ background: '#fff', border: '1px solid #e8edf2', boxShadow: '0 1px 4px rgba(30,42,74,0.05)' }}
              onMouseEnter={e => {
                (e.currentTarget as HTMLElement).style.borderColor = color;
                (e.currentTarget as HTMLElement).style.boxShadow = `0 4px 20px ${color}18`;
                (e.currentTarget as HTMLElement).style.transform = 'translateY(-2px)';
              }}
              onMouseLeave={e => {
                (e.currentTarget as HTMLElement).style.borderColor = '#e8edf2';
                (e.currentTarget as HTMLElement).style.boxShadow = '0 1px 4px rgba(30,42,74,0.05)';
                (e.currentTarget as HTMLElement).style.transform = 'translateY(0)';
              }}
            >
              <div className="w-12 h-12 rounded-xl flex items-center justify-center flex-shrink-0" style={{ background: bg }}>
                <Icon size={22} style={{ color }} strokeWidth={1.8} />
              </div>
              <div>
                <p className="font-bold text-base mb-1" style={{ color: '#1e293b' }}>{label}</p>
                <p className="text-sm leading-relaxed" style={{ color: '#64748b' }}>{desc}</p>
              </div>
            </div>
          ))}
        </div>
      </section>

      {/* ── CTA band ── */}
      <section
        className="mx-8 mb-16 rounded-2xl px-10 py-10 flex items-center justify-between"
        style={{ background: '#1e2a4a', boxShadow: '0 8px 32px rgba(30,42,74,0.25)' }}
      >
        <div>
          <h3 className="text-xl font-black text-white mb-1.5">Prêt à commencer ?</h3>
          <p className="text-sm" style={{ color: 'rgba(255,255,255,0.55)' }}>Connectez-vous avec vos identifiants CEA ou créez un nouveau compte.</p>
        </div>
        <div className="flex items-center gap-3 flex-shrink-0">
          <button
            onClick={onRegister}
            className="px-5 py-2.5 rounded-xl text-sm font-semibold transition-colors"
            style={{ background: 'rgba(255,255,255,0.1)', color: '#fff', border: '1px solid rgba(255,255,255,0.15)' }}
            onMouseEnter={e => (e.currentTarget.style.background = 'rgba(255,255,255,0.18)')}
            onMouseLeave={e => (e.currentTarget.style.background = 'rgba(255,255,255,0.1)')}
          >
            S'inscrire
          </button>
          <button
            onClick={onLogin}
            className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-bold transition-all"
            style={{ background: '#f97316', color: '#fff', boxShadow: '0 4px 12px rgba(249,115,22,0.4)' }}
            onMouseEnter={e => (e.currentTarget.style.background = '#ea580c')}
            onMouseLeave={e => (e.currentTarget.style.background = '#f97316')}
          >
            Se connecter <ArrowRight size={14} />
          </button>
        </div>
      </section>

      {/* ── Footer ── */}
      <footer
        className="text-center py-6 text-xs"
        style={{ color: '#94a3b8', borderTop: '1px solid #e8edf2' }}
      >
        CADI Web · INSTN / CEA · DEN/SFEN · Saclay · {new Date().getFullYear()}
      </footer>
    </div>
  );
}
