import { useState } from 'react';
import { Eye, EyeOff, Lock, Mail, ArrowRight } from 'lucide-react';

interface LoginProps {
  onLogin: (user: { prenom: string; nom: string; email: string; unite?: string }) => void;
  onRegister: () => void;
  onForgot: () => void;
}

export default function Login({ onLogin, onRegister, onForgot }: LoginProps) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    const trimmedEmail = email.trim().toLowerCase();

    setLoading(true);

    // Connexion locale temporaire pour continuer les tests de l'interface.
    // Elle sera remplacee plus tard par la verification complete du compte.
    setTimeout(() => {
      setLoading(false);
      onLogin({
        prenom: 'Utilisateur',
        nom: 'CADI',
        email: trimmedEmail || 'visiteur@cadi.local',
        unite: 'INSTN',
      });
    }, 300);
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#f0f3f8' }}>
      {/* Top bar */}
      <div className="flex items-center justify-between px-8 py-4">
        <img src="/image.png" alt="CADI" style={{ height: 64, width: 160, objectFit: 'contain' }} />
        <span className="text-xs font-medium" style={{ color: '#94a3b8' }}>INSTN / CEA</span>
      </div>

      {/* Card */}
      <div className="flex-1 flex items-center justify-center px-4 pb-16">
        <div className="w-full max-w-sm">

          {/* Header */}
          <div className="text-center mb-8">
            <div
              className="w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-5"
              style={{ background: '#1e2a4a' }}
            >
              <Lock size={22} style={{ color: '#f97316' }} />
            </div>
            <h1 className="text-2xl font-black mb-1.5" style={{ color: '#1e293b' }}>Connexion</h1>
            <p className="text-sm" style={{ color: '#64748b' }}>Accédez à votre espace CADI Web</p>
          </div>

          {/* Form */}
          <form onSubmit={submit} className="space-y-4">
            {error && (
              <div className="px-4 py-3 rounded-xl text-sm font-medium" style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#dc2626' }}>
                {error}
              </div>
            )}

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>
                Adresse mail
              </label>
              <div className="relative">
                <Mail size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                <input
                  type="email"
                  value={email}
                  onChange={e => setEmail(e.target.value)}
                  placeholder="votre.nom@cea.fr"
                  className="w-full pl-10 pr-4 py-3 rounded-xl text-sm outline-none transition-all"
                  style={{ background: '#fff', border: '1px solid #e2e8f0', color: '#1e293b' }}
                  onFocus={e => { e.target.style.borderColor = '#f97316'; e.target.style.boxShadow = '0 0 0 3px rgba(249,115,22,0.12)'; }}
                  onBlur={e => { e.target.style.borderColor = '#e2e8f0'; e.target.style.boxShadow = 'none'; }}
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>
                Mot de passe
              </label>
              <div className="relative">
                <Lock size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                <input
                  type={showPw ? 'text' : 'password'}
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full pl-10 pr-11 py-3 rounded-xl text-sm outline-none transition-all"
                  style={{ background: '#fff', border: '1px solid #e2e8f0', color: '#1e293b' }}
                  onFocus={e => { e.target.style.borderColor = '#f97316'; e.target.style.boxShadow = '0 0 0 3px rgba(249,115,22,0.12)'; }}
                  onBlur={e => { e.target.style.borderColor = '#e2e8f0'; e.target.style.boxShadow = 'none'; }}
                />
                <button
                  type="button"
                  onClick={() => setShowPw(v => !v)}
                  className="absolute right-3.5 top-1/2 -translate-y-1/2 transition-colors"
                  style={{ color: '#94a3b8' }}
                  onMouseEnter={e => (e.currentTarget.style.color = '#475569')}
                  onMouseLeave={e => (e.currentTarget.style.color = '#94a3b8')}
                >
                  {showPw ? <EyeOff size={15} /> : <Eye size={15} />}
                </button>
              </div>
              <div className="flex justify-end mt-1.5">
                <button
                  type="button"
                  onClick={onForgot}
                  className="text-xs font-medium transition-colors"
                  style={{ color: '#f97316' }}
                  onMouseEnter={e => (e.currentTarget.style.color = '#ea580c')}
                  onMouseLeave={e => (e.currentTarget.style.color = '#f97316')}
                >
                  Mot de passe oublié ?
                </button>
              </div>
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full flex items-center justify-center gap-2 py-3.5 rounded-xl text-sm font-bold transition-all mt-2"
              style={{ background: loading ? '#94a3b8' : '#1e2a4a', color: '#fff', boxShadow: loading ? 'none' : '0 4px 12px rgba(30,42,74,0.3)' }}
              onMouseEnter={e => { if (!loading) (e.currentTarget.style.background = '#162036'); }}
              onMouseLeave={e => { if (!loading) (e.currentTarget.style.background = '#1e2a4a'); }}
            >
              {loading ? 'Connexion…' : <><span>Se connecter</span><ArrowRight size={15} /></>}
            </button>
          </form>

          {/* Footer link */}
          <p className="text-center text-sm mt-6" style={{ color: '#94a3b8' }}>
            Pas encore de compte ?{' '}
            <button
              onClick={onRegister}
              className="font-semibold transition-colors"
              style={{ color: '#f97316' }}
              onMouseEnter={e => (e.currentTarget.style.color = '#ea580c')}
              onMouseLeave={e => (e.currentTarget.style.color = '#f97316')}
            >
              S'inscrire
            </button>
          </p>
        </div>
      </div>

      <footer className="text-center py-4 text-xs" style={{ color: '#cbd5e1' }}>
        CADI Web · INSTN / CEA
      </footer>
    </div>
  );
}
