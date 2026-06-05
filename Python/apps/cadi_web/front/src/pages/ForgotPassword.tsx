import { useState } from 'react';
import { Mail, ArrowLeft, ArrowRight, CheckCircle } from 'lucide-react';

interface ForgotPasswordProps {
  onBack: () => void;
}

export default function ForgotPassword({ onBack }: ForgotPasswordProps) {
  const [email, setEmail] = useState('');
  const [sent, setSent] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!email) { setError('Veuillez saisir votre adresse mail.'); return; }
    setError('');
    setLoading(true);
    setTimeout(() => { setLoading(false); setSent(true); }, 900);
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#f0f3f8' }}>
      <div className="flex items-center justify-between px-8 py-4">
        <img src="/image.png" alt="CADI" style={{ height: 64, width: 160, objectFit: 'contain' }} />
        <span className="text-xs font-medium" style={{ color: '#94a3b8' }}>INSTN / CEA</span>
      </div>

      <div className="flex-1 flex items-center justify-center px-4 pb-16">
        <div className="w-full max-w-sm">

          {!sent ? (
            <>
              <div className="text-center mb-8">
                <div className="w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-5" style={{ background: '#1e2a4a' }}>
                  <Mail size={22} style={{ color: '#f97316' }} />
                </div>
                <h1 className="text-2xl font-black mb-1.5" style={{ color: '#1e293b' }}>Mot de passe oublié</h1>
                <p className="text-sm leading-relaxed" style={{ color: '#64748b' }}>
                  Saisissez votre adresse mail CEA. Nous vous enverrons un lien de réinitialisation.
                </p>
              </div>

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

                <button
                  type="submit"
                  disabled={loading}
                  className="w-full flex items-center justify-center gap-2 py-3.5 rounded-xl text-sm font-bold transition-all"
                  style={{ background: loading ? '#94a3b8' : '#1e2a4a', color: '#fff', boxShadow: loading ? 'none' : '0 4px 12px rgba(30,42,74,0.3)' }}
                  onMouseEnter={e => { if (!loading) (e.currentTarget.style.background = '#162036'); }}
                  onMouseLeave={e => { if (!loading) (e.currentTarget.style.background = '#1e2a4a'); }}
                >
                  {loading ? 'Envoi…' : <><span>Envoyer le lien</span><ArrowRight size={15} /></>}
                </button>
              </form>
            </>
          ) : (
            <div className="text-center">
              <div className="w-16 h-16 rounded-2xl flex items-center justify-center mx-auto mb-5" style={{ background: '#f0fdf4', border: '2px solid #bbf7d0' }}>
                <CheckCircle size={28} className="text-emerald-500" />
              </div>
              <h1 className="text-xl font-black mb-2" style={{ color: '#1e293b' }}>Mail envoyé !</h1>
              <p className="text-sm leading-relaxed mb-6" style={{ color: '#64748b' }}>
                Un lien de réinitialisation a été envoyé à <strong style={{ color: '#1e293b' }}>{email}</strong>. Vérifiez votre boîte de réception.
              </p>
            </div>
          )}

          <p className="text-center text-sm mt-6" style={{ color: '#94a3b8' }}>
            <button onClick={onBack} className="font-semibold inline-flex items-center gap-1 transition-colors" style={{ color: '#f97316' }}
              onMouseEnter={e => (e.currentTarget.style.color = '#ea580c')}
              onMouseLeave={e => (e.currentTarget.style.color = '#f97316')}>
              <ArrowLeft size={12} /> Retour à la connexion
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
