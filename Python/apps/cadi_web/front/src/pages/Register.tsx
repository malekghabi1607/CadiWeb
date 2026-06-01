import { useState } from 'react';
import { Eye, EyeOff, Lock, Mail, User, Building, ArrowRight, ArrowLeft, Hash, ShieldCheck } from 'lucide-react';
import { supabase } from '../lib/supabase';

interface RegisterProps {
  onBack: () => void;
  onSuccess: () => void;
}

export default function Register({ onBack, onSuccess }: RegisterProps) {
  // Step 1 — badge
  const [badge, setBadge] = useState('');
  const [badgeError, setBadgeError] = useState('');
  const [badgeLoading, setBadgeLoading] = useState(false);
  const [step, setStep] = useState<1 | 2>(1);

  // Step 2 — account info
  const [prenom, setPrenom] = useState('');
  const [nom, setNom] = useState('');
  const [email, setEmail] = useState('');
  const [unite, setUnite] = useState('');
  const [password, setPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [showPw, setShowPw] = useState(false);
  const [formError, setFormError] = useState('');
  const [formLoading, setFormLoading] = useState(false);

  // ── Liste des badges valides autorisés en dur dans le code ──
  const ALLOWED_BADGES = ['123456', '234567', '345678', '888888', '999999', '111111', '654321'];

  // ── Step 1: verify badge ──
  const verifyBadge = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = badge.trim();
    if (trimmed.length !== 6 || !/^\d{6}$/.test(trimmed)) {
      setBadgeError('Le numéro de badge doit contenir exactement 6 chiffres.');
      return;
    }

    // Validation stricte en dur par rapport à la liste autorisée du code
    if (!ALLOWED_BADGES.includes(trimmed)) {
      setBadgeError('Badge non autorisé dans le code système CADI.');
      return;
    }

    // Vérification locale si déjà utilisé
    const usedBadgesStr = localStorage.getItem('cadi_used_badges') || '[]';
    const usedBadges = JSON.parse(usedBadgesStr);
    if (usedBadges.includes(trimmed)) {
      setBadgeError('Ce badge a déjà été utilisé pour créer un compte.');
      return;
    }

    setBadgeError('');
    setBadgeLoading(true);
    
    // Essai Supabase (facultatif/non-bloquant en cas d'erreur de base)
    try {
      const { data } = await supabase
        .from('badge_codes')
        .select('used')
        .eq('badge', trimmed)
        .maybeSingle();
      
      if (data?.used) {
        setBadgeError('Ce badge est déjà enregistré dans Supabase.');
        setBadgeLoading(false);
        return;
      }
    } catch {
      // Ignorer l'erreur réseau ou db de Supabase pour la robustesse locale
    }

    setBadgeLoading(false);
    setStep(2);
  };

  // ── Step 2: create account ──
  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setFormError('');
    if (!prenom || !nom || !email || !password) {
      setFormError('Veuillez remplir tous les champs obligatoires.');
      return;
    }
    if (password !== confirm) {
      setFormError('Les mots de passe ne correspondent pas.');
      return;
    }
    if (password.length < 8) {
      setFormError('Le mot de passe doit contenir au moins 8 caractères.');
      return;
    }
    setFormLoading(true);

    try {
      // Inscription Supabase
      const { data: authData } = await supabase.auth.signUp({
        email: email.trim(),
        password,
      });

      if (authData?.user) {
        // Insert profile dans Supabase
        await supabase.from('profiles').insert({
          id: authData.user.id,
          nom: nom.trim(),
          prenom: prenom.trim(),
          unite: unite.trim(),
          badge: badge.trim(),
        });

        // Mark badge as used
        await supabase
          .from('badge_codes')
          .update({ used: true, used_by: email.trim() })
          .eq('badge', badge.trim());
      }
    } catch {
      // Ignorer silencieusement pour le mode local
    }

    // PERSISTANCE LOCALSTORAGE (Toujours actif pour la robustesse demandée)
    const usedBadgesStr = localStorage.getItem('cadi_used_badges') || '[]';
    const usedBadges = JSON.parse(usedBadgesStr);
    usedBadges.push(badge.trim());
    localStorage.setItem('cadi_used_badges', JSON.stringify(usedBadges));

    const registeredUsersStr = localStorage.getItem('cadi_registered_users') || '[]';
    const registeredUsers = JSON.parse(registeredUsersStr);
    registeredUsers.push({
      email: email.trim(),
      password: password,
      prenom: prenom.trim(),
      nom: nom.trim(),
      unite: unite.trim()
    });
    localStorage.setItem('cadi_registered_users', JSON.stringify(registeredUsers));

    setFormLoading(false);
    onSuccess();
  };

  const inputCls = "w-full px-3.5 py-3 rounded-xl text-sm outline-none transition-all";
  const iStyle = { background: '#fff', border: '1px solid #e2e8f0', color: '#1e293b' };
  const onFocus = (e: React.FocusEvent<HTMLInputElement>) => {
    e.target.style.borderColor = '#f97316';
    e.target.style.boxShadow = '0 0 0 3px rgba(249,115,22,0.12)';
  };
  const onBlur = (e: React.FocusEvent<HTMLInputElement>) => {
    e.target.style.borderColor = '#e2e8f0';
    e.target.style.boxShadow = 'none';
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#f0f3f8' }}>
      {/* Top bar */}
      <div className="flex items-center justify-between px-8 py-4">
        <img src="/image.png" alt="CADI" style={{ height: 36, objectFit: 'contain' }} />
        <span className="text-xs font-medium" style={{ color: '#94a3b8' }}>INSTN / CEA</span>
      </div>

      <div className="flex-1 flex items-center justify-center px-4 pb-10">
        <div className="w-full max-w-sm">

          {/* ── STEP 1: Badge ── */}
          {step === 1 && (
            <>
              <div className="text-center mb-8">
                <div className="w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-5" style={{ background: '#1e2a4a' }}>
                  <Hash size={22} style={{ color: '#f97316' }} />
                </div>
                <h1 className="text-2xl font-black mb-1.5" style={{ color: '#1e293b' }}>Vérification du badge</h1>
                <p className="text-sm leading-relaxed" style={{ color: '#64748b' }}>
                  Saisissez votre numéro de badge à 6 chiffres pour accéder à l'inscription.
                </p>
              </div>

              <form onSubmit={verifyBadge} className="space-y-4">
                {badgeError && (
                  <div className="px-4 py-3 rounded-xl text-sm font-medium" style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#dc2626' }}>
                    {badgeError}
                  </div>
                )}
                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>
                    Numéro de badge
                  </label>
                  <div className="relative">
                    <Hash size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                    <input
                      type="text"
                      inputMode="numeric"
                      maxLength={6}
                      value={badge}
                      onChange={e => setBadge(e.target.value.replace(/\D/g, '').slice(0, 6))}
                      placeholder="000000"
                      className={inputCls}
                      style={{ ...iStyle, paddingLeft: '2.4rem', letterSpacing: '0.25em', fontWeight: 700, fontSize: '1.1rem', textAlign: 'center' }}
                      onFocus={onFocus}
                      onBlur={onBlur}
                    />
                  </div>
                  <p className="text-xs mt-1.5 text-center" style={{ color: '#94a3b8' }}>
                    Ce code vous a été fourni par l'administrateur
                  </p>
                </div>

                <button
                  type="submit"
                  disabled={badgeLoading || badge.length !== 6}
                  className="w-full flex items-center justify-center gap-2 py-3.5 rounded-xl text-sm font-bold transition-all"
                  style={{
                    background: badge.length === 6 && !badgeLoading ? '#1e2a4a' : '#e2e8f0',
                    color: badge.length === 6 && !badgeLoading ? '#fff' : '#94a3b8',
                    boxShadow: badge.length === 6 && !badgeLoading ? '0 4px 12px rgba(30,42,74,0.3)' : 'none',
                    cursor: badge.length !== 6 ? 'not-allowed' : 'pointer',
                  }}
                  onMouseEnter={e => { if (badge.length === 6 && !badgeLoading) (e.currentTarget.style.background = '#162036'); }}
                  onMouseLeave={e => { if (badge.length === 6 && !badgeLoading) (e.currentTarget.style.background = '#1e2a4a'); }}
                >
                  {badgeLoading ? 'Vérification…' : <><span>Vérifier le badge</span><ArrowRight size={15} /></>}
                </button>
              </form>

              <p className="text-center text-sm mt-6" style={{ color: '#94a3b8' }}>
                Déjà un compte ?{' '}
                <button onClick={onBack} className="font-semibold transition-colors" style={{ color: '#f97316' }}
                  onMouseEnter={e => (e.currentTarget.style.color = '#ea580c')}
                  onMouseLeave={e => (e.currentTarget.style.color = '#f97316')}>
                  Se connecter
                </button>
              </p>
            </>
          )}

          {/* ── STEP 2: Account info ── */}
          {step === 2 && (
            <>
              <div className="text-center mb-7">
                <div className="w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-4" style={{ background: '#1e2a4a' }}>
                  <ShieldCheck size={22} style={{ color: '#f97316' }} />
                </div>
                {/* Badge confirmed badge */}
                <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold mb-4"
                  style={{ background: '#f0fdf4', color: '#16a34a', border: '1px solid #bbf7d0' }}>
                  <ShieldCheck size={11} /> Badge {badge} validé
                </div>
                <h1 className="text-2xl font-black mb-1" style={{ color: '#1e293b' }}>Créer votre compte</h1>
                <p className="text-sm" style={{ color: '#64748b' }}>Renseignez vos informations personnelles</p>
              </div>

              <form onSubmit={submit} className="space-y-3.5">
                {formError && (
                  <div className="px-4 py-3 rounded-xl text-sm font-medium" style={{ background: '#fef2f2', border: '1px solid #fecaca', color: '#dc2626' }}>
                    {formError}
                  </div>
                )}

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>Prénom</label>
                    <div className="relative">
                      <User size={13} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                      <input type="text" value={prenom} onChange={e => setPrenom(e.target.value)} placeholder="Martin"
                        className={inputCls} style={{ ...iStyle, paddingLeft: '2.1rem' }} onFocus={onFocus} onBlur={onBlur} />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>Nom</label>
                    <input type="text" value={nom} onChange={e => setNom(e.target.value)} placeholder="Dupont"
                      className={inputCls} style={iStyle} onFocus={onFocus} onBlur={onBlur} />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>Adresse mail CEA</label>
                  <div className="relative">
                    <Mail size={13} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                    <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="votre.nom@cea.fr"
                      className={inputCls} style={{ ...iStyle, paddingLeft: '2.1rem' }} onFocus={onFocus} onBlur={onBlur} />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>
                    Unité <span style={{ color: '#cbd5e1', fontWeight: 400 }}>(optionnel)</span>
                  </label>
                  <div className="relative">
                    <Building size={13} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                    <input type="text" value={unite} onChange={e => setUnite(e.target.value)} placeholder="ex : DEN/SFEN"
                      className={inputCls} style={{ ...iStyle, paddingLeft: '2.1rem' }} onFocus={onFocus} onBlur={onBlur} />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>Mot de passe</label>
                  <div className="relative">
                    <Lock size={13} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                    <input type={showPw ? 'text' : 'password'} value={password} onChange={e => setPassword(e.target.value)} placeholder="8 caractères min."
                      className={inputCls} style={{ ...iStyle, paddingLeft: '2.1rem', paddingRight: '2.5rem' }} onFocus={onFocus} onBlur={onBlur} />
                    <button type="button" onClick={() => setShowPw(v => !v)}
                      className="absolute right-3 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }}>
                      {showPw ? <EyeOff size={14} /> : <Eye size={14} />}
                    </button>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider mb-1.5" style={{ color: '#475569' }}>Confirmer le mot de passe</label>
                  <div className="relative">
                    <Lock size={13} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: '#94a3b8' }} />
                    <input type={showPw ? 'text' : 'password'} value={confirm} onChange={e => setConfirm(e.target.value)} placeholder="••••••••"
                      className={inputCls} style={{ ...iStyle, paddingLeft: '2.1rem' }} onFocus={onFocus} onBlur={onBlur} />
                  </div>
                </div>

                <button
                  type="submit"
                  disabled={formLoading}
                  className="w-full flex items-center justify-center gap-2 py-3.5 rounded-xl text-sm font-bold transition-all mt-1"
                  style={{ background: formLoading ? '#94a3b8' : '#1e2a4a', color: '#fff', boxShadow: formLoading ? 'none' : '0 4px 12px rgba(30,42,74,0.3)' }}
                  onMouseEnter={e => { if (!formLoading) (e.currentTarget.style.background = '#162036'); }}
                  onMouseLeave={e => { if (!formLoading) (e.currentTarget.style.background = '#1e2a4a'); }}
                >
                  {formLoading ? 'Création…' : <><span>Créer mon compte</span><ArrowRight size={15} /></>}
                </button>
              </form>

              <button onClick={() => setStep(1)} className="w-full flex items-center justify-center gap-1.5 text-sm mt-4 transition-colors" style={{ color: '#94a3b8' }}
                onMouseEnter={e => (e.currentTarget.style.color = '#475569')}
                onMouseLeave={e => (e.currentTarget.style.color = '#94a3b8')}>
                <ArrowLeft size={13} /> Changer de badge
              </button>
            </>
          )}
        </div>
      </div>

      <footer className="text-center py-4 text-xs" style={{ color: '#cbd5e1' }}>
        CADI Web · INSTN / CEA
      </footer>
    </div>
  );
}
