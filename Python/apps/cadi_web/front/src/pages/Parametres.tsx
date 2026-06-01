import { useState } from 'react';
import { CheckCircle, Edit2, UserX, Trash2, FolderOpen, TestTube, X, AlertTriangle } from 'lucide-react';
import { Card, Badge, Btn, Field, Input, Select, SectionLabel } from '../components/ui';

type Tab = 'profil' | 'chemins' | 'administration';

type User = { name: string; mail: string; role: string; actif: boolean };

interface ParametresProps {
  currentUser?: { prenom: string; nom: string; email: string; unite?: string } | null;
  onUpdateUser?: (user: { prenom: string; nom: string; email: string; unite?: string }) => void;
}

const initialUsers: User[] = [
  { name: 'Martin Dupont',  mail: 'martin.dupont@cea.fr',  role: 'Administrateur', actif: true },
  { name: 'Sophie Moreau',  mail: 'sophie.moreau@cea.fr',  role: 'Utilisateur',    actif: true },
  { name: 'Paul Girard',    mail: 'paul.girard@cea.fr',    role: 'Lecteur',        actif: false },
];

const tabs: { id: Tab; label: string }[] = [
  { id: 'profil',         label: 'Profil' },
  { id: 'chemins',        label: 'Chemins' },
  { id: 'administration', label: 'Administration' },
];

function EditUserModal({ user, onClose, onSave }: { user: User; onClose: () => void; onSave: (u: User) => void }) {
  const [name, setName] = useState(user.name);
  const [mail, setMail] = useState(user.mail);
  const [role, setRole] = useState(user.role);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: 'rgba(0,0,0,0.4)' }} onClick={onClose}>
      <div className="w-full max-w-md rounded-2xl p-6 relative" style={{ background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        <button onClick={onClose} className="absolute top-4 right-4 w-7 h-7 rounded-lg flex items-center justify-center hover:bg-slate-100 transition-colors">
          <X size={15} style={{ color: '#94a3b8' }} />
        </button>
        <h2 className="text-base font-bold mb-5" style={{ color: '#1e293b' }}>Modifier l'utilisateur</h2>
        <div className="space-y-4">
          <Field label="Nom complet"><Input value={name} onChange={setName} /></Field>
          <Field label="Adresse mail"><Input value={mail} onChange={setMail} /></Field>
          <Field label="Rôle"><Select value={role} onChange={setRole} options={['Administrateur', 'Utilisateur', 'Lecteur']} /></Field>
        </div>
        <div className="flex gap-3 mt-6">
          <Btn variant="secondary" onClick={onClose} fullWidth size="md">Annuler</Btn>
          <Btn onClick={() => onSave({ ...user, name, mail, role })} fullWidth size="md">
            <CheckCircle size={14} /> Enregistrer
          </Btn>
        </div>
      </div>
    </div>
  );
}

function ConfirmModal({ msg, onConfirm, onClose }: { msg: string; onConfirm: () => void; onClose: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: 'rgba(0,0,0,0.4)' }} onClick={onClose}>
      <div className="w-full max-w-sm rounded-2xl p-6 relative" style={{ background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        <div className="flex items-center gap-3 mb-4">
          <div className="w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0" style={{ background: '#fef2f2' }}>
            <AlertTriangle size={18} className="text-red-500" />
          </div>
          <p className="text-sm font-semibold" style={{ color: '#1e293b' }}>{msg}</p>
        </div>
        <div className="flex gap-3">
          <Btn variant="secondary" onClick={onClose} fullWidth size="md">Annuler</Btn>
          <button
            onClick={() => { onConfirm(); onClose(); }}
            className="flex-1 px-4 py-2.5 rounded-xl text-sm font-semibold transition-colors"
            style={{ background: '#dc2626', color: '#fff' }}
            onMouseEnter={e => (e.currentTarget.style.background = '#b91c1c')}
            onMouseLeave={e => (e.currentTarget.style.background = '#dc2626')}
          >
            Confirmer
          </button>
        </div>
      </div>
    </div>
  );
}

export default function Parametres({ currentUser, onUpdateUser }: ParametresProps) {
  const [tab, setTab]     = useState<Tab>('profil');
  const [saved, setSaved] = useState(false);
  const [nom, setNom]     = useState(currentUser?.nom || 'Dupont');
  const [prenom, setPrenom] = useState(currentUser?.prenom || 'Martin');
  const [mail, setMail]   = useState(currentUser?.email || 'martin.dupont@cea.fr');
  const [unite, setUnite] = useState(currentUser?.unite || 'DEN/SFEN');
  const [site, setSite]   = useState('Saclay');

  const [chemins, setChemins] = useState([
    { key: 'ged',   label: 'GED IRIS',          val: '\\\\serveur-ged\\IRIS\\exports' },
    { key: 'orig',  label: 'Extracts originaux', val: 'C:\\CADI\\extracts\\originaux' },
    { key: 'comp',  label: 'Extracts complets',  val: 'C:\\CADI\\extracts\\complets' },
    { key: 'word',  label: 'Modèles Word',        val: 'C:\\CADI\\modeles\\word' },
    { key: 'excel', label: 'Modèles Excel',       val: 'C:\\CADI\\modeles\\excel' },
  ]);
  const [testResult, setTestResult] = useState<null | 'ok' | 'error'>(null);
  const [testing, setTesting] = useState(false);

  const [users, setUsers] = useState<User[]>(initialUsers);
  const [editingUser, setEditingUser] = useState<User | null>(null);
  const [confirmDelete, setConfirmDelete] = useState<number | null>(null);
  const [toast, setToast] = useState('');

  const showToast = (msg: string) => { setToast(msg); setTimeout(() => setToast(''), 2500); };
  const save = () => { 
    setSaved(true); 
    onUpdateUser?.({ nom, prenom, email: mail, unite });
    setTimeout(() => setSaved(false), 2500); 
  };

  const testChemins = () => {
    setTesting(true); setTestResult(null);
    setTimeout(() => { setTesting(false); setTestResult('ok'); }, 1600);
  };

  const updateChemin = (key: string, val: string) =>
    setChemins(p => p.map(c => c.key === key ? { ...c, val } : c));

  const browseChemin = (key: string) => {
    const example = key === 'ged' ? '\\\\serveur-ged\\IRIS\\exports_new' : `C:\\CADI\\${key}\\nouveau`;
    updateChemin(key, example);
    showToast('Chemin mis à jour (simulation)');
  };

  const toggleUser = (i: number) =>
    setUsers(p => p.map((u, j) => j === i ? { ...u, actif: !u.actif } : u));

  const deleteUser = (i: number) =>
    setUsers(p => p.filter((_, j) => j !== i));

  const saveUser = (u: User) => {
    setUsers(p => p.map(x => x.mail === editingUser?.mail ? u : x));
    setEditingUser(null);
    showToast('Utilisateur modifié');
  };

  return (
    <>
      {toast && (
        <div className="fixed top-4 right-4 z-50 flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium"
          style={{ background: '#1e2a4a', color: '#fff', boxShadow: '0 8px 24px rgba(0,0,0,0.2)' }}>
          <CheckCircle size={15} /> {toast}
        </div>
      )}
      {editingUser && (
        <EditUserModal user={editingUser} onClose={() => setEditingUser(null)} onSave={saveUser} />
      )}
      {confirmDelete !== null && (
        <ConfirmModal
          msg={`Supprimer ${users[confirmDelete]?.name} ?`}
          onConfirm={() => deleteUser(confirmDelete)}
          onClose={() => setConfirmDelete(null)}
        />
      )}

      <div className="space-y-5">
        <Card noPad>
          {/* Tab bar */}
          <div className="flex" style={{ borderBottom: '1px solid #f1f5f9' }}>
            {tabs.map(t => (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className="relative px-6 py-3.5 text-sm font-semibold transition-colors"
                style={{ color: tab === t.id ? '#f97316' : '#64748b' }}
              >
                {t.label}
                {tab === t.id && (
                  <span className="absolute bottom-0 left-3 right-3 h-0.5 rounded-t-full" style={{ background: '#f97316' }} />
                )}
              </button>
            ))}
          </div>

          <div className="p-6">

            {/* ── PROFIL ── */}
            {tab === 'profil' && (
              <div className="max-w-lg space-y-5">
                <SectionLabel letter="P" title="Informations personnelles" />
                <div className="grid grid-cols-2 gap-4">
                  <Field label="Nom"><Input value={nom} onChange={setNom} /></Field>
                  <Field label="Prénom"><Input value={prenom} onChange={setPrenom} /></Field>
                </div>
                <Field label="Adresse mail"><Input value={mail} onChange={setMail} /></Field>
                <Field label="Unité"><Input value={unite} onChange={setUnite} /></Field>
                <Field label="Site">
                  <Select value={site} onChange={setSite} options={['Saclay', 'Cadarache', 'Grenoble', 'Marcoule', 'Valduc']} />
                </Field>
                <Field label="Rôle"><Input value="Administrateur" readOnly /></Field>
                <Btn onClick={save} size="md">
                  {saved ? <><CheckCircle size={14} /> Enregistré</> : 'Enregistrer les modifications'}
                </Btn>
              </div>
            )}

            {/* ── CHEMINS ── */}
            {tab === 'chemins' && (
              <div className="max-w-2xl space-y-5">
                <SectionLabel letter="C" title="Configuration des chemins réseau" />
                {chemins.map(({ key, label, val }) => (
                  <Field key={key} label={label}>
                    <div className="flex gap-2">
                      <div className="flex-1">
                        <Input value={val} mono onChange={v => updateChemin(key, v)} />
                      </div>
                      <button
                        onClick={() => browseChemin(key)}
                        className="w-9 h-9 rounded-lg flex items-center justify-center flex-shrink-0 transition-colors"
                        style={{ border: '1px solid #e2e8f0', background: '#f8fafc' }}
                        onMouseEnter={e => (e.currentTarget.style.background = '#f1f5f9')}
                        onMouseLeave={e => (e.currentTarget.style.background = '#f8fafc')}
                        title="Parcourir"
                      >
                        <FolderOpen size={14} style={{ color: '#64748b' }} />
                      </button>
                    </div>
                  </Field>
                ))}
                <div className="flex items-center gap-3 pt-2 flex-wrap">
                  <Btn variant="secondary" size="md" onClick={testChemins} disabled={testing}>
                    <TestTube size={14} /> {testing ? 'Test en cours…' : 'Tester les chemins'}
                  </Btn>
                  <Btn onClick={save} size="md">
                    {saved ? <><CheckCircle size={14} /> Enregistré</> : 'Enregistrer'}
                  </Btn>
                  {testResult === 'ok' && (
                    <span className="flex items-center gap-1.5 text-sm font-medium" style={{ color: '#16a34a' }}>
                      <CheckCircle size={14} /> Tous les chemins sont accessibles
                    </span>
                  )}
                </div>
              </div>
            )}

            {/* ── ADMINISTRATION ── */}
            {tab === 'administration' && (
              <div>
                <SectionLabel letter="A" title="Gestion des utilisateurs" />
                <div className="overflow-hidden rounded-xl" style={{ border: '1px solid #f1f5f9' }}>
                  <table className="w-full">
                    <thead>
                      <tr style={{ background: '#f8fafc', borderBottom: '1px solid #f1f5f9' }}>
                        {['Utilisateur', 'Adresse mail', 'Rôle', 'Statut', 'Actions'].map(h => (
                          <th key={h} className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest" style={{ color: '#94a3b8' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {users.map((u, i) => (
                        <tr key={i} style={{ borderBottom: i < users.length - 1 ? '1px solid #f8fafc' : 'none' }}
                          onMouseEnter={e => (e.currentTarget.style.background = '#fafbfc')}
                          onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                        >
                          <td className="px-4 py-3.5">
                            <div className="flex items-center gap-3">
                              <div className="w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold text-white"
                                style={{ background: '#1e2a4a' }}>
                                {u.name.split(' ').map(n => n[0]).join('')}
                              </div>
                              <span className="text-sm font-semibold" style={{ color: '#1e293b' }}>{u.name}</span>
                            </div>
                          </td>
                          <td className="px-4 py-3.5 font-mono text-xs" style={{ color: '#64748b' }}>{u.mail}</td>
                          <td className="px-4 py-3.5 text-sm" style={{ color: '#475569' }}>{u.role}</td>
                          <td className="px-4 py-3.5">
                            <Badge variant={u.actif ? 'ok' : 'gray'}>{u.actif ? 'Actif' : 'Désactivé'}</Badge>
                          </td>
                          <td className="px-4 py-3.5">
                            <div className="flex items-center gap-1">
                              <button
                                onClick={() => setEditingUser(u)}
                                className="w-7 h-7 rounded-lg flex items-center justify-center transition-colors hover:bg-slate-100"
                                title="Modifier"
                              >
                                <Edit2 size={13} style={{ color: '#64748b' }} />
                              </button>
                              <button
                                onClick={() => { toggleUser(i); showToast(u.actif ? `${u.name} désactivé` : `${u.name} activé`); }}
                                className="w-7 h-7 rounded-lg flex items-center justify-center transition-colors hover:bg-amber-50"
                                title={u.actif ? 'Désactiver' : 'Activer'}
                              >
                                <UserX size={13} className="text-amber-500" />
                              </button>
                              <button
                                onClick={() => setConfirmDelete(i)}
                                className="w-7 h-7 rounded-lg flex items-center justify-center transition-colors hover:bg-red-50"
                                title="Supprimer"
                              >
                                <Trash2 size={13} className="text-red-400" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        </Card>
      </div>
    </>
  );
}
