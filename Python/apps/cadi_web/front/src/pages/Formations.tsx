/**
 * Écran Catalogue des Formations — CADI Web.
 * 
 * Version connectée au Backend & Base de Données PostgreSQL :
 * - Effectue un GET fetch vers '/api/formations/' au démarrage.
 * - Le modal de création envoie un POST réel pour persister la formation en base.
 * - Dispose d'une tolérance aux pannes locale en cas de serveur non actif.
 */

import { useState, useEffect } from 'react';
import { BookOpen, Search, Plus, ExternalLink, X, CheckCircle, Database } from 'lucide-react';
import { Card, Badge, Btn, Field, Input, Select } from '../components/ui';

const SEED_FORMATIONS = [
  { code: 'TEL', nom: "Techniques d'Exploitation des Réacteurs",       sessions: 4, statut: 'Actif',    updated: '26/05/2026' },
  { code: 'RC2', nom: 'Radioprotection de Niveau 2',                  sessions: 6, statut: 'Actif',    updated: '24/05/2026' },
  { code: 'TAN', nom: 'Thermique Appliquée Nucléaire',                 sessions: 2, statut: 'Suspendu', updated: '10/03/2026' },
  { code: 'CHI', nom: 'Chimie des Circuits Primaires',                 sessions: 3, statut: 'Actif',    updated: '22/05/2026' },
  { code: 'SEC', nom: 'Sécurité Incendie — Installations Nucléaires',  sessions: 5, statut: 'Actif',    updated: '20/05/2026' },
  { code: 'DEM', nom: 'Démantèlement et Assainissement',             sessions: 1, statut: 'En cours', updated: '15/04/2026' },
];

const codeColors: Record<string, string> = {
  TEL: '#3b82f6', RC2: '#16a34a', TAN: '#94a3b8',
  CHI: '#f97316', SEC: '#dc2626', DEM: '#0891b2',
};

type Formation = typeof SEED_FORMATIONS[0];

function DetailModal({ f, onClose }: { f: Formation; onClose: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm" onClick={onClose}>
      <div
        className="w-full max-w-md rounded-2xl p-6 relative bg-white shadow-2xl animate-scaleUp border border-slate-100"
        onClick={e => e.stopPropagation()}
      >
        <button onClick={onClose} className="absolute top-4 right-4 w-7 h-7 rounded-lg flex items-center justify-center hover:bg-slate-100 transition-colors">
          <X size={15} style={{ color: '#94a3b8' }} />
        </button>
        <div className="flex items-center gap-3 mb-5">
          <div className="w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0"
            style={{ background: `${codeColors[f.code] ?? '#64748b'}15` }}>
            <BookOpen size={18} style={{ color: codeColors[f.code] ?? '#64748b' }} strokeWidth={1.8} />
          </div>
          <div>
            <p className="font-bold text-base text-slate-800 leading-tight">{f.nom}</p>
            <p className="text-xs font-mono font-bold mt-0.5" style={{ color: codeColors[f.code] ?? '#64748b' }}>{f.code}</p>
          </div>
        </div>
        <div className="space-y-3 text-sm">
          {[
            ['Trigramme', f.code],
            ['Sessions', `${f.sessions} session${f.sessions > 1 ? 's' : ''}`],
            ['Statut', f.statut],
            ['Dernière MAJ', f.updated],
          ].map(([k, v]) => (
            <div key={k} className="flex justify-between items-center py-2" style={{ borderBottom: '1px solid #f1f5f9' }}>
              <span className="font-medium text-slate-400">{k}</span>
              <span className="font-semibold text-slate-800">{v}</span>
            </div>
          ))}
        </div>
        <div className="mt-5">
          <Btn onClick={onClose} fullWidth size="md">Fermer</Btn>
        </div>
      </div>
    </div>
  );
}

function NewFormationModal({ onClose, onSave }: { onClose: () => void; onSave: (f: Formation) => void }) {
  const [code, setCode] = useState('');
  const [nom, setNom] = useState('');
  const [statut, setStatut] = useState('Actif');

  const save = () => {
    if (!code.trim() || !nom.trim()) return;
    onSave({
      code: code.toUpperCase().trim(),
      nom: nom.trim(),
      sessions: 0,
      statut,
      updated: new Date().toLocaleDateString('fr-FR'),
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm" onClick={onClose}>
      <div
        className="w-full max-w-md rounded-2xl p-6 relative bg-white shadow-2xl animate-scaleUp border border-slate-100"
        onClick={e => e.stopPropagation()}
      >
        <button onClick={onClose} className="absolute top-4 right-4 w-7 h-7 rounded-lg flex items-center justify-center hover:bg-slate-100 transition-colors">
          <X size={15} style={{ color: '#94a3b8' }} />
        </button>
        <h2 className="text-sm font-extrabold uppercase tracking-wider text-slate-800 mb-5">Nouvelle formation</h2>
        <div className="space-y-4">
          <Field label="Code / Trigramme">
            <Input value={code} onChange={setCode} placeholder="ex : PHY" />
          </Field>
          <Field label="Intitulé">
            <Input value={nom} onChange={setNom} placeholder="ex : Physique des Réacteurs" />
          </Field>
          <Field label="Statut">
            <Select value={statut} onChange={setStatut} options={['Actif', 'Suspendu', 'En cours']} />
          </Field>
        </div>
        <div className="flex gap-3 mt-6">
          <Btn variant="secondary" onClick={onClose} fullWidth size="md">Annuler</Btn>
          <Btn onClick={save} fullWidth size="md" disabled={!code.trim() || !nom.trim()}>
            <CheckCircle size={14} /> Créer la formation
          </Btn>
        </div>
      </div>
    </div>
  );
}

export default function Formations() {
  const API_BASE_URL = 'http://127.0.0.1:8000/api';

  const [q, setQ] = useState('');
  const [formations, setFormations] = useState<Formation[]>([]);
  const [showNew, setShowNew] = useState(false);
  const [detail, setDetail] = useState<Formation | null>(null);
  const [connected, setConnected] = useState<boolean | null>(null);

  // ======================================================================================
  // 🔌 CHARGEMENT DE LA BASE DE DONNÉES POSTGRES VIA FASTAPI (GET)
  // ======================================================================================
  const loadFormations = () => {
    fetch(`${API_BASE_URL}/formations/`)
      .then(res => {
        if (!res.ok) throw new Error("Erreur de réponse");
        return res.json();
      })
      .then(data => {
        setFormations(data || []);
        setConnected(true);
      })
      .catch(err => {
        console.warn("Échec de connexion au catalogue FastAPI. Chargement local.", err);
        setConnected(false);
        // Fallback local via localStorage
        const existing = localStorage.getItem('cadi_formations');
        if (existing) {
          setFormations(JSON.parse(existing));
        } else {
          localStorage.setItem('cadi_formations', JSON.stringify(SEED_FORMATIONS));
          setFormations(SEED_FORMATIONS);
        }
      });
  };

  useEffect(() => {
    loadFormations();
  }, []);

  // ======================================================================================
  // 🔌 SAUVEGARDE D'UNE NOUVELLE FORMATION EN BASE POSTGRES VIA FASTAPI (POST)
  // ======================================================================================
  const addFormation = (f: Formation) => {
    fetch(`${API_BASE_URL}/formations/`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(f)
    })
      .then(res => {
        if (!res.ok) throw new Error("Échec d'enregistrement en base");
        return res.json();
      })
      .then(() => {
        loadFormations(); // Recharger le catalogue à jour
        setShowNew(false);
      })
      .catch(err => {
        console.warn("Enregistrement de secours local effectué.", err);
        const updated = [...formations, f];
        setFormations(updated);
        localStorage.setItem('cadi_formations', JSON.stringify(updated));
        setShowNew(false);
      });
  };

  const filtered = formations.filter(
    f => f.code.toLowerCase().includes(q.toLowerCase()) || f.nom.toLowerCase().includes(q.toLowerCase())
  );

  return (
    <>
      {showNew && <NewFormationModal onClose={() => setShowNew(false)} onSave={addFormation} />}
      {detail && <DetailModal f={detail} onClose={() => setDetail(null)} />}

      <div className="space-y-5">
        {/* DB Sync Indicator */}
        <div className="flex items-center justify-between p-3.5 rounded-xl border bg-white"
          style={{ borderLeft: connected ? '4px solid #3b82f6' : '4px solid #94a3b8' }}>
          <div className="flex items-center gap-2.5 text-xs font-bold text-slate-700">
            <Database size={15} className={connected ? 'text-blue-600' : 'text-slate-400'} />
            <span>
              {connected 
                ? 'Catalogue CADI connecté à la base PostgreSQL réelle — Modifications persistées en direct' 
                : 'Catalogue CADI en cache local (Serveur hors-ligne)'}
            </span>
          </div>
          <span className="text-[10px] font-mono text-slate-400">POSTGRES TABLE: formations</span>
        </div>

        {/* Stats row */}
        <div className="grid grid-cols-4 gap-4">
          {[
            { label: 'Formations totales',   val: formations.length },
            { label: 'Formations actives',   val: formations.filter(f => f.statut === 'Actif').length },
            { label: 'Sessions totales',     val: formations.reduce((s, f) => s + f.sessions, 0) },
            { label: 'Suspendues',           val: formations.filter(f => f.statut === 'Suspendu').length },
          ].map(({ label, val }) => (
            <div key={label} className="bg-white rounded-xl px-5 py-4" style={{ border: '1px solid #e8edf2', boxShadow: '0 1px 3px rgba(30,42,74,0.05)' }}>
              <p className="text-2xl font-black" style={{ color: '#1e293b' }}>{val}</p>
              <p className="text-xs mt-0.5" style={{ color: '#94a3b8' }}>{label}</p>
            </div>
          ))}
        </div>

        <Card noPad>
          <div className="flex items-center justify-between px-5 py-3.5" style={{ borderBottom: '1px solid #f1f5f9' }}>
            <div>
              <h3 className="text-sm font-semibold" style={{ color: '#1e293b' }}>Catalogue des formations</h3>
              <p className="text-xs mt-0.5" style={{ color: '#94a3b8' }}>{filtered.length} formation{filtered.length > 1 ? 's' : ''} affichée{filtered.length > 1 ? 's' : ''}</p>
            </div>
            <div className="flex items-center gap-2">
              <div className="flex items-center gap-2 px-3 py-2 rounded-lg" style={{ border: '1px solid #e2e8f0', background: '#f8fafc' }}>
                <Search size={13} style={{ color: '#94a3b8' }} />
                <input
                  type="text"
                  placeholder="Rechercher…"
                  value={q}
                  onChange={e => setQ(e.target.value)}
                  className="text-xs bg-transparent outline-none w-36 text-slate-800"
                />
              </div>
              <Btn size="sm" onClick={() => setShowNew(true)}>
                <Plus size={13} /> Nouvelle formation
              </Btn>
            </div>
          </div>

          <table className="w-full">
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #f1f5f9' }}>
                {['Code', 'Intitulé', 'Sessions', 'Statut', 'Dernière MAJ', ''].map(h => (
                  <th key={h} className="text-left px-5 py-2.5 text-xs font-bold uppercase tracking-widest text-slate-400">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((f, i) => (
                <tr key={i} style={{ borderBottom: i < filtered.length - 1 ? '1px solid #f8fafc' : 'none' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#fafbfc')}
                  onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                >
                  <td className="px-5 py-3.5">
                    <div className="flex items-center gap-2.5">
                      <div className="w-8 h-8 rounded-lg flex items-center justify-center"
                        style={{ background: `${codeColors[f.code] ?? '#64748b'}15` }}>
                        <BookOpen size={14} style={{ color: codeColors[f.code] ?? '#64748b' }} strokeWidth={1.8} />
                      </div>
                      <span className="font-mono font-bold text-sm" style={{ color: codeColors[f.code] ?? '#64748b' }}>{f.code}</span>
                    </div>
                  </td>
                  <td className="px-5 py-3.5 text-sm font-semibold text-slate-700">{f.nom}</td>
                  <td className="px-5 py-3.5">
                    <span className="text-sm font-semibold text-slate-700">{f.sessions}</span>
                    <span className="text-xs ml-1 text-slate-400">session{f.sessions > 1 ? 's' : ''}</span>
                  </td>
                  <td className="px-5 py-3.5">
                    <Badge variant={f.statut === 'Actif' ? 'ok' : f.statut === 'En cours' ? 'blue' : 'gray'}>
                      {f.statut}
                    </Badge>
                  </td>
                  <td className="px-5 py-3.5 text-xs font-mono text-slate-400">{f.updated}</td>
                  <td className="px-5 py-3.5">
                    <button
                      onClick={() => setDetail(f)}
                      className="w-7 h-7 rounded-lg flex items-center justify-center transition-colors hover:bg-slate-100"
                      title="Voir le détail"
                    >
                      <ExternalLink size={13} style={{ color: '#94a3b8' }} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      </div>
    </>
  );
}
