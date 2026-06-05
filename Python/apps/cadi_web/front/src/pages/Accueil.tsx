/**
 * Écran d'Accueil de CADI Web.
 * 
 * Version connectée au Backend FastAPI :
 * - Charge les véritables indicateurs clés (KPIs), alertes et historiques depuis l'API locale.
 * - Inclut des commentaires d'intégration extrêmement clairs pour personnaliser les points de terminaison.
 * - Gère un fallback élégant si le serveur backend n'est pas démarré.
 */

import { useState, useEffect } from 'react';
import {
  Database, BarChart2, FileText, Table2,
  ArrowRight, AlertTriangle, TrendingUp,
  Clock, CheckCircle, XCircle, AlertCircle,
} from 'lucide-react';
import { Card, Badge } from '../components/ui';
import { API_BASE_URL } from '../lib/api';

const quickCards = [
  {
    id: 'iris',
    title: 'IRIS',
    desc: 'Mettre à jour les exports',
    Icon: Database,
    color: '#3b82f6',
    bg: 'linear-gradient(135deg,#eff6ff,#dbeafe)',
    border: '#bfdbfe',
    stat: '4 exports',
  },
  {
    id: 'evalstat',
    title: 'EvalStat',
    desc: 'Traiter les évaluations',
    Icon: BarChart2,
    color: '#16a34a',
    bg: 'linear-gradient(135deg,#f0fdf4,#dcfce7)',
    border: '#bbf7d0',
    stat: '15 réponses',
  },
  {
    id: 'bilans',
    title: 'Bilans',
    desc: 'Générer les bilans',
    Icon: FileText,
    color: '#7c3aed',
    bg: 'linear-gradient(135deg,#f5f3ff,#ede9fe)',
    border: '#ddd6fe',
    stat: '2 en attente',
  },
  {
    id: 'formations',
    title: 'Catalogue Formations',
    desc: 'Consulter les formations',
    Icon: Table2,
    color: '#f97316',
    bg: 'linear-gradient(135deg,#fff7ed,#ffedd5)',
    border: '#fed7aa',
    stat: 'Synchronisé',
  },
];

function StatutIcon({ statut }: { statut: string }) {
  if (statut === 'ok' || statut === 'SUCCESS') return <Badge variant="ok">OK</Badge>;
  if (statut === 'warning' || statut === 'WARNING') return <Badge variant="warning">Avertissement</Badge>;
  return <Badge variant="error">Erreur</Badge>;
}

interface AccueilProps {
  onNavigate: (screen: string) => void;
}

export default function Accueil({ onNavigate }: AccueilProps) {
  // --- ÉTATS CONNECTÉS AU BACKEND ---
  const [kpis, setKpis] = useState<any[]>([]);
  const [recentRows, setRecentRows] = useState<any[]>([]);
  const [alerts, setAlerts] = useState<any[]>([]);
  const [connected, setConnected] = useState<boolean | null>(null);

  // ======================================================================================
  // 🔌 CONNECTEUR BACKEND FASTAPI (CONNEXION RÉELLE)
  // ======================================================================================
  useEffect(() => {
    // Appel de l'endpoint du tableau de bord de votre backend Python FastAPI
    fetch(`${API_BASE_URL}/dashboard/`)
      .then(res => {
        if (!res.ok) throw new Error("Erreur de réponse serveur");
        return res.json();
      })
      .then(data => {
        // Enregistrement des vraies données renvoyées par votre base PostgreSQL locale !
        setKpis([
          { label: 'Traitements cette semaine', value: data.kpis.weekly_treatments || '0', trend: '+3', Icon: TrendingUp },
          { label: 'Sessions actives', value: data.kpis.active_sessions || '0', trend: '+1', Icon: CheckCircle },
          { label: 'Erreurs non résolues', value: data.kpis.unresolved_errors || '0', trend: '0', Icon: XCircle },
        ]);
        setRecentRows(data.recent_activity || []);
        setAlerts(data.alerts || []);
        setConnected(true);
      })
      .catch(err => {
        console.warn("Connexion au backend FastAPI indisponible, chargement du mode local autonome.", err);
        setConnected(false);
        // Fallback local structuré (permet de travailler même si le serveur Python est éteint)
        setKpis([
          { label: 'Traitements cette semaine', value: '18', trend: '+3', Icon: TrendingUp },
          { label: 'Sessions actives', value: '7', trend: '+1', Icon: CheckCircle },
          { label: 'Erreurs non résolues', value: '2', trend: '-1', Icon: XCircle },
        ]);
        setRecentRows([
          { date: '26/05/2026', module: 'IRIS', formation: 'TEL — 12995', statut: 'ok', fichier: 'export_R04110_20260526.csv' },
          { date: '25/05/2026', module: 'EvalStat', formation: 'RC2 — 12880', statut: 'warning', fichier: 'Evaluation-RC2-12880.xlsx' },
          { date: '24/05/2026', module: 'Bilans', formation: 'TAN — 12741', statut: 'error', fichier: 'Bilan-TAN-2025-S2.docx' },
        ]);
        setAlerts([
          { msg: 'Export IRIS R04301 (Ventes) non mis à jour depuis 48h', time: 'Il y a 2h', Icon: AlertTriangle },
          { msg: 'Fiche de coûts manquante pour la formation TAN', time: 'Il y a 5h', Icon: AlertCircle },
          { msg: '3 fichiers CSV EvalStat en attente de traitement', time: 'Hier', Icon: Clock },
        ]);
      });
  }, []);

  return (
    <div className="space-y-6">

      {/* Indicateur de connectivité au serveur de base de données réelle */}
      <div className="flex items-center justify-between p-3.5 rounded-xl border bg-white"
        style={{ borderLeft: connected ? '4px solid #10b981' : '4px solid #f59e0b' }}>
        <div className="flex items-center gap-2.5">
          <span className={`w-2.5 h-2.5 rounded-full ${connected ? 'bg-emerald-500 animate-pulse' : 'bg-amber-500 animate-pulse'}`} />
          <span className="text-xs font-bold text-slate-700">
            {connected 
              ? 'Connecté au serveur CADI FastAPI — Base PostgreSQL locale active' 
              : 'Mode autonome local — En attente de connexion avec le serveur CADI FastAPI'}
          </span>
        </div>
        <span className="text-[10px] font-mono text-slate-400">API: {API_BASE_URL}</span>
      </div>

      {/* KPI strip */}
      <div className="grid grid-cols-3 gap-4">
        {kpis.map(({ label, value, trend, Icon }) => (
          <div
            key={label}
            className="bg-white rounded-xl px-5 py-4 flex items-center gap-4"
            style={{ border: '1px solid #e8edf2', boxShadow: '0 1px 3px rgba(30,42,74,0.05)' }}
          >
            <div className="w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0"
              style={{ background: '#f8fafc', border: '1px solid #e8edf2' }}>
              <Icon size={18} style={{ color: '#475569' }} />
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-2xl font-black leading-none" style={{ color: '#1e293b' }}>{value}</p>
              <p className="text-xs mt-0.5 truncate" style={{ color: '#94a3b8' }}>{label}</p>
            </div>
            <span
              className="text-xs font-bold px-2 py-0.5 rounded-full"
              style={{ background: trend.startsWith('-') ? '#fef2f2' : '#f0fdf4', color: trend.startsWith('-') ? '#dc2626' : '#16a34a' }}
            >
              {trend}
            </span>
          </div>
        ))}
      </div>

      {/* Quick-access cards */}
      <div className="grid grid-cols-4 gap-4">
        {quickCards.map(({ id, title, desc, Icon, color, bg, border, stat }) => (
          <div
            key={id}
            className="relative overflow-hidden rounded-xl p-5 flex flex-col gap-4 cursor-pointer group transition-transform hover:-translate-y-0.5"
            style={{ background: bg, border: `1px solid ${border}`, boxShadow: '0 1px 3px rgba(30,42,74,0.06)' }}
            onClick={() => onNavigate(id)}
          >
            <div className="flex items-start justify-between">
              <div
                className="w-10 h-10 rounded-xl flex items-center justify-center"
                style={{ background: 'rgba(255,255,255,0.7)', backdropFilter: 'blur(8px)' }}
              >
                <Icon size={20} style={{ color }} strokeWidth={1.8} />
              </div>
              <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: 'rgba(255,255,255,0.6)', color }}>
                {stat}
              </span>
            </div>
            <div>
              <div className="font-bold text-sm" style={{ color: '#1e293b' }}>{title}</div>
              <div className="text-xs mt-0.5" style={{ color: '#64748b' }}>{desc}</div>
            </div>
            <div className="flex items-center gap-1 text-xs font-semibold mt-auto transition-all group-hover:gap-2" style={{ color }}>
              Accéder <ArrowRight size={13} />
            </div>
          </div>
        ))}
      </div>

      {/* Table + Alerts */}
      <div className="grid grid-cols-3 gap-4">

        {/* Recent treatments */}
        <Card noPad className="col-span-2">
          <div
            className="flex items-center justify-between px-5 py-3.5"
            style={{ borderBottom: '1px solid #f1f5f9' }}
          >
            <div>
              <h3 className="text-sm font-semibold" style={{ color: '#1e293b' }}>Derniers traitements</h3>
              <p className="text-xs mt-0.5" style={{ color: '#94a3b8' }}>Activité réelle extraite du système CADI</p>
            </div>
            <Badge variant="gray">{recentRows.length} entrées</Badge>
          </div>
          <table className="w-full">
            <thead>
              <tr style={{ background: '#fafbfc', borderBottom: '1px solid #f1f5f9' }}>
                {['Date', 'Module', 'Formation', 'Statut', 'Fichier'].map(h => (
                  <th key={h} className="text-left px-5 py-2.5 text-xs font-bold uppercase tracking-widest" style={{ color: '#94a3b8' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {recentRows.map((row, i) => (
                <tr
                  key={i}
                  className="transition-colors"
                  style={{ borderBottom: i < recentRows.length - 1 ? '1px solid #f8fafc' : 'none' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#fafbfc')}
                  onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                >
                  <td className="px-5 py-3 text-xs font-mono" style={{ color: '#64748b' }}>{row.date}</td>
                  <td className="px-5 py-3 text-sm font-semibold" style={{ color: '#1e293b' }}>{row.module}</td>
                  <td className="px-5 py-3 text-sm" style={{ color: '#475569' }}>{row.formation}</td>
                  <td className="px-5 py-3"><StatutIcon statut={row.statut} /></td>
                  <td className="px-5 py-3 text-xs font-mono max-w-0 truncate" style={{ color: '#94a3b8', maxWidth: 180 }}>{row.fichier}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>

        {/* Alerts */}
        <Card noPad>
          <div
            className="flex items-center justify-between px-5 py-3.5"
            style={{ borderBottom: '1px solid #f1f5f9' }}
          >
            <h3 className="text-sm font-semibold" style={{ color: '#1e293b' }}>Alertes système</h3>
            <Badge variant="warning">{alerts.length}</Badge>
          </div>
          <div className="p-4 space-y-2.5">
            {alerts.map(({ msg, time }, i) => (
              <div
                key={i}
                className="flex gap-3 rounded-xl p-3.5"
                style={{ background: '#fffbeb', border: '1px solid #fde68a' }}
              >
                <AlertTriangle size={14} className="text-amber-500 flex-shrink-0 mt-0.5" />
                <div className="min-w-0">
                  <p className="text-xs leading-snug font-medium" style={{ color: '#92400e' }}>{msg}</p>
                  <p className="text-xs mt-1.5 font-semibold" style={{ color: '#b45309' }}>{time}</p>
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}
