/**
 * Écran de la Messagerie.
 * Gère le journal des transmissions sécurisées et l'historique des envois
 * de bilans officiels signés au chef de l'unité DEN/SFEN.
 */
import { useState, useEffect } from 'react';
import { Mail, Calendar, User, FileText, ChevronRight, Search, Eye, AlertCircle, CheckCircle } from 'lucide-react';
import { Card, CardHeader, Badge } from '../components/ui';

interface MailRecord {
  id: string;
  destinataire: string;
  objet: string;
  titre: string;
  message: string;
  date: string;
  fichiers: string[];
}

const SEED_HISTORY: MailRecord[] = [
  {
    id: 'seed-1',
    destinataire: "Chef d'unité — DEN/SFEN (chef-dunite@cea.fr)",
    objet: "[CADI Web] Bilan Sessions TEL — Période : Année (2025)",
    titre: "Transmission du bilan signé",
    message: `Bonjour,

Veuillez trouver ci-joint le bilan officiel signé pour la formation "TEL" couvrant la période Année de l'année 2025.

Ce document intègre la fiche de coûts consolidée et les spécifications pédagogiques conformes.

Cordialement,
Martin Dupont
Responsable de formation — INSTN Saclay`,
    date: "28/05/2026 14:32",
    fichiers: ["Bilan-Sessions-TEL-2025-Annee_signed.docx"]
  },
  {
    id: 'seed-2',
    destinataire: "Chef d'unité — DEN/SFEN (chef-dunite@cea.fr)",
    objet: "[CADI Web] Bilan Formation RC2 — Période : 1er semestre (2026)",
    titre: "Envoi du bilan de session RC2",
    message: `Bonjour,

Comme demandé par la direction de l'INSTN, voici le bilan consolidé pour le premier semestre 2026 de la formation Radioprotection Niveau 2 (RC2).

La fiche de coûts est annexée au document.

Cordialement,
Martin Dupont
Responsable de formation — INSTN Saclay`,
    date: "15/05/2026 09:15",
    fichiers: ["Bilan-Formation-RC2-2026-Semestre1_signed.docx"]
  }
];

interface MessagerieProps {
  currentUser?: { prenom: string; nom: string; email: string; unite?: string } | null;
}

export default function Messagerie({ currentUser }: MessagerieProps) {
  const [history, setHistory] = useState<MailRecord[]>([]);
  const [selectedMail, setSelectedMail] = useState<MailRecord | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  // Load history from localStorage or seed it if empty
  useEffect(() => {
    const existing = localStorage.getItem('cadi_mail_history');
    if (existing) {
      setHistory(JSON.parse(existing));
    } else {
      localStorage.setItem('cadi_mail_history', JSON.stringify(SEED_HISTORY));
      setHistory(SEED_HISTORY);
    }
  }, []);

  // Filter messages based on search query
  const filteredHistory = history.filter(
    m => m.objet.toLowerCase().includes(searchTerm.toLowerCase()) || 
         m.message.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="grid grid-cols-5 gap-6">

      {/* Left Column: Sent Messages List */}
      <div className="col-span-3 space-y-4">
        <Card noPad>
          <div className="flex items-center justify-between px-5 py-3.5 border-b border-slate-100">
            <div>
              <h3 className="text-sm font-extrabold text-slate-800 uppercase tracking-wider">Journal des transmissions</h3>
              <p className="text-xs text-slate-400 mt-0.5">Historique des bilans envoyés au Chef de l'unité</p>
            </div>
            <Badge variant="navy">{filteredHistory.length} envoyés</Badge>
          </div>

          {/* Search bar */}
          <div className="p-4 bg-slate-50/50 border-b border-slate-100 flex items-center gap-2">
            <Search size={14} className="text-slate-400" />
            <input
              type="text"
              placeholder="Rechercher un bilan ou un mot-clé..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
              className="text-xs bg-transparent outline-none w-full text-slate-800"
            />
          </div>

          {/* List */}
          <div className="divide-y divide-slate-100 max-h-[60vh] overflow-y-auto">
            {filteredHistory.length > 0 ? (
              filteredHistory.map((m) => {
                const isSelected = selectedMail?.id === m.id;
                return (
                  <div
                    key={m.id}
                    onClick={() => setSelectedMail(m)}
                    className="p-4 cursor-pointer transition-all hover:bg-slate-50 flex items-start gap-4"
                    style={{ background: isSelected ? '#fff7ed' : 'transparent' }}
                  >
                    <div
                      className="w-8 h-8 rounded-lg flex items-center justify-center flex-shrink-0"
                      style={{
                        background: isSelected ? '#fed7aa' : '#f1f5f9',
                        color: isSelected ? '#f97316' : '#64748b'
                      }}
                    >
                      <Mail size={15} />
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-bold text-slate-400 flex items-center gap-1">
                          <Calendar size={10} />
                          {m.date}
                        </span>
                        <Badge variant="ok">Transmis (OK)</Badge>
                      </div>
                      <h4 className="text-xs font-bold text-slate-800 mt-1 truncate">{m.objet}</h4>
                      <p className="text-[11px] text-slate-400 mt-0.5 truncate leading-relaxed">
                        {m.message}
                      </p>
                    </div>
                    <ChevronRight size={14} className="text-slate-300 mt-4 flex-shrink-0" />
                  </div>
                );
              })
            ) : (
              <div className="p-8 text-center text-slate-400 text-xs flex flex-col items-center justify-center gap-2">
                <AlertCircle size={24} className="text-slate-300" />
                Aucun message ou bilan transmis trouvé dans l'historique.
              </div>
            )}
          </div>
        </Card>
      </div>

      {/* Right Column: Message Details Viewer */}
      <div className="col-span-2">
        {selectedMail ? (
          <Card noPad className="sticky top-20 border border-orange-100 shadow-md">
            <CardHeader
              title="Aperçu de la transmission"
              subtitle="Enregistrement officiel de la messagerie"
              icon={Mail}
              action={
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-bold border border-emerald-200 flex items-center gap-1">
                  <CheckCircle size={10} /> Envoyé
                </span>
              }
            />
            
            <div className="p-5 space-y-4">
              
              <div className="space-y-2 pb-3 border-b border-slate-100 text-xs">
                <div className="flex items-center gap-2">
                  <span className="w-16 font-extrabold text-slate-400 uppercase">Expéditeur :</span>
                  <span className="text-slate-700 font-semibold flex items-center gap-1">
                    <User size={11} className="text-slate-400" />
                    {currentUser ? `${currentUser.prenom} ${currentUser.nom}` : 'Martin Dupont'} (CADI Web)
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="w-16 font-extrabold text-slate-400 uppercase">Destinataire :</span>
                  <span className="text-slate-700 font-semibold">{selectedMail.destinataire}</span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="w-16 font-extrabold text-slate-400 uppercase">Envoyé le :</span>
                  <span className="text-slate-700 font-semibold flex items-center gap-1">
                    <Calendar size={11} className="text-slate-400" />
                    {selectedMail.date}
                  </span>
                </div>
              </div>

              <div>
                <h4 className="text-[10px] font-extrabold uppercase text-slate-400 tracking-wider mb-1.5">Objet du message</h4>
                <p className="text-xs font-bold text-slate-800 bg-slate-50 p-2.5 rounded-lg border border-slate-100">{selectedMail.objet}</p>
              </div>

              <div>
                <h4 className="text-[10px] font-extrabold uppercase text-slate-400 tracking-wider mb-1.5">Corps du message</h4>
                <pre className="text-xs font-sans text-slate-600 bg-slate-50 p-3 rounded-lg border border-slate-100 whitespace-pre-wrap leading-relaxed">
                  {selectedMail.message.replace(/Martin Dupont/g, currentUser ? `${currentUser.prenom} ${currentUser.nom}` : 'Martin Dupont')}
                </pre>
              </div>

              <div>
                <h4 className="text-[10px] font-extrabold uppercase text-slate-400 tracking-wider mb-1.5">Pièces jointes signées (.docx)</h4>
                <div className="space-y-1.5">
                  {selectedMail.fichiers.map((f, i) => (
                    <div key={i} className="flex items-center justify-between p-2.5 rounded-xl border border-emerald-200 bg-emerald-50/50 text-xs">
                      <div className="flex items-center gap-2 min-w-0">
                        <FileText size={13} className="text-emerald-600 flex-shrink-0" />
                        <span className="font-mono font-bold truncate text-emerald-800">{f}</span>
                      </div>
                      <button
                        onClick={() => alert(`Téléchargement de la pièce jointe archivée : ${f}`)}
                        className="p-1 hover:bg-emerald-100 text-emerald-700 rounded transition-colors"
                        title="Télécharger l'archive signée"
                      >
                        <Eye size={13} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* Status logs */}
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-100 space-y-1.5 text-[10px] text-slate-400">
                <div className="flex items-center gap-1.5 text-emerald-600 font-semibold">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                  Généré et Signé Numériquement dans CADI Web
                </div>
                <div className="flex items-center gap-1.5 text-emerald-600 font-semibold">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                  Transmis via messagerie sécurisée CEA
                </div>
                <div className="flex items-center gap-1.5 text-slate-500 font-semibold">
                  <span className="w-1.5 h-1.5 rounded-full bg-slate-400" />
                  Archivé dans la GED d'unité (GED_IRIS)
                </div>
              </div>

            </div>
          </Card>
        ) : (
          <div className="h-64 rounded-xl border border-dashed border-slate-200 bg-white flex flex-col items-center justify-center text-center p-6 text-slate-400 text-xs gap-2">
            <Eye size={20} className="text-slate-300" />
            Sélectionnez une transmission dans la liste pour voir les détails et télécharger les pièces jointes associées.
          </div>
        )}
      </div>

    </div>
  );
}
