import { useEffect, useRef, useState } from 'react';
import { CheckCircle, ExternalLink, FileUp, Play, Square, XCircle } from 'lucide-react';
import { Badge, Btn, Card, CardHeader } from '../components/ui';
import { API_BASE_URL } from '../lib/api';

// =============================================================================
// Types
// =============================================================================
//
// Cette page dialogue avec le routeur FastAPI `/api/iris`.
// Les types ci-dessous decrivent les objets JSON renvoyes par le back.
// Les garder ici rend les changements d'API visibles a la compilation TypeScript.

type Mode = 'auto' | 'manual';
type RunStatus = 'idle' | 'success' | 'error';

type DumpRow = {
  // Code technique IRIS: R04110, R0304, R04301, R04500.
  code: string;
  // Libelle fonctionnel affiche dans le tableau.
  type: string;
  // Informations du repertoire source configure cote Python.
  source: string;
  sourcePath: string;
  sourceUrl: string;
  // Informations du fichier cible: soit dernier fichier COMPLET produit,
  // soit dernier dump source si aucun resultat n'existe encore.
  fichier: string;
  targetPath: string;
  targetUrl: string;
  // Date deja formatee par le back pour garder un affichage stable.
  updatedAt: string;
  // Statut global ou statut de progression envoye par le job.
  statut: string;
  // Selection utilisateur pour inclure/exclure l'export du traitement.
  selected: boolean;
  // Etat local de verification pre-traitement.
  verification?: 'idle' | 'ok' | 'error';
  verificationMessage?: string;
};

type ResultRow = {
  export: string;
  fichier: string;
  path?: string;
  url?: string;
  date: string;
  statut: string;
};

type ManualSelectedFile = {
  name: string;
  path: string;
  url?: string;
};

type DumpsResponse = {
  dumps: DumpRow[];
  outputFiles?: ResultRow[];
};

type VerificationResult = {
  code: string;
  valid: boolean;
  status: 'ok' | 'error';
  message: string;
  file?: string;
  updatedAt?: string;
};

type VerificationResponse = {
  valid: boolean;
  message: string;
  checks: VerificationResult[];
};


type JobStatus = {
  // Etat du job cote serveur. Tant que `running`, le front continue le polling.
  status: 'running' | 'done' | 'error' | 'cancelled';
  // Progression reelle calculee par le back, entre 0 et 100.
  progress?: number;
  // Etape courante affichee sous la barre de progression.
  step?: string;
  // Historique court des etapes recues pendant le traitement.
  steps_history?: string[];
  // Fichiers produits par le traitement termine.
  files?: ResultRow[];
  // Message final ou message d'erreur fonctionnelle.
  message?: string;
  // Flag utilise pour afficher un resultat comme erreur.
  warning?: boolean;
  // Code IRIS actuellement traite, utile pour colorer la ligne active.
  current_code?: string;
  // Statut par code IRIS, envoye par le backend pendant le job.
  statut_codes?: Record<string, string>;
};

/**
 * Interroge le serveur toutes les 600 ms jusqu'a la fin du job.
 *
 * Le traitement IRIS peut durer longtemps. Le POST de lancement cree donc un
 * job cote FastAPI et renvoie seulement son identifiant. Cette fonction garde
 * l'interface synchronisee avec ce job: progression, etape courante, statuts
 * par ligne et fichiers produits.
 *
 * `signal` permet d'interrompre le polling quand l'utilisateur clique sur
 * Arreter ou quand le composant doit abandonner les requetes en cours.
 */
async function pollJobStatus(
  jobId: string,
  signal: AbortSignal,
  onUpdate: (progress: number, step: string) => void,
  onStatuts?: (statuts: Record<string, string>) => void,
  onStepsHistory?: (steps: string[]) => void,
): Promise<JobStatus> {
  while (true) {
    // Delai volontaire: assez court pour une UX reactive, assez long pour ne pas
    // appeler le backend local en boucle trop agressive.
    await new Promise(resolve => setTimeout(resolve, 600));
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');

    // Le backend renvoie l'etat complet du job a chaque appel.
    const resp = await fetch(`${API_BASE_URL}/iris/job/${jobId}`, { signal });
    if (!resp.ok) throw new Error(`Erreur serveur : ${resp.status}`);

    const job = await resp.json() as JobStatus;
    // Les callbacks permettent au composant parent de mettre a jour seulement ce
    // qui l'interesse sans melanger la logique de polling et le rendu React.
    onUpdate(job.progress ?? 0, job.step ?? '');
    if (job.statut_codes && onStatuts) onStatuts(job.statut_codes);
    if (job.steps_history && onStepsHistory) onStepsHistory(job.steps_history);

    if (job.status === 'done' || job.status === 'cancelled') return job;
    if (job.status === 'error') throw new Error(job.message ?? 'Erreur de traitement');
  }
}

// =============================================================================
// Composants utilitaires
// =============================================================================


/**
 * Colonne statut: rond colore + texte.
 *
 * Le statut peut venir de deux endroits:
 * - `row.statut`, envoye par le backend ou par le job en cours;
 * - `row.verification`, calcule apres l'appel `/validate-structure`.
 *
 * Une erreur de verification est prioritaire: elle indique a l'utilisateur quel
 * fichier bloque avant de lancer le traitement long.
 */
function StatusCell({ row }: { row: DumpRow }) {
  const dot = (color: string) => (
    <span className={`inline-block w-2 h-2 rounded-full flex-shrink-0 ${color}`} />
  );

  if (row.statut === 'blue') {
    return (
      <div className="flex items-center gap-1.5">
        {dot('bg-orange-400 animate-pulse')}
        <span className="text-xs font-semibold text-orange-700">En cours</span>
      </div>
    );
  }
  if (row.statut === 'gray') {
    return (
      <div className="flex items-center gap-1.5">
        {dot('bg-slate-300')}
        <span className="text-xs font-semibold text-slate-400">En attente</span>
      </div>
    );
  }
  if (row.statut === 'error' || row.verification === 'error') {
    return (
      <div className="flex items-center gap-1.5">
        {dot('bg-red-500')}
        <span className="text-xs font-semibold text-red-700">{row.verificationMessage ?? 'Erreur'}</span>
      </div>
    );
  }
  if (row.statut === 'ok' || row.verification === 'ok') {
    return (
      <div className="flex items-center gap-1.5">
        {dot('bg-emerald-500')}
        <span className="text-xs font-semibold text-emerald-700">À jour</span>
      </div>
    );
  }
  return (
    <div className="flex items-center gap-1.5">
      {dot('bg-slate-300')}
      <span className="text-xs font-semibold text-slate-400">En attente</span>
    </div>
  );
}

/**
 * Bouton d'ouverture de fichier ou dossier.
 *
 * Le navigateur ne peut pas ouvrir librement un chemin Windows local. On envoie
 * donc le chemin au backend local, qui verifie qu'il appartient bien aux zones
 * IRIS autorisees, puis declenche l'ouverture cote poste.
 */
function PathLink({ href, openTarget, label }: { href?: string; openTarget?: string; label: string }) {
  const pathToOpen = openTarget || href;
  if (!pathToOpen) return <span className="font-mono text-xs text-slate-500">{label}</span>;

  const handleOpenPath = async () => {
    const response = await fetch(`${API_BASE_URL}/iris/open-path`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path: pathToOpen }),
    });
    if (!response.ok) {
      const message = await response.json().catch(() => null);
      alert(`Ouverture impossible : ${message?.detail ?? response.statusText}`);
    }
  };

  return (
    <button
      type="button"
      onClick={handleOpenPath}
      title={openTarget || href}
      className="inline-flex max-w-[220px] items-center gap-1 font-mono text-xs font-bold text-orange-700 hover:text-orange-900 hover:underline"
    >
      <span className="truncate">{label}</span>
      <ExternalLink size={12} className="flex-shrink-0" />
    </button>
  );
}

// =============================================================================
// Page principale
// =============================================================================

export default function IRIS() {

  // --- Donnees chargees depuis le backend ---
  // `dumps` alimente le tableau principal.
  // `resultRows` contient les fichiers produits par le dernier traitement ou
  // les derniers fichiers COMPLET trouves au chargement.
  const [dumps, setDumps] = useState<DumpRow[]>([]);
  const [resultRows, setResultRows] = useState<ResultRow[]>([]);
  const [loadingDumps, setLoadingDumps] = useState(false);
  const [loadError, setLoadError] = useState('');

  // --- Etat UI pur ---
  // `checks` duplique volontairement `dump.selected`: le tableau l'utilise pour
  // le rendu, et `checks` permet de recuperer vite les codes selectionnes.
  // `manualFiles` stocke les chemins choisis en mode manuel par la fenetre
  // Python/Tkinter ouverte par le backend local.
  const [checks, setChecks] = useState<Record<string, boolean>>({});
  const [mode, setMode] = useState<Mode>('auto');
  const [manualFiles, setManualFiles] = useState<Record<string, ManualSelectedFile[]>>({});
  const [manualToast, setManualToast] = useState('');
  const [selectingManualCode, setSelectingManualCode] = useState<string | null>(null);

  // --- Verification pre-traitement ---
  // Elle controle les prerequis fichier avant de lancer une operation longue.
  // Le vrai traitement metier reste cote Python dans `vte.domain.iris`.
  const [verifyMessage, setVerifyMessage] = useState('');
  const [verifyStatus, setVerifyStatus] = useState<RunStatus>('idle');

  // --- Traitement et progression ---
  // `running`: un job est lance ou en cours de polling.
  // `done`: le panneau de resultat reste visible apres la fin.
  // `progress`: valeur animee affichee dans la barre.
  // `targetRef`: valeur cible envoyee par le backend.
  const [running, setRunning] = useState(false);
  const [done, setDone] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentStep, setCurrentStep] = useState('');
  const [stepsLog, setStepsLog] = useState<string[]>([]);
  const [treatmentMessage, setTreatmentMessage] = useState('');
  const [treatmentStatus, setTreatmentStatus] = useState<RunStatus>('idle');

  // References techniques conservees hors rendu React:
  // - AbortController pour annuler les fetch/polling;
  // - job id courant pour demander aussi l'annulation cote backend.
  const abortRef = useRef<AbortController | null>(null);
  const currentJobIdRef = useRef<string | null>(null);
  const manualToastTimeoutRef = useRef<number | null>(null);
  // Cible de progression: stockee en ref pour que le timer lise toujours la
  // derniere valeur sans recreer un interval a chaque progression recue.
  const targetRef = useRef(0);

  const setTargetProgress = (val: number) => { targetRef.current = Math.min(val, 100); };

  // Quand le traitement se termine, on aligne immediatement l'affichage sur la
  // valeur finale pour eviter une barre qui continue a rattraper apres le message.
  useEffect(() => {
    if (done) setProgress(Math.min(targetRef.current, 100));
  }, [done]);

  // Timer unique stable:
  // Le back envoie parfois des sauts de progression. On anime `progress` vers
  // `targetRef.current` pour garder une barre fluide sans multiplier les timers.
  useEffect(() => {
    if (!running) return;
    const id = setInterval(() => {
      setProgress(prev => {
        const t = targetRef.current;
        if (prev === t) return prev;
        if (prev > t) return t; // descend immédiatement si reset
        const step = Math.max(0.3, (t - prev) * 0.05);
        return Math.min(Math.round((prev + step) * 10) / 10, Math.min(t, 100));
      });
    }, 300);
    return () => clearInterval(id);
  }, [running]); // ne redémarre QUE quand running change

  // =============================================================================
  // Chargement initial
  // =============================================================================

  const loadDumps = async (preserveResults = false) => {
    // `preserveResults` sert apres un traitement: on veut rafraichir le tableau
    // des dumps, mais conserver la liste des fichiers produits par le job.
    setLoadingDumps(true);
    setLoadError('');
    const ctrl = new AbortController();
    // Timeout court: si FastAPI demarre encore, on affichera un message et le
    // retry automatique relancera le chargement quelques secondes plus tard.
    const timeout = setTimeout(() => ctrl.abort(), 5000);
    try {
      const response = await fetch(`${API_BASE_URL}/iris/dumps`, { signal: ctrl.signal });
      if (!response.ok) throw new Error(`Serveur : erreur ${response.status}`);
      const data = await response.json() as DumpsResponse;
      // La verification appartient a l'execution courante du front, pas au back.
      // On la remet donc a zero quand on recharge les lignes.
      const rows = data.dumps.map(dump => ({ ...dump, verification: 'idle' as const, verificationMessage: '' }));
      setDumps(rows);
      if (!preserveResults) setResultRows(data.outputFiles ?? []);
      setChecks(Object.fromEntries(rows.map(dump => [dump.code.toLowerCase(), dump.selected])));
    } catch (error) {
      // Erreur volontairement douce: en local, le cas le plus courant est que le
      // backend n'a pas encore fini de demarrer.
      setDumps([]);
      setResultRows([]);
      setChecks({});
      const msg = 'Backend en cours de démarrage... reconnexion automatique.';
      setLoadError(msg);
    } finally {
      clearTimeout(timeout);
      setLoadingDumps(false);
    }
  };

  // Chargement initial + retry auto toutes les 5s si le backend est absent.
  useEffect(() => { void loadDumps(); }, []);
  useEffect(() => {
    return () => {
      if (manualToastTimeoutRef.current !== null) {
        window.clearTimeout(manualToastTimeoutRef.current);
      }
    };
  }, []);
  useEffect(() => {
    if (!loadError) return;
    const id = setTimeout(() => void loadDumps(), 5000);
    return () => clearTimeout(id);
  }, [loadError]);

  // Reprise automatique si la page est rechargee pendant un traitement.
  // Le job reste en memoire cote FastAPI; le front garde seulement son id dans
  // sessionStorage pour reprendre le polling apres un refresh.
  useEffect(() => {
    const jobId = sessionStorage.getItem('iris_pending_job');
    if (!jobId) return;

    const ctrl = new AbortController();
    abortRef.current = ctrl;
    setRunning(true);
    setDone(false);
    setTargetProgress(10);
    setCurrentStep('Reprise du traitement en cours...');

    pollJobStatus(jobId, ctrl.signal, (p, s) => { setTargetProgress(p); setCurrentStep(s); }, applyStatuts, setStepsLog)
      .then(job => {
        setProgress(100);
        setTargetProgress(100);
        setDone(true);
        if (job.status === 'cancelled') {
          setTreatmentStatus('error');
          setTreatmentMessage(job.message ?? 'Traitement annulé — tu peux relancer.');
          setCurrentStep('Annulé');
        } else {
          setTreatmentStatus(job.warning ? 'error' : 'success');
          setTreatmentMessage(job.message ?? 'Traitement terminé.');
          setCurrentStep(job.warning ? 'Erreur' : 'Traitement terminé ✓');
        }
        setResultRows(job.files ?? []);
        if (job.status !== 'cancelled' && !job.warning) void loadDumps(true);
      })
      .catch(err => {
        if ((err as Error).name !== 'AbortError') {
          setDone(true);
          setTargetProgress(100);
          setTreatmentStatus('error');
          setTreatmentMessage((err as Error).message);
          setCurrentStep('Erreur');
        }
      })
      .finally(() => {
        sessionStorage.removeItem('iris_pending_job'); currentJobIdRef.current = null;
        setRunning(false);
      });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // =============================================================================
  // Handlers sélection
  // =============================================================================

  // Convertit la selection locale en codes IRIS attendus par l'API.
  const getSelectedExportCodes = () =>
    Object.entries(checks).filter(([, s]) => s).map(([c]) => c.toUpperCase());

  // Applique les statuts par code recus pendant le polling du job.
  const applyStatuts = (statuts: Record<string, string>) => {
    setDumps(prev => prev.map(dump => ({
      ...dump,
      statut: statuts[dump.code] ?? dump.statut,
    })));
  };

  // Remet a zero uniquement l'etat d'execution. Les fichiers selectionnes et le
  // mode restent inchanges pour ne pas surprendre l'utilisateur.
  const resetExecutionState = () => {
    setDone(false);
    setProgress(0);
    setTargetProgress(0);
    setCurrentStep('');
    setStepsLog([]);
    setTreatmentMessage('');
    setTreatmentStatus('idle');
    setResultRows([]);
  };

  const showManualToast = (message: string) => {
    // Notification non bloquante: elle remplace l'ancienne fenetre `alert()`.
    // Le fichier reste refuse, mais l'utilisateur peut continuer a utiliser la page.
    if (manualToastTimeoutRef.current !== null) {
      window.clearTimeout(manualToastTimeoutRef.current);
    }
    setManualToast(message);
    manualToastTimeoutRef.current = window.setTimeout(() => {
      setManualToast('');
      manualToastTimeoutRef.current = null;
    }, 4500);
  };

  const chooseManualFiles = async (dump: DumpRow) => {
    // Le backend ouvre une vraie fenetre Python/Tkinter locale avec un titre
    // personnalise. Cela evite le selecteur Chrome intitule "Ouvrir".
    resetExecutionState();
    setManualToast('');
    setSelectingManualCode(dump.code);

    try {
      const response = await fetch(`${API_BASE_URL}/iris/select-manual-files`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ export: dump.code }),
      });

      if (!response.ok) {
        const errPayload = await response.json().catch(() => null);
        const raw = errPayload?.detail;
        throw new Error(typeof raw === 'string' ? raw : raw != null ? JSON.stringify(raw) : response.statusText);
      }

      const data = await response.json() as { files: ManualSelectedFile[] };
      setManualFiles(prev => ({ ...prev, [dump.code]: data.files }));

      if (data.files.length === 0) {
        showManualToast(`Aucun fichier selectionne pour ${dump.type} (${dump.code}).`);
      } else {
        setManualToast('');
      }
    } catch (error) {
      showManualToast(error instanceof Error ? error.message : 'Selection de fichiers impossible.');
    } finally {
      setSelectingManualCode(null);
    }
  };

  const toggleDumpSelect = (index: number) => {
    // La ligne du tableau et le dictionnaire `checks` doivent rester synchrones:
    // la ligne sert au rendu, `checks` sert aux appels API.
    setDumps(prev => prev.map((dump, i) => {
      if (i !== index) return dump;
      const selected = !dump.selected;
      setChecks(current => ({ ...current, [dump.code.toLowerCase()]: selected }));
      return { ...dump, selected };
    }));
    resetExecutionState();
  };

  // =============================================================================
  // Vérification (appelée automatiquement avant traitement)
  // =============================================================================

  const verifyFileStructure = async (signal?: AbortSignal): Promise<boolean> => {
    // On verifie seulement les exports coches. Le backend renvoie ensuite un
    // resultat detaille par code pour colorer chaque ligne du tableau.
    const selectedExports = getSelectedExportCodes();
    if (selectedExports.length === 0) return false;

    setVerifyStatus('idle');
    setVerifyMessage('');
    setCurrentStep('Vérification des fichiers sources...');
    setDumps(prev => prev.map(dump => ({
      ...dump,
      verification: selectedExports.includes(dump.code) ? 'idle' : dump.verification,
      verificationMessage: selectedExports.includes(dump.code) ? 'Vérification en cours...' : dump.verificationMessage,
    })));

    try {
      // Validation rapide cote backend: presence, extension, taille et lecture
      // basique. Le contenu metier detaille reste traite par les classes IRIS.
      const response = await fetch(`${API_BASE_URL}/iris/validate-structure`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ codes: selectedExports }),
        signal,
      });
      if (!response.ok) throw new Error(await response.text());

      const data = await response.json() as VerificationResponse;
      // Map pour appliquer facilement le resultat de chaque code sur sa ligne.
      const checksByCode = new Map(data.checks.map(check => [check.code, check]));

      setVerifyStatus(data.valid ? 'success' : 'error');
      setVerifyMessage(data.message);
      setDumps(prev => prev.map(dump => ({
        ...dump,
        verification: selectedExports.includes(dump.code)
          ? checksByCode.get(dump.code)?.status ?? 'error'
          : dump.verification,
        verificationMessage: selectedExports.includes(dump.code)
          ? checksByCode.get(dump.code)?.message ?? 'Vérification absente'
          : dump.verificationMessage,
      })));

      if (data.valid) {
        setCurrentStep('Fichiers vérifiés — prêt à traiter');
      } else {
        setCurrentStep('Vérification échouée — corrigez les erreurs avant de relancer');
      }

      return data.valid;
    } catch (error) {
      // Si l'utilisateur annule, on propage l'AbortError pour que le traitement
      // principal affiche "Arrete" au lieu d'une erreur technique.
      if ((error as Error).name === 'AbortError') throw error;
      setVerifyStatus('error');
      setVerifyMessage(error instanceof Error ? error.message : 'Erreur pendant la vérification IRIS.');
      setCurrentStep('Erreur lors de la vérification');
      setDumps(prev => prev.map(dump => ({
        ...dump,
        verification: dump.selected ? 'error' : dump.verification,
        verificationMessage: dump.selected ? 'Erreur de vérification' : dump.verificationMessage,
      })));
      return false;
    } finally {
      }
  };

  // =============================================================================
  // Arrêter le traitement
  // =============================================================================

  const stopTreatment = () => {
    // Signale l'annulation au backend en premier (arrêt propre du thread)
    const jobId = currentJobIdRef.current;
    if (jobId) {
      fetch(`${API_BASE_URL}/iris/job/${jobId}`, { method: 'DELETE' }).catch(() => {});
      currentJobIdRef.current = null;
    }
    // Puis annule les requetes frontend encore ouvertes.
    abortRef.current?.abort();
    sessionStorage.removeItem('iris_pending_job'); currentJobIdRef.current = null;
  };

  // =============================================================================
  // Lancer le traitement
  // =============================================================================

  const launchTreatment = async () => {
    // Blocage anti double-clic: un second lancement pendant un job creerait deux
    // traitements concurrents sur les memes fichiers Excel.
    if (running) {
      alert('Un traitement est déjà en cours. Attendez la fin ou cliquez sur Arrêter.');
      return;
    }

    const selectedExports = getSelectedExportCodes();
    if (selectedExports.length === 0) {
      alert('Aucun export sélectionné pour le traitement.');
      return;
    }

    if (mode === 'manual') {
      // En mode manuel, chaque export coche doit recevoir au moins un fichier.
      // Le back controle aussi, mais on evite un aller-retour inutile.
      const manquants = dumps.filter(d => d.selected && !(manualFiles[d.code]?.length > 0));
      if (manquants.length > 0) {
        alert(`Fichiers manquants pour : ${manquants.map(d => d.type).join(', ')}`);
        return;
      }
    }

    // Cree un AbortController pour permettre l'arret propre du flux complet:
    // verification, POST de lancement, puis polling du job.
    const controller = new AbortController();
    abortRef.current = controller;

    setRunning(true);
    setDone(false);
    setProgress(0);
    setTargetProgress(5);
    setStepsLog([]);
    setTreatmentStatus('idle');
    setTreatmentMessage('');
    setCurrentStep('Démarrage...');
    setResultRows([]);

    try {
      // Etape 1: verification auto uniquement en mode auto.
      // En mode manuel, les fichiers ont deja ete valides lors de la selection
      // (Tkinter + _valider_fichier_manuel cote back). Appeler validate-structure
      // ici bloquerait a tort si un fichier de la config auto est absent alors
      // que l'utilisateur a fourni ses propres chemins.
      if (mode === 'auto') {
        setTargetProgress(10);
        const verificationOk = await verifyFileStructure(controller.signal);
        if (!verificationOk) {
          setDone(true);
          setTargetProgress(100);
          setTreatmentStatus('error');
          setTreatmentMessage('Traitement bloqué — corrigez les fichiers indiqués dans la colonne Vérification.');
          return;
        }
        // La vérification a servi de garde avant lancement. Dès que le traitement
        // démarre réellement, on retire son bandeau pour ne garder que le suivi job.
        setVerifyStatus('idle');
        setVerifyMessage('');
      }

      setTargetProgress(25);

      // Etape 2: traitement.
      // Mode manuel: on traite les exports un par un pour associer les chemins
      // choisis par la fenetre Python a un code IRIS precis.
      if (mode === 'manual') {
        const selectedDumps = dumps.filter(d => d.selected);
        const allFiles: ResultRow[] = [];

        for (let i = 0; i < selectedDumps.length; i++) {
          const dump = selectedDumps[i];
          const stepBase = 25 + Math.round((i / selectedDumps.length) * 70);
          setTargetProgress(stepBase);
          setCurrentStep(`${dump.type} — transmission des chemins selectionnes...`);

          setTargetProgress(stepBase + 5);
          setCurrentStep(`${dump.type} — lecture et concaténation des fichiers Excel...`);

          const resp = await fetch(`${API_BASE_URL}/iris/traiter-manuel-chemins`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              export: dump.code,
              paths: (manualFiles[dump.code] ?? []).map(file => file.path),
            }),
            signal: controller.signal,
          });

          if (!resp.ok) {
            const errPayload = await resp.json().catch(() => null);
            const raw = errPayload?.detail;
            const msg = typeof raw === 'string'
              ? raw
              : raw != null ? JSON.stringify(raw) : resp.statusText || `Erreur ${resp.status}`;
            throw new Error(msg);
          }

          const { job_id } = await resp.json() as { job_id: string };
          // On memorise le job pour pouvoir reprendre le polling apres refresh
          // ou demander une annulation serveur.
          currentJobIdRef.current = job_id;
          sessionStorage.setItem('iris_pending_job', job_id);

          const job = await pollJobStatus(job_id, controller.signal, (p, s) => {
            // Chaque job manuel represente une portion de la progression totale.
            // On remappe donc la progression du back dans la plage de l'export.
            setTargetProgress(stepBase + Math.round(p * 0.3));
            setCurrentStep(s);
          }, applyStatuts, setStepsLog);

          sessionStorage.removeItem('iris_pending_job'); currentJobIdRef.current = null;
          if (job.status === 'cancelled') {
            setTargetProgress(100);
            setDone(true);
            setTreatmentStatus('error');
            setTreatmentMessage(job.message ?? 'Traitement annulé — tu peux relancer.');
            setCurrentStep('Annulé');
            setResultRows(allFiles);
            return;
          }
          allFiles.push(...(job.files ?? []));
          setCurrentStep(`${dump.type} — terminé ✓`);
        }

        setTargetProgress(100);
        setDone(true);
        setTreatmentStatus('success');
        setTreatmentMessage('Traitement manuel terminé avec succès.');
        setCurrentStep('Tous les exports traités ✓');
        setResultRows(allFiles);
        void loadDumps(true);
        return;
      }

      // Mode auto: le back choisit les fichiers depuis la configuration IRIS et
      // traite tous les exports selectionnes dans un seul job.
      setCurrentStep('Envoi de la demande au serveur...');
      setTargetProgress(20);

      const response = await fetch(`${API_BASE_URL}/iris/traiter`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ exports: selectedExports, format: 'xlsx', mode }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const errorPayload = await response.json().catch(() => null);
        const raw = errorPayload?.detail;
        throw new Error(typeof raw === 'string' ? raw : raw != null ? JSON.stringify(raw) : response.statusText);
      }

      const { job_id } = await response.json() as { job_id: string };
      // Meme mecanique que le mode manuel: id en ref pour l'annulation, id en
      // sessionStorage pour la reprise apres rechargement de page.
      currentJobIdRef.current = job_id;
      sessionStorage.setItem('iris_pending_job', job_id);

      setCurrentStep('Traitement en cours côté serveur...');
      setTargetProgress(30);

      const job = await pollJobStatus(job_id, controller.signal, (p, s) => {
        setTargetProgress(p);
        setCurrentStep(s);
      }, applyStatuts, setStepsLog);

      sessionStorage.removeItem('iris_pending_job'); currentJobIdRef.current = null;

      setTargetProgress(100);
      setDone(true);
      if (job.status === 'cancelled') {
        setTreatmentStatus('error');
        setTreatmentMessage(job.message ?? 'Traitement annulé — tu peux relancer.');
        setCurrentStep('Annulé');
        setResultRows(job.files ?? []);
        return;
      }
      setTreatmentStatus(job.warning ? 'error' : 'success');
      setTreatmentMessage(job.message ?? 'Traitement IRIS terminé.');
      setCurrentStep(job.warning ? 'Erreur' : 'Traitement terminé ✓');
      setResultRows(job.files ?? []);
      if (!job.warning) void loadDumps(true);

    } catch (error) {
      // Nettoyage sessionStorage dans tous les cas d'erreur: si on laisse un id
      // obsolete, le prochain chargement essaiera de reprendre un job invalide.
      sessionStorage.removeItem('iris_pending_job'); currentJobIdRef.current = null;

      if ((error as Error).name === 'AbortError') {
        setTargetProgress(100);
        setDone(true);
        setTreatmentStatus('error');
        setTreatmentMessage('Traitement arrêté — tu peux relancer.');
        setCurrentStep('Arrêté');
      } else if (
        error instanceof TypeError &&
        (error.message === 'Failed to fetch' || error.message.includes('fetch'))
      ) {
        setTargetProgress(100);
        setDone(true);
        setTreatmentStatus('error');
        setTreatmentMessage('Connexion perdue — relance le serveur puis réessaie.');
        setCurrentStep('Connexion perdue');
      } else {
        setTargetProgress(100);
        setDone(true);
        setTreatmentStatus('error');
        setTreatmentMessage(error instanceof Error ? error.message : 'Erreur pendant le traitement IRIS.');
        setCurrentStep('Erreur');
      }
    } finally {
      setRunning(false);
    }
  };

  // =============================================================================
  // Rendu
  // =============================================================================

  // Les resultats affiches sont filtres sur la selection courante pour eviter de
  // montrer d'anciens fichiers d'un export que l'utilisateur vient de decocher.
  const selectedCodesForRender = getSelectedExportCodes();
  const visibleResultRows = resultRows.filter(row =>
    selectedCodesForRender.some(code => row.export.startsWith(code) || row.fichier.startsWith(code))
  );

  return (
    <div className="space-y-4">
      {manualToast && (
        <div className="fixed right-6 top-24 z-50 flex max-w-md items-start gap-2 rounded-lg border border-red-100 bg-red-50 px-4 py-3 text-xs font-semibold leading-relaxed text-red-800 shadow-lg">
          <XCircle size={15} className="mt-0.5 flex-shrink-0" />
          <span>{manualToast}</span>
        </div>
      )}

      <Card noPad>
        <CardHeader title="Gestion des exports IRIS" step={1} />

        <div className="p-4 space-y-4">

          {/* Tableau dumps */}
          <div className="overflow-hidden rounded-xl border border-slate-100">
            <table className="w-full">
              <thead>
                <tr style={{ background: '#f8fafc', borderBottom: '1px solid #f1f5f9' }}>
                  <th className="px-4 py-2.5 text-left text-xs font-bold uppercase tracking-widest text-slate-400 w-12">Choisir</th>
                  <th className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest text-slate-400">Code</th>
                  <th className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest text-slate-400">Type</th>
                  <th className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest text-slate-400">Fichier cible</th>
                  <th className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest text-slate-400">Dernière mise à jour</th>
                  <th className="text-left px-4 py-2.5 text-xs font-bold uppercase tracking-widest text-slate-400">Statut</th>
                </tr>
              </thead>
              <tbody>
                {dumps.map((row, index) => (
                  <tr key={row.code} style={{ borderBottom: index < dumps.length - 1 ? '1px solid #f8fafc' : 'none' }} className="transition-colors hover:bg-slate-50">
                    <td className="px-4 py-3 text-center">
                      <input
                        type="checkbox"
                        checked={row.selected}
                        onChange={() => toggleDumpSelect(index)}
                        disabled={running}
                        className="w-4 h-4 rounded text-orange-500"
                        style={{ accentColor: '#f97316' }}
                      />
                    </td>
                    <td className="px-4 py-3 font-mono text-xs font-bold" style={{ color: '#1e293b' }}>{row.code}</td>
                    <td className="px-4 py-3 text-sm font-medium" style={{ color: '#1e293b' }}>{row.type}</td>
                    <td className={`px-4 py-3 font-mono text-xs text-slate-500 ${running ? 'opacity-40 pointer-events-none' : ''}`}>
                      <PathLink href={row.targetUrl} openTarget={row.targetPath} label={row.fichier} />
                    </td>
                    <td className="px-4 py-3 font-mono text-xs text-slate-500">{row.updatedAt}</td>
                    <td className="px-4 py-3"><StatusCell row={row} /></td>
                  </tr>
                ))}
              </tbody>
            </table>

            {loadingDumps && (
              <div className="p-4 text-xs font-semibold text-slate-400 bg-slate-50 flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-slate-400 animate-pulse" />
                Connexion au backend...
              </div>
            )}
          </div>


          {/* Message vérification global */}
          {verifyMessage && !running && (
            <div className={`p-3 rounded-xl text-xs border font-semibold leading-relaxed ${verifyStatus === 'success' ? 'bg-emerald-50 text-emerald-800 border-emerald-100' : 'bg-red-50 text-red-800 border-red-100'}`}>
              {verifyMessage}
            </div>
          )}

          {/* Barre mode + zones selection manuel + boutons */}
          <div className="flex flex-wrap items-start gap-3 rounded-xl border border-slate-100 bg-white p-3">

            {/* Toggle Auto / Manuel */}
            <div className="flex items-center rounded-lg border border-slate-200 bg-slate-50 p-1 self-start">
              {([
                { value: 'auto' as const, label: 'Auto' },
                { value: 'manual' as const, label: 'Manuel' },
              ] as const).map(({ value, label }) => (
                <button
                  key={value}
                  onClick={() => { setMode(value); resetExecutionState(); }}
                  disabled={running}
                  className="min-w-20 rounded-md px-3 py-1.5 text-center text-xs font-bold transition-all"
                  style={{
                    background: mode === value ? '#fff' : 'transparent',
                    color: mode === value ? '#c2410c' : '#64748b',
                    boxShadow: mode === value ? '0 1px 2px rgba(15,23,42,0.08)' : 'none',
                  }}
                >
                  {label}
                </button>
              ))}
            </div>

            {/* Zones de selection en mode manuel */}
            {mode === 'manual' && (
              <div className="flex flex-col gap-2 flex-1">
                {dumps.filter(d => d.selected).length === 0 && (
                  <p className="text-xs text-slate-400">Cochez au moins un export dans le tableau.</p>
                )}
                {dumps.filter(d => d.selected).map(dump => (
                  <button
                    key={dump.code}
                    type="button"
                    onClick={() => void chooseManualFiles(dump)}
                    disabled={running}
                    className="flex cursor-pointer items-center gap-3 rounded-lg border border-dashed border-orange-200 bg-orange-50 px-3 py-2 text-left hover:bg-orange-100 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    <FileUp size={15} className="flex-shrink-0 text-orange-600" />
                    <div className="flex-1 min-w-0">
                      <p className="text-xs font-bold text-slate-700">{dump.type} ({dump.code})</p>
                      <p className="text-xs text-slate-500 truncate">
                        {selectingManualCode === dump.code
                          ? 'Sélection en cours...'
                          : (manualFiles[dump.code] ?? []).length > 0
                          ? `${(manualFiles[dump.code] ?? []).length} fichier(s) sélectionné(s)`
                          : `Choisir uniquement les fichiers ${dump.code}...`}
                      </p>
                    </div>
                  </button>
                ))}
              </div>
            )}

            {/* Boutons Traiter + Arrêter */}
            <div className="flex items-center gap-2 self-start ml-auto">
              <Btn
                onClick={launchTreatment}
                disabled={running || dumps.length === 0}
                size="sm"
              >
                <Play size={15} className={running ? 'animate-pulse' : ''} />
                {running ? 'En cours...' : 'Traiter'}
              </Btn>
              {running && (
                <Btn onClick={stopTreatment} size="sm" variant="secondary">
                  <Square size={14} />
                  Arrêter
                </Btn>
              )}
            </div>
          </div>

          {/* Zone progression + étapes + résultats */}
          {(running || done) && (
            <div className="space-y-3 rounded-xl border border-slate-100 bg-slate-50 p-4">

              {/* Barre de progression */}
              <div className="flex items-center justify-between text-xs">
                <span className="font-semibold text-slate-600">
                  {done ? (treatmentStatus === 'error' ? 'Terminé avec erreur' : 'Terminé') : 'En cours...'}
                </span>
                <span className="font-bold text-orange-600">{progress}%</span>
              </div>
              {!done && currentStep && (
                <p className="text-xs font-mono text-slate-400 truncate">{currentStep}</p>
              )}
              <div className="h-2 w-full overflow-hidden rounded-full bg-slate-200">
                <div
                  className="h-full rounded-full transition-all duration-500"
                  style={{
                    width: `${progress}%`,
                    background: treatmentStatus === 'error' ? '#dc2626' : done ? '#16a34a' : '#f97316',
                  }}
                />
              </div>

              {/* 2 dernières étapes — disparaissent dès que c'est terminé */}
              {stepsLog.length > 0 && !done && (
                <div className="rounded-lg border border-slate-100 bg-white p-2 space-y-0.5" style={{ fontFamily: 'monospace' }}>
                  {stepsLog.slice(-2).map((step, i, arr) => {
                    const isCurrent = i === arr.length - 1 && !done;
                    return (
                      <div key={stepsLog.length - arr.length + i} className="flex items-center gap-2 text-xs py-0.5">
                        {isCurrent
                          ? <span className="w-1.5 h-1.5 rounded-full bg-orange-400 animate-pulse flex-shrink-0" />
                          : <span className="text-emerald-500 flex-shrink-0">✓</span>
                        }
                        <span className={isCurrent ? 'text-orange-700 font-semibold' : 'text-slate-400'}>
                          {step}
                        </span>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* Message final */}
              {done && treatmentMessage && (
                <div className={`flex items-center gap-2 rounded-lg border px-3 py-2 text-xs font-semibold ${treatmentStatus === 'error' ? 'border-red-100 bg-red-50 text-red-800' : 'border-emerald-100 bg-emerald-50 text-emerald-800'}`}>
                  {treatmentStatus === 'error' ? <XCircle size={14} /> : <CheckCircle size={14} />}
                  <span>{treatmentMessage}</span>
                </div>
              )}

              {/* Fichiers produits */}
              {done && treatmentStatus === 'success' && visibleResultRows.length > 0 && (
                <div className="overflow-hidden rounded-lg border border-slate-100 bg-white">
                  <div className="flex items-center justify-between border-b border-slate-100 px-4 py-2.5">
                    <p className="text-xs font-bold uppercase tracking-widest text-slate-400">Résultats</p>
                    <Badge variant="gray">{visibleResultRows.length} fichier(s)</Badge>
                  </div>
                  <div className="divide-y divide-slate-100">
                    {visibleResultRows.map((row, index) => (
                      <div key={`${row.fichier}-${index}`} className="grid grid-cols-[1fr_2fr_auto_auto] items-center gap-3 px-4 py-3">
                        <p className="text-xs font-bold text-slate-700">{row.export}</p>
                        <PathLink href={row.url} openTarget={row.path} label={row.fichier} />
                        <p className="whitespace-nowrap text-xs font-mono text-slate-400">{row.date}</p>
                        <Badge variant={row.statut === 'ok' ? 'ok' : 'warning'}>
                          {row.statut === 'ok' ? 'OK' : 'Warning'}
                        </Badge>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}

        </div>
      </Card>
    </div>
  );
}
