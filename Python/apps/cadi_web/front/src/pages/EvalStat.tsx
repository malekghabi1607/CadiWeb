import { useEffect, useMemo, useState } from 'react';
import { Loader2, Play, RefreshCw } from 'lucide-react';
import { API_BASE_URL } from '../lib/api';
import { Badge, Btn, Card, CardHeader, Field, Input, TD, THead, TRow } from '../components/ui';

type Status = 'idle' | 'success' | 'warning' | 'error';
type OpenAfterMode = 'auto' | 'excel' | 'folder' | 'none';

type SessionRow = {
  codeIris: number;
  formation: string;
  trigrammeRp: string;
  sessionRef: string;
  sessionTitle: string;
  sessionStatus: string;
  startDate: string;
  endDate: string;
  nbNommes: number | null;
  csvPath: string;
  csvUrl: string;
  csvExists: boolean;
};

type SessionsResponse = {
  formation: string;
  sessions: SessionRow[];
  irisFile: string | null;
  message?: string | null;
  searchDirectory?: string | null;
};

type FormationsResponse = {
  formations: string[];
};

type ProcessRowResult = {
  success: boolean;
  status: string;
  message: string;
  codeIris: number | null;
  trigramme: string | null;
  files?: {
    csv?: { path: string; url: string; name: string };
    sessionExcel?: { path: string; url: string; name: string };
    formationExcel?: { path: string; url: string; name: string };
  };
};

type ProcessResponse = {
  success: boolean;
  rows: ProcessRowResult[];
};

type ChooseCsvResponse = {
  formation: string;
  session: SessionRow;
  irisFile: string | null;
};

async function apiGet<T>(path: string): Promise<T> {
  const res = await fetch(`${API_BASE_URL}${path}`);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const detail = typeof data.detail === 'string' ? data.detail : data.detail?.message || data.message;
    throw new Error(detail || `Erreur API ${res.status}`);
  }
  return data;
}

async function apiPost<T>(path: string, body: object): Promise<T> {
  const res = await fetch(`${API_BASE_URL}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const detail = typeof data.detail === 'string' ? data.detail : data.detail?.message || data.message;
    throw new Error(detail || `Erreur API ${res.status}`);
  }
  return data;
}

function formatDate(isoDate: string) {
  if (!isoDate) return '-';
  const [year, month, day] = isoDate.split('-');
  return `${day}/${month}/${year}`;
}

function parentPath(path: string) {
  const index = Math.max(path.lastIndexOf('\\'), path.lastIndexOf('/'));
  return index > 0 ? path.slice(0, index) : path;
}

function splitPath(path: string) {
  return path.split(/[\\/]/).filter(Boolean);
}

function commonDirectory(paths: string[]) {
  if (paths.length === 0) return '';
  const directories = paths.map(path => splitPath(parentPath(path)));
  const first = directories[0];
  let commonLength = first.length;

  for (const parts of directories.slice(1)) {
    commonLength = Math.min(commonLength, parts.length);
    for (let index = 0; index < commonLength; index += 1) {
      if (parts[index].toLowerCase() !== first[index].toLowerCase()) {
        commonLength = index;
        break;
      }
    }
  }

  if (commonLength === 0) return parentPath(paths[0]);
  const isUnc = paths[0].startsWith('\\\\');
  const isDrive = /^[A-Za-z]:/.test(paths[0]);
  const prefix = isUnc ? '\\\\' : isDrive ? `${paths[0].slice(0, 2)}\\` : '';
  return `${prefix}${first.slice(isDrive ? 1 : 0, commonLength).join('\\')}`;
}

function treatmentBadge(row: SessionRow, result: ProcessRowResult | undefined, loading: string, selectedCodes: Set<number>) {
  if (result) {
    if (result.success) return { label: 'Traite', variant: 'ok' as const };
    if (result.status.toLowerCase().includes('vide')) return { label: 'Ignore', variant: 'warning' as const };
    return { label: 'Erreur', variant: 'error' as const };
  }
  if (loading === 'process' && selectedCodes.has(row.codeIris)) {
    return { label: 'En cours', variant: 'warning' as const };
  }
  if (selectedCodes.has(row.codeIris)) {
    return { label: 'En attente', variant: 'gray' as const };
  }
  return { label: 'Non coche', variant: 'gray' as const };
}

function isIgnoredResult(row: ProcessRowResult) {
  const text = `${row.status} ${row.message}`.toLowerCase();
  return text.includes('vide') || text.includes('ignore') || text.includes('ignor') || text.includes('exclu');
}

function treatmentSummary(rows: ProcessRowResult[]) {
  const treated = rows.filter(row => row.success).length;
  const ignored = rows.filter(row => !row.success && isIgnoredResult(row)).length;
  const errors = rows.length - treated - ignored;
  return `${treated} traite(s), ${ignored} ignore(s), ${errors} erreur(s).`;
}

export default function EvalStat() {
  const [loading, setLoading] = useState('');
  const [loadingElapsed, setLoadingElapsed] = useState(0);
  const [status, setStatus] = useState<Status>('idle');
  const [message, setMessage] = useState('Saisissez un trigramme de formation.');
  const [formations, setFormations] = useState<string[]>([]);
  const [formationFilter, setFormationFilter] = useState('');
  const [showFormationSuggestions, setShowFormationSuggestions] = useState(false);
  const [sessions, setSessions] = useState<SessionRow[]>([]);
  const [selectedCodes, setSelectedCodes] = useState<Set<number>>(new Set());
  const [processResults, setProcessResults] = useState<Map<number, ProcessRowResult>>(new Map());
  const [openAfterMode, setOpenAfterMode] = useState<OpenAfterMode>('auto');
  const [searchDirectory, setSearchDirectory] = useState('');

  const formationSuggestions = useMemo(() => {
    const query = formationFilter.trim().toUpperCase();
    if (!query) return [];
    return formations.filter(formation => formation.startsWith(query)).slice(0, 12);
  }, [formations, formationFilter]);

  const visibleSessions = sessions.slice(0, 500);
  const selectableRows = visibleSessions.filter(row => row.csvExists);
  const selectedRows = visibleSessions.filter(row => row.csvExists && selectedCodes.has(row.codeIris));
  const allSelected = selectableRows.length > 0 && selectedRows.length === selectableRows.length;

  useEffect(() => {
    if (!loading) return;

    setLoadingElapsed(0);
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      setLoadingElapsed(Math.floor((Date.now() - startedAt) / 1000));
    }, 500);

    return () => window.clearInterval(timer);
  }, [loading]);

  useEffect(() => {
    const loadFormations = async () => {
      setLoading('formations');
      try {
        const data = await apiGet<FormationsResponse>('/evalstat/formations');
        setFormations(data.formations);
        setMessage('Saisissez un trigramme de formation.');
      } catch (error) {
        setStatus('error');
        setMessage(error instanceof Error ? error.message : 'Chargement des formations impossible.');
      } finally {
        setLoading('');
      }
    };
    void loadFormations();
  }, []);

  const searchSessions = async () => {
    const trigramme = formationFilter.trim().toUpperCase();
    if (!trigramme) {
      setStatus('warning');
      setMessage('Saisissez un trigramme de formation avant de rechercher.');
      return;
    }

    setShowFormationSuggestions(false);
    setLoading('search');
    setProcessResults(new Map());
    try {
      const data = await apiGet<SessionsResponse>(`/evalstat/sessions?formation=${encodeURIComponent(trigramme)}`);
      setSessions(data.sessions);
      setSearchDirectory(data.searchDirectory || '');
      setSelectedCodes(new Set(data.sessions.filter(row => row.csvExists).map(row => row.codeIris)));
      setStatus(data.sessions.length > 0 ? 'success' : 'warning');
      setMessage(
        data.sessions.length > 0
          ? `${data.sessions.length} CSV EvalStat trouve(s) pour ${trigramme}.`
          : data.message || `Aucun CSV EvalStat trouve pour ${trigramme} dans FORMATIONS_C.`
      );
    } catch (error) {
      setStatus('error');
      setMessage(error instanceof Error ? error.message : 'Chargement des sessions impossible.');
    } finally {
      setLoading('');
    }
  };

  const toggleRow = (row: SessionRow) => {
    if (!row.csvExists || loading === 'process') return;
    setSelectedCodes(codes => {
      const next = new Set(codes);
      if (next.has(row.codeIris)) {
        next.delete(row.codeIris);
      } else {
        next.add(row.codeIris);
      }
      return next;
    });
  };

  const toggleAll = () => {
    if (loading === 'process') return;
    setSelectedCodes(allSelected ? new Set() : new Set(selectableRows.map(row => row.codeIris)));
  };

  const openPath = async (path: string) => {
    if (!path) return;
    const response = await fetch(`${API_BASE_URL}/evalstat/open-path`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path }),
    });
    if (!response.ok) {
      const data = await response.json().catch(() => null);
      throw new Error(data?.detail || response.statusText);
    }
  };

  const openAfterTreatment = async (rows: ProcessRowResult[]) => {
    if (openAfterMode === 'none') return;

    const processedRows = rows.filter(row => row.success && row.files?.sessionExcel?.path);
    if (processedRows.length === 0) return;

    const excelPaths = processedRows
      .map(row => row.files?.sessionExcel?.path || '')
      .filter(Boolean);
    const firstExcel = excelPaths[0] || '';
    const formationExcel = processedRows[0].files?.formationExcel?.path || firstExcel;
    const target =
      openAfterMode === 'folder'
        ? commonDirectory(excelPaths)
        : openAfterMode === 'excel'
        ? processedRows.length === 1
          ? firstExcel
          : formationExcel
        : processedRows.length === 1
        ? firstExcel
        : commonDirectory(excelPaths);

    await openPath(target);
  };

  const uploadManualCsv = async (file: File | undefined) => {
    if (!file) return;

    const trigramme = formationFilter.trim().toUpperCase();
    if (!trigramme) {
      setStatus('warning');
      setMessage('Saisissez un trigramme.');
      return;
    }

    setLoading('manual');
    try {
      const formData = new FormData();
      formData.append('formation', trigramme);
      formData.append('file', file);

      const response = await fetch(`${API_BASE_URL}/evalstat/upload-csv`, {
        method: 'POST',
        body: formData,
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) {
        const detail = typeof data.detail === 'string' ? data.detail : data.detail?.message || data.message;
        throw new Error(detail || `Erreur API ${response.status}`);
      }
      const payload = data as ChooseCsvResponse;

      setSessions(rows => {
        const filtered = rows.filter(row => row.codeIris !== payload.session.codeIris);
        return [...filtered, payload.session].sort((a, b) => b.startDate.localeCompare(a.startDate));
      });
      setSelectedCodes(codes => new Set(codes).add(payload.session.codeIris));
      setStatus('success');
      setMessage('CSV ajoute. Vous pouvez lancer le traitement.');
    } catch (error) {
      setStatus('error');
      setMessage(error instanceof Error ? error.message : 'Choix du CSV impossible.');
    } finally {
      setLoading('');
      const input = document.getElementById('evalstat-manual-csv') as HTMLInputElement | null;
      if (input) input.value = '';
    }
  };

  const processSelected = async () => {
    if (selectedRows.length === 0) {
      setStatus('warning');
      setMessage('Cochez au moins une session avec un CSV trouve avant de traiter.');
      return;
    }

    setLoading('process');
    try {
      const data = await apiPost<ProcessResponse>('/evalstat/process', {
        formation: formationFilter.trim().toUpperCase(),
        paths: selectedRows.map(row => row.csvPath),
      });
      const results = new Map<number, ProcessRowResult>();
      data.rows.forEach(row => {
        if (row.codeIris !== null) results.set(row.codeIris, row);
      });
      setProcessResults(results);
      setStatus(data.success ? 'success' : 'warning');
      setMessage(treatmentSummary(data.rows));
      try {
        await openAfterTreatment(data.rows);
      } catch (openError) {
        setStatus('warning');
        setMessage(
          openError instanceof Error
            ? `${treatmentSummary(data.rows)} Ouverture impossible : ${openError.message}`
            : `${treatmentSummary(data.rows)} Ouverture impossible.`
        );
      }
    } catch (error) {
      setStatus('error');
      setMessage(error instanceof Error ? error.message : 'Traitement impossible.');
    } finally {
      setLoading('');
    }
  };

  const loadingLabel =
    loading === 'formations'
      ? 'Chargement des formations'
      : loading === 'search'
      ? 'Recherche des sessions'
      : loading === 'process'
      ? 'Traitement en cours'
      : loading === 'manual'
      ? 'Choix du CSV'
      : '';

  const displayedMessage = loading
    ? `${loadingLabel}... ${loadingElapsed} s`
    : message;

  return (
    <div className="space-y-4">
      <Card noPad>
        <CardHeader title="Rechercher les sessions" subtitle="Saisir un trigramme de formation puis lancer la recherche." step={1} />
        <div className="p-4 space-y-4">
          {(displayedMessage || loading) && (
            <div
              className={`rounded-xl border px-4 py-3 text-xs font-semibold leading-relaxed ${
                status === 'success'
                  ? 'bg-emerald-50 border-emerald-100 text-emerald-800'
                  : status === 'warning'
                  ? 'bg-amber-50 border-amber-100 text-amber-800'
                  : status === 'error'
                  ? 'bg-red-50 border-red-100 text-red-800'
                  : 'bg-slate-50 border-slate-100 text-slate-600'
              }`}
            >
              <div className="flex items-center justify-between gap-3">
                <span>{displayedMessage}</span>
                {loading && (
                  <Badge variant="gray">
                    <Loader2 size={10} className="animate-spin" />
                    {loadingElapsed} s
                  </Badge>
                )}
              </div>
              {loading && (
                <div className="relative mt-3 h-1.5 overflow-hidden rounded-full bg-white/70">
                  <div className="absolute inset-y-0 left-0 w-2/5 rounded-full bg-orange-500 animate-progress-indeterminate" />
                </div>
              )}
            </div>
          )}

          <Field label="Trigramme formation">
            <div className="relative">
              <Input
                value={formationFilter}
                onChange={value => {
                  setFormationFilter(value.toUpperCase());
                  setShowFormationSuggestions(true);
                  setSessions([]);
                  setSearchDirectory('');
                  setSelectedCodes(new Set());
                  setProcessResults(new Map());
                }}
                placeholder="ex : TEL, 948, 22B"
              />
              {showFormationSuggestions && formationFilter.trim() && (
                <div className="absolute z-20 mt-1 w-full overflow-hidden rounded-lg border border-slate-200 bg-white shadow-lg">
                  {formationSuggestions.length > 0 ? (
                    formationSuggestions.map(formation => (
                      <button
                        key={formation}
                        type="button"
                        className="block w-full px-3 py-2 text-left text-sm font-semibold text-slate-700 hover:bg-slate-50"
                        onMouseDown={event => {
                          event.preventDefault();
                          setFormationFilter(formation);
                          setShowFormationSuggestions(false);
                          setSessions([]);
                          setSearchDirectory('');
                          setSelectedCodes(new Set());
                          setProcessResults(new Map());
                        }}
                      >
                        {formation}
                      </button>
                    ))
                  ) : (
                    <div className="px-3 py-2 text-sm font-semibold text-red-700">
                      Aucune formation ne commence par "{formationFilter.trim().toUpperCase()}"
                    </div>
                  )}
                </div>
              )}
            </div>
          </Field>

          <div className="flex justify-end">
            <Btn onClick={searchSessions} disabled={!!loading || !formationFilter}>
              {loading === 'search' ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
              Rechercher
            </Btn>
          </div>

          {status === 'warning' && formationFilter.trim() && visibleSessions.length === 0 && (
            <div className="rounded-xl border border-amber-100 bg-amber-50 p-4">
              <div className="mb-3 text-sm font-semibold text-amber-900">
                Aucun CSV trouve.
              </div>
              <input
                id="evalstat-manual-csv"
                type="file"
                accept=".csv,text/csv"
                className="hidden"
                onChange={event => {
                  void uploadManualCsv(event.target.files?.[0]);
                }}
              />
              <div className="flex flex-wrap justify-end gap-3">
                {searchDirectory && (
                  <Btn
                    onClick={() => {
                      void openPath(searchDirectory).catch(error => {
                        setStatus('error');
                        setMessage(error instanceof Error ? error.message : 'Ouverture du dossier impossible.');
                      });
                    }}
                  >
                    Ouvrir le repertoire le plus proche
                  </Btn>
                )}
                <Btn
                  onClick={() => document.getElementById('evalstat-manual-csv')?.click()}
                  disabled={!!loading}
                >
                  {loading === 'manual' ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
                  Choisir un CSV
                </Btn>
                <Btn onClick={searchSessions} disabled={!!loading}>
                  {loading === 'search' ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
                  Relancer la recherche
                </Btn>
              </div>
            </div>
          )}
        </div>
      </Card>

      {visibleSessions.length > 0 && (
        <Card noPad>
          <CardHeader
            title="Cocher les sessions a traiter"
            subtitle={`${sessions.length} session(s), affichage limite a ${visibleSessions.length}`}
            step={2}
          />
          <div className="p-4 space-y-4">
            <div className="overflow-hidden rounded-xl border border-slate-100">
              <table className="w-full">
                <THead cols={['', 'Code IRIS', 'Trigramme RP', 'Date debut ses.', 'Date fin ses.', 'Nb. nommes', 'Statut Session', 'N° Session', 'Etat']} />
                <tbody>
                  <TRow>
                    <TD>
                      <input
                        type="checkbox"
                        checked={allSelected}
                        disabled={selectableRows.length === 0 || loading === 'process'}
                        className="disabled:cursor-not-allowed disabled:opacity-35"
                        onChange={toggleAll}
                        title="Tout cocher / decocher"
                      />
                    </TD>
                    <TD colSpan={8} muted>Tout cocher / decocher</TD>
                  </TRow>
                  {visibleSessions.map(row => {
                    const isProcessing = loading === 'process';
                    const treatment = treatmentBadge(row, processResults.get(row.codeIris), loading, selectedCodes);

                    return (
                      <TRow key={row.codeIris}>
                        <TD>
                          <input
                            type="checkbox"
                            checked={selectedCodes.has(row.codeIris)}
                            disabled={!row.csvExists || isProcessing}
                            className="disabled:cursor-not-allowed disabled:opacity-35"
                            onChange={() => toggleRow(row)}
                          />
                        </TD>
                        <TD mono>{row.codeIris}</TD>
                        <TD>{row.trigrammeRp || '-'}</TD>
                        <TD mono>{formatDate(row.startDate)}</TD>
                        <TD mono>{formatDate(row.endDate)}</TD>
                        <TD mono>{row.nbNommes ?? '-'}</TD>
                        <TD>{row.sessionStatus || '-'}</TD>
                        <TD mono>{row.sessionRef || '-'}</TD>
                        <TD>
                          <Badge variant={treatment.variant}>{treatment.label}</Badge>
                        </TD>
                      </TRow>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <Field label="Apres traitement">
                <select
                  value={openAfterMode}
                  onChange={event => setOpenAfterMode(event.target.value as OpenAfterMode)}
                  disabled={loading === 'process'}
                  className="h-10 rounded-lg border border-slate-200 bg-white px-3 text-sm font-semibold text-slate-700 outline-none transition focus:border-orange-400 disabled:cursor-not-allowed disabled:bg-slate-50 disabled:text-slate-400"
                >
                  <option value="auto">Automatique</option>
                  <option value="excel">Ouvrir le fichier Excel</option>
                  <option value="folder">Ouvrir le repertoire</option>
                  <option value="none">Ne rien ouvrir</option>
                </select>
              </Field>

              <Btn onClick={processSelected} disabled={!!loading || selectedRows.length === 0}>
                {loading === 'process' ? <Loader2 size={14} className="animate-spin" /> : <Play size={14} />}
                Traiter ({selectedRows.length})
              </Btn>
            </div>
          </div>
        </Card>
      )}
    </div>
  );
}
