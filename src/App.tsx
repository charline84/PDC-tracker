import React, { useState } from 'react';
import { 
  Building2, 
  Search, 
  Download, 
  Trash2, 
  Plus, 
  Loader2, 
  FileSpreadsheet, 
  Info,
  ChevronRight,
  MapPin,
  ExternalLink,
  Briefcase,
  Calendar,
  HardHat,
  Map,
  ArrowRight,
  Database,
  RefreshCw,
  Server,
  Github
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import ExcelJS from 'exceljs';

interface CompanyResult {
  nom_complet: string;
  nom_raison_sociale: string;
  siren: string;
  date_creation: string;
  date_fermeture: string;
  activite_principale: string;
  activite_principale_naf25: string;
  dirigeants: any[];
  siege: {
    adresse: string;
    code_postal: string;
    commune: string;
    activite_principale: string;
    etat_administratif: string;
  };
  _search_term: string;
}

interface AdsResult {
  id: string;
  type: string;
  date_depot: string;
  date_obtention: string;
  statut: string;
  demandeur: string;
  siren_demandeur: string;
  adresse: string;
  description: string;
}

export default function App() {
  const [activeModule, setActiveModule] = useState<'entreprises' | 'ads'>('entreprises');
  
  // Entreprises State
  const [companies, setCompanies] = useState<string[]>(['']);
  const [isSearching, setIsSearching] = useState(false);
  const [results, setResults] = useState<CompanyResult[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'search' | 'results'>('search');

  // ADS State
  const [adsSearchType, setAdsSearchType] = useState<'company' | 'permit' | 'geo'>('company');
  const [adsQuery, setAdsQuery] = useState('');
  const [isSearchingAds, setIsSearchingAds] = useState(false);
  const [adsResults, setAdsResults] = useState<AdsResult[]>([]);
  const [adsError, setAdsError] = useState<string | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);

  // --- ENTREPRISES LOGIC ---
  const addCompany = () => {
    setCompanies([...companies, '']);
  };

  const removeCompany = (index: number) => {
    const newCompanies = companies.filter((_, i) => i !== index);
    setCompanies(newCompanies.length ? newCompanies : ['']);
  };

  const updateCompany = (index: number, value: string) => {
    const newCompanies = [...companies];
    newCompanies[index] = value;
    setCompanies(newCompanies);
  };

  const handleSearch = async (e?: React.FormEvent, searchValues?: string[]) => {
    e?.preventDefault();
    const validCompanies = searchValues || companies.filter(c => c.trim() !== '');
    if (validCompanies.length === 0) {
      setError('Veuillez entrer au moins un nom de société.');
      return;
    }

    setCompanies(validCompanies.length ? validCompanies : ['']);
    setActiveModule('entreprises');

    setIsSearching(true);
    setError(null);
    setResults([]);

    try {
      const allResults: CompanyResult[] = [];
      const API_URL = "https://recherche-entreprises.api.gouv.fr/search";

      for (const company of validCompanies) {
        let page = 1;
        let totalPages = 1;

        while (page <= totalPages) {
          if (page > 1 || allResults.length > 0) {
             await new Promise(resolve => setTimeout(resolve, 300));
          }

          const response = await fetch(`${API_URL}?q=${encodeURIComponent(company)}&page=${page}&per_page=25`);
          if (!response.ok) throw new Error("Erreur réseau");
          
          const data = await response.json();
          if (data && data.results) {
            const resultsWithMeta = data.results.map((r: any) => ({
              ...r,
              _search_term: company
            }));
            allResults.push(...resultsWithMeta);
            totalPages = Math.min(data.total_pages, 5); // Limite raisonnable client-side
          } else {
            break;
          }
          page++;
        }
      }

      setResults(allResults);
      setActiveTab('results');
    } catch (err) {
      console.error(err);
      setError('Erreur de connexion à l\'API data.gouv.fr.');
    } finally {
      setIsSearching(false);
    }
  };

  const handleExport = async () => {
    if (results.length === 0) return;

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Résultats");

      const headerStyle: Partial<ExcelJS.Style> = {
        font: { bold: true, color: { argb: "FFFFFF" }, size: 11, name: 'Arial' },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2F5496' } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: {
          top: { style: 'thin', color: { argb: 'D9D9D9' } },
          left: { style: 'thin', color: { argb: 'D9D9D9' } },
          bottom: { style: 'thin', color: { argb: 'D9D9D9' } },
          right: { style: 'thin', color: { argb: 'D9D9D9' } }
        }
      };

      const columns = [
        { header: 'Recherche originale', key: 'search_term', width: 25 },
        { header: 'Nom de la société', key: 'nom', width: 45 },
        { header: 'SIREN', key: 'siren', width: 14 },
        { header: 'Date de création', key: 'date_creation', width: 15 },
        { header: 'Date de clôture', key: 'date_fermeture', width: 15 },
        { header: 'Code NAF/APE', key: 'code_naf', width: 15 },
        { header: 'NAF 2025', key: 'naf_2025', width: 15 },
        { header: 'Dirigeants / Sociétés partenaires', key: 'dirigeants', width: 55 },
        { header: 'Adresse du siège', key: 'adresse', width: 40 },
        { header: 'Code postal', key: 'cp', width: 12 },
        { header: 'Commune', key: 'commune', width: 20 },
        { header: 'État', key: 'etat', width: 10 },
      ];

      worksheet.columns = columns;

      worksheet.getRow(1).eachCell((cell) => {
        cell.style = headerStyle;
      });
      worksheet.getRow(1).height = 35;

      results.forEach((result, index) => {
        const siege = result.siege || {};
        
        const dirigeantsArr = result.dirigeants || [];
        const dirigeantsStr = dirigeantsArr.map((d: any) => {
          if (d.denomination) return d.denomination;
          const nomComplet = [d.prenoms, d.nom].filter(Boolean).join(" ");
          return d.qualite ? `${nomComplet} (${d.qualite})` : nomComplet;
        }).filter(Boolean).join(", ");

        const adresseStr = [siege.adresse, siege.code_postal, siege.commune].filter(Boolean).join(", ");
        const etat = siege.etat_administratif === 'A' ? 'Active' : (siege.etat_administratif === 'F' ? 'Fermée' : siege.etat_administratif);

        const row = worksheet.addRow({
          search_term: result._search_term || "",
          nom: result.nom_complet || result.nom_raison_sociale || "",
          siren: result.siren || "",
          date_creation: result.date_creation || "",
          date_fermeture: result.date_fermeture || "",
          code_naf: result.activite_principale || "",
          naf_2025: result.activite_principale_naf25 || "",
          dirigeants: dirigeantsStr,
          adresse: adresseStr,
          cp: siege.code_postal || "",
          commune: siege.commune || "",
          etat: etat
        });

        if (index % 2 === 1) {
          row.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F6FC' } };
          });
        }
        
        row.eachCell((cell) => {
          cell.alignment = { vertical: 'top', wrapText: true };
          cell.font = { name: 'Arial', size: 10 };
          cell.border = {
            top: { style: 'thin', color: { argb: 'D9D9D9' } },
            left: { style: 'thin', color: { argb: 'D9D9D9' } },
            bottom: { style: 'thin', color: { argb: 'D9D9D9' } },
            right: { style: 'thin', color: { argb: 'D9D9D9' } }
          };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `investissements_${new Date().toISOString().split('T')[0]}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      alert('Erreur lors de l\'exportation Excel.');
    }
  };

  // --- ADS LOGIC (Supabase Integration Placeholder) ---
  const handleAdsSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!adsQuery.trim()) {
      setAdsError('Veuillez entrer une valeur pour la recherche.');
      return;
    }

    setIsSearchingAds(true);
    setAdsError(null);
    setAdsResults([]);

    try {
      // Pour l'instant, on simule l'appel Supabase jusqu'à sa configuration
      // Dans le futur: await supabase.from('permis_construire').select('*').eq(...)
      await new Promise(resolve => setTimeout(resolve, 800));
      
      setAdsError('Supabase non configuré. Veuillez suivre les instructions en bas de page pour connecter votre base de données contenant les données sitadel.');

    } catch (err) {
      setAdsError('Erreur de connexion au serveur ADS.');
    } finally {
      setIsSearchingAds(false);
    }
  };

  const syncDatabase = async () => {
    setIsSyncing(true);
    // Simulation of a cron/sync job
    setTimeout(() => {
      setIsSyncing(false);
      alert("La synchronisation des données depuis data.gouv.fr se fait généralement via un script backend (Github Actions) car le fichier est trop lourd pour le navigateur. (Voir les instructions)");
    }, 1500);
  };

  const formatDate = (dateStr: string) => {
    if (!dateStr) return 'N/A';
    try {
      const d = new Date(dateStr);
      return d.toLocaleDateString('fr-FR');
    } catch {
      return dateStr;
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-20">
        <div className="max-w-5xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2">
              <div className="bg-blue-600 p-2 rounded-lg">
                <Building2 className="text-white w-5 h-5" />
              </div>
              <h1 className="text-xl font-bold tracking-tight text-slate-800 hidden sm:block">Entreprise Tracker</h1>
            </div>

            <nav className="hidden md:flex gap-1">
              <button
                onClick={() => setActiveModule('entreprises')}
                className={`px-4 py-2 font-medium text-sm rounded-lg transition-colors ${
                  activeModule === 'entreprises' ? 'bg-blue-50 text-blue-700' : 'text-slate-600 hover:bg-slate-50'
                }`}
              >
                Entreprises & Participations
              </button>
              <button
                onClick={() => setActiveModule('ads')}
                className={`px-4 py-2 font-medium text-sm rounded-lg transition-colors ${
                  activeModule === 'ads' ? 'bg-amber-50 text-amber-700' : 'text-slate-600 hover:bg-slate-50'
                }`}
              >
                Permis de Construire (OpenData)
              </button>
            </nav>
          </div>
          
          {activeModule === 'entreprises' && (
            <nav className="flex gap-1 p-1 bg-slate-100 rounded-lg">
              <button 
                onClick={() => setActiveTab('search')}
                className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all ${
                  activeTab === 'search' 
                  ? 'bg-white text-blue-600 shadow-sm' 
                  : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Recherche
              </button>
              <button 
                onClick={() => results.length > 0 && setActiveTab('results')}
                disabled={results.length === 0}
                className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all ${
                  activeTab === 'results' 
                  ? 'bg-white text-blue-600 shadow-sm' 
                  : results.length > 0 ? 'text-slate-600 hover:text-slate-900' : 'text-slate-300 cursor-not-allowed'
                }`}
              >
                Résultats ({results.length})
              </button>
            </nav>
          )}
        </div>
        <div className="md:hidden flex border-t border-slate-100">
             <button
                onClick={() => setActiveModule('entreprises')}
                className={`flex-1 py-3 font-medium text-sm transition-colors text-center ${
                  activeModule === 'entreprises' ? 'border-b-2 border-blue-600 text-blue-700' : 'text-slate-600 bg-slate-50'
                }`}
              >
                Entreprises
              </button>
              <button
                onClick={() => setActiveModule('ads')}
                className={`flex-1 py-3 font-medium text-sm transition-colors text-center ${
                  activeModule === 'ads' ? 'border-b-2 border-amber-600 text-amber-700' : 'text-slate-600 bg-slate-50'
                }`}
              >
                Permis
              </button>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 py-8">
        <AnimatePresence mode="wait">
          {activeModule === 'entreprises' && (
            <motion.div
              key="entreprises-module"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
            >
              <AnimatePresence mode="wait">
                {activeTab === 'search' ? (
                  <motion.div 
                    key="search-view"
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    className="max-w-2xl mx-auto"
                  >
                    <div className="mb-8 text-center">
                      <h2 className="text-3xl font-bold text-slate-900 mb-2">Analyse des participations</h2>
                      <p className="text-slate-600">Entrez les noms des sociétés pour identifier leurs investissements via data.gouv.fr. Architecture Client-Side 100% compatible avec GitHub Pages.</p>
                    </div>

                    <form onSubmit={(e) => handleSearch(e)} className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 overflow-hidden">
                      <div className="p-6 space-y-4">
                        <div className="space-y-3">
                          {companies.map((company, index) => (
                            <div key={index} className="flex gap-2">
                              <div className="relative flex-1">
                                <div className="absolute inset-y-0 left-3 flex items-center pointer-events-none">
                                  <Search className="w-4 h-4 text-slate-400" />
                                </div>
                                <input
                                  type="text"
                                  value={company}
                                  onChange={(e) => updateCompany(index, e.target.value)}
                                  placeholder="Nom de la société, SIREN (ex: TotalEnergies)"
                                  className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all placeholder:text-slate-400"
                                  autoFocus={index === companies.length - 1}
                                />
                              </div>
                              {companies.length > 1 && (
                                <button
                                  type="button"
                                  onClick={() => removeCompany(index)}
                                  className="p-3 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-xl transition-colors"
                                >
                                  <Trash2 className="w-5 h-5" />
                                </button>
                              )}
                            </div>
                          ))}
                        </div>

                        <button
                          type="button"
                          onClick={addCompany}
                          className="w-full flex items-center justify-center gap-2 py-2 border-2 border-dashed border-slate-200 rounded-xl text-slate-500 hover:text-blue-600 hover:border-blue-200 hover:bg-blue-50 transition-all text-sm font-medium"
                        >
                          <Plus className="w-4 h-4" />
                          Ajouter une autre société
                        </button>
                      </div>

                      {error && (
                        <div className="px-6 pb-4">
                          <div className="bg-red-50 border border-red-100 text-red-600 px-4 py-3 rounded-xl text-sm flex items-start gap-2">
                            <Info className="w-4 h-4 mt-0.5" />
                            <span>{error}</span>
                          </div>
                        </div>
                      )}

                      <div className="bg-slate-50 px-6 py-4 flex flex-col sm:flex-row sm:items-center justify-between border-t border-slate-200 gap-4">
                        <div className="text-xs text-slate-500">
                          Mode <span className="font-semibold text-blue-600">Client-Side (GH Pages Ready)</span> | Propulsé par <span className="font-semibold">api.gouv.fr</span>
                        </div>
                        <button
                          type="submit"
                          disabled={isSearching}
                          className="flex items-center justify-center gap-2 px-8 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white rounded-xl font-bold shadow-lg shadow-blue-600/20 transition-all cursor-pointer"
                        >
                          {isSearching ? (
                            <>
                              <Loader2 className="w-5 h-5 animate-spin" />
                              Recherche...
                            </>
                          ) : (
                            <>
                              Lancer l'analyse
                              <ChevronRight className="w-5 h-5" />
                            </>
                          )}
                        </button>
                      </div>
                    </form>
                  </motion.div>
                ) : (
                  <motion.div 
                    key="results-view"
                    initial={{ opacity: 0, x: 20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, x: -20 }}
                    className="space-y-6"
                  >
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div>
                        <h2 className="text-2xl font-bold text-slate-900">Résultats de la recherche</h2>
                        <p className="text-slate-500">{results.length} entreprises identifiées</p>
                      </div>
                      <button
                        onClick={handleExport}
                        className="flex items-center justify-center gap-2 px-6 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold shadow-lg shadow-emerald-600/20 transition-all cursor-pointer"
                      >
                        <Download className="w-5 h-5" />
                        Télécharger Excel
                      </button>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {results.map((result, idx) => (
                        <a
                          key={`${result.siren}-${idx}`}
                          href={`https://www.pappers.fr/entreprise/${result.siren}`}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="block h-full group"
                        >
                          <motion.div 
                            initial={{ opacity: 0, y: 10 }}
                            animate={{ opacity: 1, y: 0 }}
                            transition={{ delay: Math.min(idx * 0.05, 1) }}
                            className="bg-white p-5 rounded-2xl border border-slate-200 hover:border-blue-400 hover:shadow-lg transition-all h-full flex flex-col relative overflow-hidden"
                          >
                            <div className="absolute top-0 right-0 w-16 h-16 bg-blue-50 rounded-bl-full -z-0 opacity-0 group-hover:opacity-100 transition-opacity" />
                            
                            <div className="flex justify-between items-start mb-4 relative z-10">
                              <div>
                                <h3 className="font-bold text-slate-900 group-hover:text-blue-700 transition-colors">
                                  {result.nom_complet || result.nom_raison_sociale}
                                </h3>
                                <div className="flex flex-wrap items-center gap-2 mt-2">
                                  <span className="text-[10px] font-mono text-slate-500 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">SIREN: {result.siren}</span>
                                  <span className={`text-[10px] uppercase font-bold px-2 py-0.5 rounded border ${
                                    result.siege?.etat_administratif === 'A' || !result.date_fermeture ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-rose-50 text-rose-700 border-rose-200'
                                  }`}>
                                    {result.siege?.etat_administratif === 'A' || !result.date_fermeture ? 'Active' : 'Fermée'}
                                  </span>
                                  {result.date_creation && (
                                    <span className="text-[10px] text-slate-500 bg-slate-50 px-2 py-0.5 rounded border border-slate-100 inline-flex items-center gap-1">
                                      <Calendar className="w-3 h-3" /> {formatDate(result.date_creation)}
                                    </span>
                                  )}
                                </div>
                              </div>
                              <div className="bg-slate-50 p-2 rounded-lg text-slate-400 group-hover:text-blue-600 group-hover:bg-blue-100 transition-all shadow-sm">
                                <ExternalLink className="w-4 h-4" />
                              </div>
                            </div>

                            <div className="space-y-3 flex-1 relative z-10">
                              <div className="flex gap-3 text-sm flex-col">
                                <div className="flex gap-1.5 items-start">
                                  <Briefcase className="w-4 h-4 text-slate-400 flex-shrink-0 mt-0.5" />
                                  <div className="text-slate-600 flex flex-wrap gap-x-2 gap-y-1 items-center">
                                    <span className="font-semibold text-slate-700">{result.activite_principale}</span>
                                    {result.activite_principale_naf25 && (
                                      <span className="text-xs bg-slate-100 text-slate-500 px-1.5 rounded">NAF 2025: {result.activite_principale_naf25}</span>
                                    )}
                                  </div>
                                </div>
                              </div>
                            </div>
                            
                            <div className="mt-4 pt-4 border-t border-slate-50 flex items-center justify-between text-[10px] text-slate-400 relative z-10">
                              <span>Source: "{result._search_term}"</span>
                              <span className="flex items-center gap-1 group-hover:text-blue-600">Voir sur Pappers <ArrowRight className="w-3 h-3" /></span>
                            </div>
                          </motion.div>
                        </a>
                      ))}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </motion.div>
          )}

          {activeModule === 'ads' && (
            <motion.div
              key="ads-module"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <div className="mb-6 text-center max-w-2xl mx-auto">
                <div className="inline-flex items-center justify-center p-3 bg-amber-100 rounded-full mb-4">
                  <Database className="w-8 h-8 text-amber-600" />
                </div>
                <h2 className="text-3xl font-bold text-slate-900 mb-2">Permis de Construire (OpentData)</h2>
                <p className="text-slate-600 mb-4">Module connecté à la base de données gouvernementale (Sitadel). Les données étant volumineuses, elles nécessitent une base de données de relais (ex: Supabase).</p>
                
                {/* Database Sync Panel */}
                <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex flex-col md:flex-row items-center justify-between gap-4 text-left">
                  <div>
                     <h3 className="font-bold text-amber-900 text-sm">Gestion de la base de données (Supabase)</h3>
                     <p className="text-xs text-amber-700 mt-1">Dernière mise à jour : <strong>Vérification...</strong></p>
                  </div>
                  <button 
                    onClick={syncDatabase}
                    disabled={isSyncing}
                    className="px-4 py-2 bg-white border border-amber-300 text-amber-700 rounded-lg text-sm font-bold shadow-sm hover:bg-amber-100 transition-colors flex items-center gap-2"
                  >
                    <RefreshCw className={`w-4 h-4 ${isSyncing ? 'animate-spin' : ''}`} />
                    {isSyncing ? 'Synchronisation...' : 'Télécharger depuis data.gouv.fr'}
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 overflow-hidden max-w-3xl mx-auto">
                <div className="flex border-b border-slate-100">
                  <button 
                    onClick={() => { setAdsSearchType('company'); setAdsQuery(''); }}
                    className={`flex-1 py-4 font-medium text-sm transition-all flex flex-col items-center gap-1 ${adsSearchType === 'company' ? 'bg-amber-50 text-amber-700 border-b-2 border-amber-600' : 'text-slate-500 hover:bg-slate-50'}`}
                  >
                    <Building2 className="w-5 h-5" />
                    Par Société / Porteur
                  </button>
                  <button 
                    onClick={() => { setAdsSearchType('permit'); setAdsQuery(''); }}
                    className={`flex-1 py-4 font-medium text-sm transition-all flex flex-col items-center gap-1 border-l border-slate-100 ${adsSearchType === 'permit' ? 'bg-amber-50 text-amber-700 border-b-2 border-amber-600' : 'text-slate-500 hover:bg-slate-50'}`}
                  >
                    <HardHat className="w-5 h-5" />
                    Numéro de Permis
                  </button>
                  <button 
                    onClick={() => { setAdsSearchType('geo'); setAdsQuery(''); }}
                    className={`flex-1 py-4 font-medium text-sm transition-all flex flex-col items-center gap-1 border-l border-slate-100 ${adsSearchType === 'geo' ? 'bg-amber-50 text-amber-700 border-b-2 border-amber-600' : 'text-slate-500 hover:bg-slate-50'}`}
                  >
                    <Map className="w-5 h-5" />
                    Zone Géographique
                  </button>
                </div>
                
                <form onSubmit={handleAdsSearch} className="p-6">
                  <div className="relative mb-4">
                    <div className="absolute inset-y-0 left-4 flex items-center pointer-events-none">
                      <Search className="w-5 h-5 text-slate-400" />
                    </div>
                    <input
                      type="text"
                      className="w-full pl-12 pr-4 py-4 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 outline-none transition-all placeholder:text-slate-400 font-medium"
                      placeholder="Attente de connexion à Supabase..."
                      value={adsQuery}
                      onChange={(e) => setAdsQuery(e.target.value)}
                    />
                  </div>

                  <div className="flex items-center justify-between">
                    <p className="text-xs text-amber-600 flex items-center gap-1 font-medium bg-amber-50 px-2 py-1 rounded">
                      <Server className="w-3.5 h-3.5" />
                      Client Supabase non connecté
                    </p>
                    <button
                      type="submit"
                      disabled={isSearchingAds}
                      className="px-6 py-2.5 bg-amber-600 hover:bg-amber-700 disabled:bg-amber-400 text-white rounded-xl font-bold shadow-md shadow-amber-600/20 transition-all cursor-pointer flex items-center gap-2"
                    >
                      {isSearchingAds ? <><Loader2 className="w-4 h-4 animate-spin"/> Recherche...</> : 'Rechercher'}
                    </button>
                  </div>
                </form>
              </div>

              {adsError && (
                <div className="max-w-3xl mx-auto">
                  <div className="bg-slate-800 text-slate-200 px-6 py-6 rounded-xl text-sm flex flex-col gap-4 shadow-xl border border-slate-700">
                    <div className="flex items-center gap-2 text-amber-400 font-bold text-lg">
                      <Database className="w-5 h-5" />
                      Configuration Requise
                    </div>
                    <p className="text-slate-300">
                      Le fichier des autorisations d'urbanisme complet pèse plusieurs gigaoctets. L'architecture optimale pour héberger votre application gratuitement sur <strong>GitHub Pages</strong> consiste à utiliser <strong>Supabase</strong> comme base de données.
                    </p>
                  </div>
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>
      
      <footer className="max-w-5xl mx-auto px-4 py-8 text-center border-t border-slate-200 mt-12 bg-white flex flex-col sm:flex-row items-center justify-between">
        <p className="text-slate-500 text-sm">&copy; {new Date().getFullYear()} Entreprise Tracker.</p>
        <div className="flex gap-4 mt-4 sm:mt-0 text-xs text-slate-400 font-medium items-center">
          <Github className="w-4 h-4" />
          <span>Architecture Serverless</span>
          <span>•</span>
          <span>data.gouv.fr (Sitadel/Recherche)</span>
        </div>
      </footer>
    </div>
  );
}

