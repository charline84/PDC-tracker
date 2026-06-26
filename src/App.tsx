import React, { useState, useMemo } from 'react';
import NAF_CODES from '@socialgouv/codes-naf';
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
  Github,
  LogOut
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import ExcelJS from 'exceljs';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, BarChart, Bar, XAxis, YAxis, CartesianGrid } from 'recharts';
import MapComponent from './components/MapComponent';
import { auth, db } from './lib/firebase';
import { GoogleAuthProvider, signInWithPopup, onAuthStateChanged, signOut, sendSignInLinkToEmail, isSignInWithEmailLink, signInWithEmailLink } from 'firebase/auth';
import { doc, setDoc, getDoc, updateDoc, collection, getDocs, query, where } from 'firebase/firestore';

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

export interface AdsResult {
  id: string;
  type: string;
  date_depot: string;
  date_obtention: string;
  statut: string;
  demandeur: string;
  siren_demandeur: string;
  adresse: string;
  description: string;
  surface?: number;
  logements?: number;
}

const getNafLabel = (codeStr: string) => {
  if (!codeStr) return '';
  const codePart = codeStr.trim().split(' ')[0];
  const found = (NAF_CODES as any[]).find((n) => n.id === codePart);
  if (found) {
    return `${codePart} - ${found.label}`;
  }
  return codeStr;
};

export default function App() {
  const [activeModule, setActiveModule] = useState<'entreprises' | 'ads' | 'admin'>('entreprises');
  
  // Entreprises State
  const [companies, setCompanies] = useState<string[]>(['']);
  const [geoFilter, setGeoFilter] = useState('');
  const [dateFilter, setDateFilter] = useState<'all' | '6_months' | '1_year' | '2_years' | '5_years' | 'more_5_years'>('all');
  const [isSearching, setIsSearching] = useState(false);
  const [results, setResults] = useState<CompanyResult[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'search' | 'results'>('search');

  // ADS State
  const [adsSearchType, setAdsSearchType] = useState<'company' | 'permit' | 'geo'>('geo');
  const [adsQuery, setAdsQuery] = useState('');
  const [isSearchingAds, setIsSearchingAds] = useState(false);
  const [adsResults, setAdsResults] = useState<AdsResult[]>([]);
  const [adsError, setAdsError] = useState<string | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncProgress, setSyncProgress] = useState<string | null>(null);
  const [isMockData, setIsMockData] = useState(false);

  const [adsFilterLogements, setAdsFilterLogements] = useState('all');
  const [adsFilterDate, setAdsFilterDate] = useState('all');
  const [adsFilterStatus, setAdsFilterStatus] = useState('all');
  const [adsViewMode, setAdsViewMode] = useState<'list' | 'map'>('list');

  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [authChecking, setAuthChecking] = useState(true);
  const [userData, setUserData] = useState<any>(null);
  const [notApproved, setNotApproved] = useState(false);
  const [authError, setAuthError] = useState('');
  
  // Admin state
  const [unapprovedUsers, setUnapprovedUsers] = useState<any[]>([]);
  const [approvedUsers, setApprovedUsers] = useState<any[]>([]);
  const [isAdminLoading, setIsAdminLoading] = useState(false);

  // Email Link state
  const [loginEmail, setLoginEmail] = useState('');
  const [emailLinkSent, setEmailLinkSent] = useState(false);
  const [isVerifyingLink, setIsVerifyingLink] = useState(false);

  React.useEffect(() => {
    const checkEmailLink = async () => {
      if (isSignInWithEmailLink(auth, window.location.href)) {
        setIsVerifyingLink(true);
        let email = window.localStorage.getItem('emailForSignIn');
        if (!email) {
          email = window.prompt('Veuillez fournir votre e-mail pour confirmation');
        }
        if (email) {
          try {
            await signInWithEmailLink(auth, email, window.location.href);
            window.localStorage.removeItem('emailForSignIn');
            // Remove the token from URL to avoid issues on reload
            window.history.replaceState(null, '', window.location.pathname);
          } catch (err: any) {
            console.error("Erreur lors de la connexion par lien", err);
            setAuthError(err.message || 'Le lien est invalide ou a expiré.');
          }
        }
        setIsVerifyingLink(false);
      }
    };
    checkEmailLink();
  }, []);

  React.useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        try {
          const userDocRef = doc(db, 'users', user.uid);
          const userDocSnap = await getDoc(userDocRef);
          
          if (userDocSnap.exists()) {
            const data = userDocSnap.data();
            if (!data.isApproved) {
              setIsAuthenticated(false);
              setNotApproved(true);
              setUserData(null);
            } else {
              setIsAuthenticated(true);
              setNotApproved(false);
              setUserData(data);
            }
          } else {
            // New user who hasn't been saved to DB yet or something strange
            const isCharline = user.email?.toLowerCase() === 'charline.haim@gmail.com';
            const newUserData = {
              email: user.email?.toLowerCase() || '',
              role: isCharline ? 'admin' : 'user',
              isApproved: isCharline ? true : false,
              createdAt: new Date().toISOString()
            };
            await setDoc(userDocRef, newUserData);
            
            if (!newUserData.isApproved) {
              setIsAuthenticated(false);
              setNotApproved(true);
              setUserData(null);
            } else {
              setIsAuthenticated(true);
              setNotApproved(false);
              setUserData(newUserData);
            }
          }
        } catch (err: any) {
          console.error("Error fetching user role", err);
          if (err.message && err.message.includes('Missing or insufficient permissions')) {
            setAuthError("Les règles de sécurité de votre base de données Firebase ne sont pas configurées. Allez dans Firebase Console > Firestore Database > Règles et copiez-y le contenu du fichier firestore.rules.");
          }
          setIsAuthenticated(false);
          setNotApproved(false);
          setUserData(null);
        }
      } else {
        setIsAuthenticated(false);
        setNotApproved(false);
        setUserData(null);
      }
      setAuthChecking(false);
    });
    return unsubscribe;
  }, []);

  const fetchAdminUsers = async () => {
    if (userData?.role !== 'admin') return;
    setIsAdminLoading(true);
    try {
      const qUnapproved = query(collection(db, 'users'), where('isApproved', '==', false));
      const snapshotUnapproved = await getDocs(qUnapproved);
      const unapproved: any[] = [];
      snapshotUnapproved.forEach((doc) => {
        unapproved.push({ id: doc.id, ...doc.data() });
      });
      setUnapprovedUsers(unapproved);

      const qApproved = query(collection(db, 'users'), where('isApproved', '==', true));
      const snapshotApproved = await getDocs(qApproved);
      const approved: any[] = [];
      snapshotApproved.forEach((doc) => {
        // Exclude current user from the list if desired, but we can just show everyone.
        if (doc.id !== auth.currentUser?.uid) {
          approved.push({ id: doc.id, ...doc.data() });
        }
      });
      setApprovedUsers(approved);
    } catch (err) {
      console.error("Error fetching users", err);
    } finally {
      setIsAdminLoading(false);
    }
  };

  const approveUser = async (userId: string) => {
    try {
      await updateDoc(doc(db, 'users', userId), {
        isApproved: true
      });
      fetchAdminUsers();
    } catch (err) {
      console.error("Error approving user", err);
    }
  };

  const revokeUser = async (userId: string) => {
    try {
      await updateDoc(doc(db, 'users', userId), {
        isApproved: false
      });
      fetchAdminUsers();
    } catch (err) {
      console.error("Error revoking user", err);
    }
  };

  React.useEffect(() => {
    if (activeModule === 'admin') {
      fetchAdminUsers();
    }
  }, [activeModule, userData]);

  const [geoUtilQuery, setGeoUtilQuery] = useState('');
  const [geoUtilResults, setGeoUtilResults] = useState<any[]>([]);
  const [isSearchingGeoUtil, setIsSearchingGeoUtil] = useState(false);

  const handleGeoUtilSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!geoUtilQuery.trim()) return;
    setIsSearchingGeoUtil(true);
    try {
      const isPostalCode = /^\d+$/.test(geoUtilQuery.trim());
      const url = isPostalCode 
        ? `https://geo.api.gouv.fr/communes?codePostal=${geoUtilQuery.trim()}&fields=nom,code,codesPostaux,population`
        : `https://geo.api.gouv.fr/communes?nom=${encodeURIComponent(geoUtilQuery.trim())}&fields=nom,code,codesPostaux,population&boost=population&limit=5`;
        
      const res = await fetch(url);
      const data = await res.json();
      setGeoUtilResults(data);
    } catch (err) {
      console.error(err);
    } finally {
      setIsSearchingGeoUtil(false);
    }
  };

  const filteredAdsResults = useMemo(() => {
    return adsResults.filter(r => {
      // Logements
      const nbLogements = r.logements || 0;
      if (adsFilterLogements === 'lt_15' && nbLogements >= 15) return false;
      if (adsFilterLogements === '15_50' && (nbLogements < 15 || nbLogements > 50)) return false;
      if (adsFilterLogements === 'gt_50' && nbLogements <= 50) return false;

      // Date
      if (adsFilterDate !== 'all') {
        const d = r.date_depot || r.date_obtention;
        if (!d) return false;
        
        const datePermis = new Date(d);
        const now = new Date();
        const monthDiff = (now.getFullYear() - datePermis.getFullYear()) * 12 + now.getMonth() - datePermis.getMonth();

        if (adsFilterDate === '3_months' && monthDiff > 3) return false;
        if (adsFilterDate === '6_months' && monthDiff > 6) return false;
        if (adsFilterDate === '1_year' && monthDiff > 12) return false;
        if (adsFilterDate === '3_years' && monthDiff > 36) return false;
      }

      // Status
      if (adsFilterStatus !== 'all') {
         const statut = r.statut.toLowerCase();
         if (adsFilterStatus === 'autorise' && !(statut.includes('autoris') || statut.includes('accord'))) return false;
         if (adsFilterStatus === 'en_cours' && !statut.includes('cours')) return false;
         if (adsFilterStatus === 'ouverture' && !statut.includes('ouverture')) return false;
         if (adsFilterStatus === 'refuse' && !statut.includes('refus')) return false;
         if (adsFilterStatus === 'depot' && !statut.includes('dépôt') && !statut.includes('depot')) return false;
      }

      return true;
    });
  }, [adsResults, adsFilterLogements, adsFilterDate, adsFilterStatus]);

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
    if (validCompanies.length === 0 && !geoFilter.trim()) {
      setError('Veuillez entrer au moins un nom de société ou une zone géographique.');
      return;
    }

    setCompanies(validCompanies.length ? validCompanies : ['']);
    setActiveModule('entreprises');

    setIsSearching(true);
    setError(null);
    setResults([]);

    try {
      let filteredResults: CompanyResult[] = [];
      const API_URL = "https://recherche-entreprises.api.gouv.fr/search";
      
      const geoParam = geoFilter.trim() 
        ? (geoFilter.trim().length <= 3 ? `departement=${geoFilter.trim()}` : `code_postal=${geoFilter.trim()}`) 
        : '';

      const itemsToIterate = validCompanies.length > 0 ? validCompanies : [''];

      for (const company of itemsToIterate) {
        let page = 1;
        let totalPages = 1;

        while (page <= totalPages) {
          if (page > 1 || filteredResults.length > 0) {
             await new Promise(resolve => setTimeout(resolve, 300));
          }

          const queryParam = company ? `q=${encodeURIComponent(company)}` : '';
          const params = [queryParam, geoParam, `page=${page}`, `per_page=25`].filter(Boolean).join('&');

          const response = await fetch(`${API_URL}?${params}`);
          if (!response.ok) throw new Error("Erreur réseau");
          
          const data = await response.json();
          if (data && data.results) {
            let resultsWithMeta = data.results.map((r: any) => ({
              ...r,
              _search_term: company || (geoFilter.trim() ? `Zone: ${geoFilter.trim()}` : '')
            }));
            
            // Client side date filtering
            if (dateFilter !== 'all') {
              const now = new Date();
              let msFilter = 0;
              let isMore = false;
              if (dateFilter === '6_months') msFilter = 6 * 30 * 24 * 60 * 60 * 1000;
              else if (dateFilter === '1_year') msFilter = 365 * 24 * 60 * 60 * 1000;
              else if (dateFilter === '2_years') msFilter = 2 * 365 * 24 * 60 * 60 * 1000;
              else if (dateFilter === '5_years') msFilter = 5 * 365 * 24 * 60 * 60 * 1000;
              else if (dateFilter === 'more_5_years') {
                msFilter = 5 * 365 * 24 * 60 * 60 * 1000;
                isMore = true;
              }
              
              resultsWithMeta = resultsWithMeta.filter((r: any) => {
                if (!r.date_creation) return false;
                const creationDate = new Date(r.date_creation);
                const diff = now.getTime() - creationDate.getTime();
                return isMore ? diff > msFilter : diff <= msFilter;
              });
            }

            filteredResults.push(...resultsWithMeta);
            totalPages = Math.min(data.total_pages, 5); // Limite raisonnable client-side
          } else {
            break;
          }
          page++;
        }
      }

      setResults(filteredResults);
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
          code_naf: getNafLabel(result.activite_principale) || "",
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

  // --- ADS LOGIC (API Fetchs) ---
  const generateMockPermits = (query: string, count: number = 42): AdsResult[] => {
    return Array.from({ length: count }).map((_, i) => ({
      id: `PC-${2023 + Math.floor(Math.random() * 2)}-${Math.floor(Math.random() * 900000)}-${i}`,
      type: i % 3 === 0 ? 'Aménagement' : (i % 2 === 0 ? 'Logement' : 'Local Commercial'),
      date_depot: `2023-${String(Math.floor(Math.random() * 12) + 1).padStart(2, '0')}-${String(Math.floor(Math.random() * 28) + 1).padStart(2, '0')}`,
      date_obtention: Math.random() > 0.5 ? `2024-${String(Math.floor(Math.random() * 12) + 1).padStart(2, '0')}-${String(Math.floor(Math.random() * 28) + 1).padStart(2, '0')}` : '',
      statut: Math.random() > 0.3 ? 'Autorisé' : 'En cours d\'instruction',
      demandeur: ['SCI Horizon', 'NEXITY', 'BOUYGUES IMMOBILIER', 'Foncière Commune', 'Particulier'][Math.floor(Math.random() * 5)],
      siren_demandeur: ['499311221', '330112344', '551122340', '', ''][Math.floor(Math.random() * 5)],
      adresse: `${Math.floor(Math.random() * 100) + 1} Avenue de la République, ${query || '75010 Paris'}`,
      description: `Construction d'une surface de ${Math.floor(Math.random() * 5000) + 100}m² pour un projet de type ${i % 3 === 0 ? 'Aménagement' : 'Logement'}.`,
      surface: Math.floor(Math.random() * 5000) + 100
    })).sort((a, b) => new Date(b.date_depot).getTime() - new Date(a.date_depot).getTime());
  };

  const triggerAdsSearch = async (query: string, type: 'company' | 'permit' | 'geo') => {
    setIsSearchingAds(true);
    setAdsError(null);
    setAdsResults([]);
    setIsMockData(false);

    try {
      let baseUrl = 'https://tabular-api.data.gouv.fr/api/resources/65a9e264-7a20-46a9-9d98-66becb817bc3/data/?page_size=200';
      if (type === 'company' && query.trim()) {
         if (/^\d{9}$/.test(query.trim())) {
           baseUrl += `&SIREN_DEM__exact=${query.trim()}`;
         } else {
           baseUrl += `&DENOM_DEM__contains=${encodeURIComponent(query.trim())}`;
         }
      } else if (type === 'permit' && query.trim()) {
         baseUrl += `&NUM_DAU__contains=${encodeURIComponent(query.trim())}`;
      } else if (type === 'geo' && query.trim()) {
         if (/^\d+$/.test(query.trim())) {
           baseUrl += `&ADR_CODPOST_TER__contains=${encodeURIComponent(query.trim())}`;
         } else {
           baseUrl += `&ADR_LOCALITE_TER__contains=${encodeURIComponent(query.trim().toUpperCase())}`;
         }
      }
      
      let allDataRaw: any[] = [];
      let page = 1;
      const maxPages = 20; // 4000 résultats max client-side
      
      while (page <= maxPages) {
        const url = `${baseUrl}&page=${page}`;
        const response = await fetch(url);
        const resData = await response.json();
        
        if (!response.ok) {
           if (page === 1) throw new Error(resData.error || resData.errors?.[0]?.detail || 'Erreur API');
           break;
        }

        if (!resData.data || resData.data.length === 0) break;
        
        allDataRaw = allDataRaw.concat(resData.data);
        
        const total = resData.meta?.total || 0;
        if (allDataRaw.length >= total) break;
        
        page++;
      }

      const allData = allDataRaw.map((d: any) => {
        let statut = 'En cours / Projet';
        if (d.ETAT_DAU === 5 || d.ETAT_DAU === 8) statut = 'Autorisé';
        if (d.ETAT_DAU === 6) statut = 'Ouverture de chantier';
        if (d.ETAT_DAU === 9) statut = 'Refusé';
        if (d.ETAT_DAU === 2) statut = 'Dépôt';
        
        let typeStr = 'Autre';
        if (d.NATURE_PROJET_COMPLETEE === 1) typeStr = 'Nouvelle construction';
        else if (d.NATURE_PROJET_COMPLETEE === 2) typeStr = 'Extension';
        
        return {
          id: d.NUM_DAU || d.__id?.toString(),
          type: typeStr,
          date_depot: d.DR_DEPOT || `${d.AN_DEPOT || ''}`,
          date_obtention: d.DATE_REELLE_AUTORISATION || '',
          statut: statut,
          demandeur: d.DENOM_DEM || 'Non renseigné',
          siren_demandeur: d.SIREN_DEM || '',
          adresse: [d.ADR_NUM_TER, d.ADR_LIBVOIE_TER, d.ADR_LIEUDIT_TER, d.ADR_CODPOST_TER, d.ADR_LOCALITE_TER].filter(Boolean).join(' '),
          description: `Surface Habitable: ${d.SURF_HAB_CREEE || 0}m², Logements créés: ${d.NB_LGT_TOT_CREES || 0}`,
          surface: d.SURF_HAB_CREEE || 0,
          logements: d.NB_LGT_TOT_CREES || 0
        };
      });

      setAdsResults(allData);
    } catch (err: any) {
      console.error(err);
      setAdsError('Erreur lors de la récupération des données réelles: ' + err.message);
    } finally {
      setIsSearchingAds(false);
    }
  };

  const handleAdsSearch = (e: React.FormEvent) => {
    e.preventDefault();
    triggerAdsSearch(adsQuery, adsSearchType);
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

  const handleEmailLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!loginEmail) return;
    setAuthError('');
    try {
      const actionCodeSettings = {
        url: window.location.origin,
        handleCodeInApp: true,
      };
      await sendSignInLinkToEmail(auth, loginEmail, actionCodeSettings);
      window.localStorage.setItem('emailForSignIn', loginEmail);
      setEmailLinkSent(true);
    } catch (err: any) {
      console.error(err);
      setAuthError(err.message || 'Erreur lors de l\'envoi du lien par email');
    }
  };

  const handleLogin = async () => {
    setAuthError('');
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
    } catch (err: any) {
      console.error(err);
      setAuthError(err.message || 'Erreur lors de la connexion Google');
    }
  };

  const handleLogout = () => {
    signOut(auth);
  };

  if (authChecking) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <div className="flex flex-col items-center">
          <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-4" />
          <p className="text-slate-500 font-medium">Chargement...</p>
        </div>
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-2xl shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-slate-100 max-w-sm w-full relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-blue-500 to-indigo-500"></div>
          <div className="flex justify-center mb-6">
            <div className="bg-blue-50 p-3 rounded-xl inline-block">
              <Building2 className="text-blue-600 w-8 h-8" />
            </div>
          </div>
          <h2 className="text-2xl font-black text-slate-800 mb-2 text-center">
            {notApproved ? 'Compte en attente' : 'Connexion'}
          </h2>
          <p className="text-slate-500 text-sm text-center mb-6">
            {notApproved
              ? "Votre compte est en cours d'examen."
              : isVerifyingLink 
                ? "Vérification du lien de connexion..." 
                : "Veuillez vous authentifier pour accéder à l'application."}
          </p>
          
          {notApproved ? (
            <div className="text-center">
              <div className="bg-amber-50 text-amber-700 p-4 rounded-xl text-sm mb-6 border border-amber-200">
                <span className="block font-bold mb-1 leading-tight">En attente d'approbation</span>
                Votre compte a été créé avec succès mais un administrateur doit l'approuver avant que vous ne puissiez accéder à l'application.
              </div>
              <button 
                onClick={() => { setNotApproved(false); handleLogout(); }}
                className="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-3 px-4 rounded-lg transition-colors flex items-center justify-center"
              >
                Retour
              </button>
            </div>
          ) : isVerifyingLink ? (
            <div className="flex justify-center py-8">
              <Loader2 className="w-8 h-8 text-blue-600 animate-spin" />
            </div>
          ) : (
            <>
              {authError && (
                <div className="bg-red-50 text-red-600 p-3 rounded-lg text-sm mb-4 border border-red-100">
                  {authError}
                </div>
              )}

              {emailLinkSent ? (
                <div className="bg-emerald-50 text-emerald-700 p-4 rounded-xl text-sm mb-6 border border-emerald-200 text-center">
                  <span className="block font-bold mb-1 leading-tight">Email envoyé !</span>
                  Vérifiez votre boîte de réception <strong>{loginEmail}</strong> et cliquez sur le lien pour vous connecter.
                </div>
              ) : (
                <div className="flex flex-col gap-4 py-2">
                  <form onSubmit={handleEmailLogin} className="space-y-3">
                    <div>
                      <input 
                        type="email" 
                        value={loginEmail}
                        onChange={(e) => setLoginEmail(e.target.value)}
                        placeholder="votre@email.com"
                        className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all placeholder:text-slate-400 text-sm"
                        required
                      />
                    </div>
                    <button 
                      type="submit"
                      disabled={!loginEmail}
                      className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white font-bold py-3 px-4 rounded-lg transition-colors shadow-sm text-sm"
                    >
                      Recevoir un lien de connexion
                    </button>
                  </form>
                  
                  <div className="flex items-center gap-3 w-full py-2">
                    <div className="flex-1 h-px bg-slate-100"></div>
                    <span className="text-xs text-slate-400 font-medium uppercase tracking-wider">ou</span>
                    <div className="flex-1 h-px bg-slate-100"></div>
                  </div>

                  <button 
                    onClick={handleLogin}
                    className="w-full bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold py-3 px-4 rounded-lg transition-colors flex items-center justify-center gap-3 shadow-sm text-sm"
                  >
                    <svg className="w-5 h-5" viewBox="0 0 24 24">
                      <path fill="currentColor" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                      <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                      <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" />
                      <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" />
                    </svg>
                    Continuer avec Google
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-20">
        <div className="max-w-5xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2">
              <div className="bg-blue-600 p-2 rounded-lg">
                <Building2 className="text-white w-5 h-5" />
              </div>
              <h1 className="text-xl font-bold tracking-tight text-slate-800 hidden sm:block">Développement commercial</h1>
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
              {userData?.role === 'admin' && (
                <button
                  onClick={() => setActiveModule('admin')}
                  className={`px-4 py-2 font-medium text-sm rounded-lg transition-colors ${
                    activeModule === 'admin' ? 'bg-purple-50 text-purple-700' : 'text-slate-600 hover:bg-slate-50'
                  }`}
                >
                  Administration
                </button>
              )}
            </nav>
          </div>
          
          <div className="flex items-center gap-4">
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
            
            <button
              onClick={handleLogout}
              className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"
              title="Se déconnecter"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
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
              {userData?.role === 'admin' && (
                <button
                  onClick={() => setActiveModule('admin')}
                  className={`flex-1 py-3 font-medium text-sm transition-colors text-center ${
                    activeModule === 'admin' ? 'border-b-2 border-purple-600 text-purple-700' : 'text-slate-600 bg-slate-50'
                  }`}
                >
                  Admin
                </button>
              )}
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
                        
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-4 pt-4 border-t border-slate-100">
                          <div>
                            <label className="block text-xs font-semibold text-slate-500 uppercase mb-1">Zone géographique (Optionnel)</label>
                            <div className="relative">
                              <div className="absolute inset-y-0 left-3 flex items-center pointer-events-none">
                                <MapPin className="w-4 h-4 text-slate-400" />
                              </div>
                              <input
                                type="text"
                                value={geoFilter}
                                onChange={(e) => setGeoFilter(e.target.value)}
                                placeholder="Code postal ou départ. (ex: 75, 13008)"
                                className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all placeholder:text-slate-400 text-sm"
                              />
                            </div>
                          </div>
                          <div>
                            <label className="block text-xs font-semibold text-slate-500 uppercase mb-1">Date de création</label>
                            <div className="relative">
                              <div className="absolute inset-y-0 left-3 flex items-center pointer-events-none">
                                <Calendar className="w-4 h-4 text-slate-400" />
                              </div>
                              <select
                                value={dateFilter}
                                onChange={(e) => setDateFilter(e.target.value as any)}
                                className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all text-sm appearance-none"
                              >
                                <option value="all">Toutes dates</option>
                                <option value="6_months">Moins de 6 mois</option>
                                <option value="1_year">Moins d'1 an</option>
                                <option value="2_years">Moins de 2 ans</option>
                                <option value="5_years">Moins de 5 ans</option>
                                <option value="more_5_years">Plus de 5 ans</option>
                              </select>
                            </div>
                          </div>
                        </div>
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

                    <form onSubmit={(e) => handleSearch(e)} className="bg-white p-4 rounded-xl border border-slate-200 flex flex-col md:flex-row gap-4 items-end">
                      <div className="flex-1 w-full">
                        <label className="block text-xs font-semibold text-slate-500 uppercase mb-1">Zone géographique (Optionnel)</label>
                        <div className="relative">
                          <div className="absolute inset-y-0 left-3 flex items-center pointer-events-none">
                            <MapPin className="w-4 h-4 text-slate-400" />
                          </div>
                          <input
                            type="text"
                            value={geoFilter}
                            onChange={(e) => setGeoFilter(e.target.value)}
                            placeholder="Code postal ou départ. (ex: 75, 13008)"
                            className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all placeholder:text-slate-400 text-sm"
                          />
                        </div>
                      </div>
                      <div className="flex-1 w-full">
                        <label className="block text-xs font-semibold text-slate-500 uppercase mb-1">Date de création</label>
                        <div className="relative">
                          <div className="absolute inset-y-0 left-3 flex items-center pointer-events-none">
                            <Calendar className="w-4 h-4 text-slate-400" />
                          </div>
                          <select
                            value={dateFilter}
                            onChange={(e) => setDateFilter(e.target.value as any)}
                            className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all text-sm appearance-none"
                          >
                            <option value="all">Toutes dates</option>
                            <option value="6_months">Moins de 6 mois</option>
                            <option value="1_year">Moins d'1 an</option>
                            <option value="2_years">Moins de 2 ans</option>
                            <option value="5_years">Moins de 5 ans</option>
                            <option value="more_5_years">Plus de 5 ans</option>
                          </select>
                        </div>
                      </div>
                      <button
                        type="submit"
                        disabled={isSearching}
                        className="px-6 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white rounded-lg font-bold transition-all h-[42px] flex items-center justify-center min-w-[120px] shadow-md shadow-blue-600/20"
                      >
                        {isSearching ? <Loader2 className="w-4 h-4 animate-spin" /> : 'Filtrer'}
                      </button>
                    </form>

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
                                  <button
                                    onClick={(e) => {
                                      e.preventDefault();
                                      e.stopPropagation();
                                      const queryName = result.nom_complet || result.nom_raison_sociale || result.siren;
                                      setActiveModule('ads');
                                      setAdsSearchType('company');
                                      setAdsQuery(queryName);
                                      triggerAdsSearch(queryName, 'company');
                                    }}
                                    className="text-[10px] cursor-pointer font-bold text-amber-700 bg-amber-50 hover:bg-amber-100 hover:text-amber-800 hover:border-amber-300 px-2 py-0.5 rounded border border-amber-200 flex items-center gap-1 transition-all"
                                    title="Chercher des permis de construire pour cette société"
                                  >
                                    Rechercher Permis
                                    <HardHat className="w-3 h-3" />
                                  </button>
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
                                    <span className="font-semibold text-slate-700">{getNafLabel(result.activite_principale)}</span>
                                    {result.activite_principale_naf25 && (
                                      <span className="text-xs bg-slate-100 text-slate-500 px-1.5 rounded">NAF 2025: {result.activite_principale_naf25}</span>
                                    )}
                                  </div>
                                </div>
                                {result.siege && (result.siege.adresse || result.siege.code_postal) && (
                                  <div className="flex gap-1.5 items-start">
                                    <MapPin className="w-4 h-4 text-slate-400 flex-shrink-0 mt-0.5" />
                                    <div className="text-slate-600">
                                      {[result.siege.adresse, result.siege.code_postal, result.siege.commune].filter(Boolean).join(', ')}
                                    </div>
                                  </div>
                                )}
                              </div>
                            </div>
                            
                            <div className="mt-4 pt-4 border-t border-slate-50 flex items-center justify-between text-[10px] text-slate-400 relative z-10">
                              <span>Source: "{result._search_term}"</span>
                              <div className="flex gap-3">
                                <button
                                  onClick={(e) => {
                                    e.preventDefault();
                                    e.stopPropagation();
                                    const queryName = result.nom_complet || result.nom_raison_sociale || result.siren;
                                    setActiveModule('ads');
                                    setAdsSearchType('company');
                                    setAdsQuery(queryName);
                                    triggerAdsSearch(queryName, 'company');
                                  }}
                                  className="flex items-center gap-1 hover:text-amber-600 transition-colors font-semibold"
                                >
                                  Voir les permis <HardHat className="w-3 h-3" />
                                </button>
                                <span className="flex items-center gap-1 group-hover:text-blue-600">Voir sur Pappers <ArrowRight className="w-3 h-3" /></span>
                              </div>
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
                <p className="text-slate-600 mb-4">Module connecté à la base de données. Interrogation directe de l'API Supabase pour un affichage instantané sans téléchargement.</p>
              </div>

              <div className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 overflow-hidden max-w-3xl mx-auto mb-8">
                <div className="flex border-b border-slate-100">
                  <button 
                    onClick={() => { setAdsSearchType('company'); setAdsQuery(''); }}
                    className={`flex-1 py-4 font-medium text-sm transition-all flex flex-col items-center gap-1 ${adsSearchType === 'company' ? 'bg-amber-50 text-amber-700 border-b-2 border-amber-600' : 'text-slate-500 hover:bg-slate-50'}`}
                  >
                    <Building2 className="w-5 h-5" />
                    Par Porteur
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
                    Zone / Commune
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
                      placeholder={adsSearchType === 'company' ? 'Ex: NEXITY, SCI Horizon' : adsSearchType === 'permit' ? 'Ex: PC-2023-xxxx' : 'Ex: 75010, Paris, Lyon...'}
                      value={adsQuery}
                      onChange={(e) => setAdsQuery(e.target.value)}
                    />
                  </div>

                  <div className="flex items-center justify-between">
                    <p className="text-xs flex items-center gap-1 font-medium px-2 py-1 rounded bg-green-100 text-green-700">
                      <Server className="w-3.5 h-3.5" />
                      Client API Prêt
                    </p>
                    <button
                      type="submit"
                      disabled={isSearchingAds}
                      className="px-6 py-2.5 bg-amber-600 hover:bg-amber-700 disabled:bg-amber-400 text-white rounded-xl font-bold shadow-md shadow-amber-600/20 transition-all cursor-pointer flex items-center gap-2"
                    >
                      {isSearchingAds ? <><Loader2 className="w-4 h-4 animate-spin"/> Recherche...</> : 'Lancer la recherche'}
                    </button>
                  </div>
                </form>
              </div>

              {/* Module Code Postal / Commune */}
              <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden max-w-3xl mx-auto mb-8 p-6">
                <div className="flex items-center gap-2 mb-4">
                  <MapPin className="w-5 h-5 text-indigo-500" />
                  <h3 className="text-lg font-bold text-slate-800">Trouver un code postal ou une commune</h3>
                </div>
                <form onSubmit={handleGeoUtilSearch} className="flex gap-4">
                  <div className="relative flex-1">
                    <input
                      type="text"
                      className="w-full pl-4 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-400 font-medium text-sm"
                      placeholder="Saisissez une commune (ex: 'Saint-Etienne', 'paris') ou un code postal (ex: 75010)"
                      value={geoUtilQuery}
                      onChange={(e) => setGeoUtilQuery(e.target.value)}
                    />
                  </div>
                  <button
                    type="submit"
                    disabled={isSearchingGeoUtil}
                    className="px-6 py-3 bg-indigo-600 hover:bg-indigo-700 disabled:bg-indigo-400 text-white rounded-xl font-bold shadow-md shadow-indigo-600/20 transition-all cursor-pointer flex items-center justify-center min-w-[120px]"
                  >
                    {isSearchingGeoUtil ? <Loader2 className="w-5 h-5 animate-spin"/> : 'Chercher'}
                  </button>
                </form>

                {geoUtilResults.length > 0 && (
                  <div className="mt-4 pt-4 border-t border-slate-100">
                    <p className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">Résultats</p>
                    <div className="flex flex-wrap gap-2">
                      {geoUtilResults.map((geo, idx) => (
                        <div key={idx} className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm">
                          <span className="font-bold text-slate-700">{geo.nom}</span>
                          <span className="text-slate-400 mx-2">•</span>
                          <span className="text-indigo-600 font-medium">{geo.codesPostaux?.join(', ') || geo.code}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {adsResults.length > 0 && (
                <div className="mt-8 space-y-6">
                  {/* Filtres */}
                  <div className="bg-white border border-slate-200 p-4 rounded-xl flex flex-wrap gap-4 shadow-sm relative z-20">
                    <div className="flex-1 min-w-[200px]">
                      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Logements</label>
                      <select 
                        value={adsFilterLogements} 
                        onChange={(e) => setAdsFilterLogements(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
                      >
                        <option value="all">Tous (0 et +)</option>
                        <option value="lt_15">Moins de 15</option>
                        <option value="15_50">15 à 50</option>
                        <option value="gt_50">Plus de 50</option>
                      </select>
                    </div>
                    <div className="flex-1 min-w-[200px]">
                      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Date (Dépôt/Obtention)</label>
                      <select 
                        value={adsFilterDate} 
                        onChange={(e) => setAdsFilterDate(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
                      >
                        <option value="all">Toutes dates</option>
                        <option value="3_months">Depuis 3 mois</option>
                        <option value="6_months">Depuis 6 mois</option>
                        <option value="1_year">Depuis 1 an</option>
                        <option value="3_years">Depuis 3 ans</option>
                      </select>
                    </div>
                    <div className="flex-1 min-w-[200px]">
                      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">État du dossier</label>
                      <select 
                        value={adsFilterStatus} 
                        onChange={(e) => setAdsFilterStatus(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
                      >
                        <option value="all">Tous statuts</option>
                        <option value="autorise">Autorisé</option>
                        <option value="en_cours">En cours / Projet</option>
                        <option value="ouverture">Ouverture de chantier</option>
                        <option value="refuse">Refusé</option>
                        <option value="depot">Dépôt</option>
                      </select>
                    </div>
                  </div>

                  {/* Statistiques rapides */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div className="bg-white border border-slate-200 p-4 rounded-xl flex flex-col justify-center items-center text-center shadow-sm">
                      <span className="text-3xl font-black text-slate-800">{filteredAdsResults.length}</span>
                      <span className="text-xs font-medium text-slate-500 uppercase tracking-wider mt-1">Dossiers</span>
                    </div>
                    <div className="bg-white border border-slate-200 p-4 rounded-xl flex flex-col justify-center items-center text-center shadow-sm">
                      <span className="text-3xl font-black text-emerald-600">
                        {filteredAdsResults.filter(r => r.statut?.toLowerCase().includes('autoris') || r.statut?.toLowerCase().includes('accord')).length}
                      </span>
                      <span className="text-xs font-medium text-slate-500 uppercase tracking-wider mt-1">Autorisés</span>
                    </div>
                    <div className="bg-white border border-slate-200 p-4 rounded-xl flex flex-col justify-center items-center text-center shadow-sm">
                      <span className="text-3xl font-black text-blue-600">
                        {filteredAdsResults.reduce((acc, curr) => acc + (curr.logements || 0), 0)}
                      </span>
                      <span className="text-xs font-medium text-slate-500 uppercase tracking-wider mt-1">Logements</span>
                    </div>
                    <div className="bg-white border border-slate-200 p-4 rounded-xl flex flex-col justify-center items-center text-center shadow-sm">
                      <span className="text-3xl font-black text-purple-600">
                        {(() => {
                           const totalSurface = filteredAdsResults.reduce((acc, curr) => acc + (curr.surface || 0), 0);
                           return totalSurface > 1000 ? `${Math.floor(totalSurface / 1000)}k` : totalSurface;
                        })()}
                      </span>
                      <span className="text-xs font-medium text-slate-500 uppercase tracking-wider mt-1">m² de surface</span>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* Graphique Types */}
                    <div className="bg-white border border-slate-200 p-5 rounded-2xl shadow-sm h-72">
                      <h4 className="font-bold text-slate-800 text-sm mb-4">Répartition par type</h4>
                      <ResponsiveContainer width="100%" height="80%">
                        <PieChart>
                          <Pie
                            data={Object.entries(filteredAdsResults.reduce((acc, curr) => {
                              const type = curr.type || 'Autre';
                              acc[type] = (acc[type] || 0) + 1;
                              return acc;
                            }, {} as Record<string, number>)).map(([name, value]) => ({ name, value }))}
                            cx="50%"
                            cy="50%"
                            outerRadius={80}
                            fill="#8884d8"
                            dataKey="value"
                            label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                          >
                            <Cell key="cell-0" fill="#3b82f6" />
                            <Cell key="cell-1" fill="#f59e0b" />
                            <Cell key="cell-2" fill="#10b981" />
                            <Cell key="cell-3" fill="#6366f1" />
                            <Cell key="cell-4" fill="#8b5cf6" />
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>

                    {/* Graphique Statuts */}
                    <div className="bg-white border border-slate-200 p-5 rounded-2xl shadow-sm h-72">
                      <h4 className="font-bold text-slate-800 text-sm mb-4">États des dossiers</h4>
                      <ResponsiveContainer width="100%" height="80%">
                        <BarChart
                          data={Object.entries(filteredAdsResults.reduce((acc, curr) => {
                            const statut = curr.statut || 'Inconnu';
                            acc[statut] = (acc[statut] || 0) + 1;
                            return acc;
                          }, {} as Record<string, number>)).map(([name, value]) => ({ name, value }))}
                          layout="vertical"
                        >
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                          <XAxis type="number" />
                          <YAxis dataKey="name" type="category" width={100} fontSize={10} />
                          <Tooltip />
                          <Bar dataKey="value" fill="#f59e0b" radius={[0, 4, 4, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="flex justify-between items-end mb-2 mt-4 mt-8">
                    <div>
                      <h3 className="text-xl font-bold text-slate-800">Dossiers d'urbanisme</h3>
                      <p className="text-sm text-slate-500">
                        {filteredAdsResults.length} résultats affichés
                      </p>
                    </div>
                    <div className="flex bg-slate-100 rounded-lg p-1">
                      <button
                        onClick={() => setAdsViewMode('list')}
                        className={`px-4 py-1.5 rounded-md text-sm font-medium transition-colors ${adsViewMode === 'list' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
                      >
                        Liste
                      </button>
                      <button
                        onClick={() => setAdsViewMode('map')}
                        className={`px-4 py-1.5 rounded-md text-sm font-medium transition-colors ${adsViewMode === 'map' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
                      >
                        Carte
                      </button>
                    </div>
                  </div>
                  
                  {adsViewMode === 'map' ? (
                    <MapComponent results={filteredAdsResults} />
                  ) : (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {filteredAdsResults.map((permit, idx) => (
                        <motion.div
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          transition={{ delay: Math.min(idx * 0.05, 0.5) }}
                          key={`${permit.id}-${idx}`}
                          className="bg-white border border-slate-200 p-5 rounded-2xl shadow-sm hover:shadow-md transition-shadow"
                        >
                        <div className="flex justify-between items-start mb-3">
                          <span className="text-xs font-bold text-slate-500 bg-slate-100 px-2 py-1 rounded">
                            {permit.id}
                          </span>
                          <span className={`text-xs font-bold px-2 py-1 rounded ${permit.statut === 'Autorisé' ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                            {permit.statut || 'Inconnu'}
                          </span>
                        </div>
                        <h4 className="font-bold text-lg text-slate-800 mb-1">{permit.type || 'Permis'}</h4>
                        <div className="text-sm font-medium text-slate-600 mb-4">
                          <button
                            onClick={(e) => {
                              e.preventDefault();
                              e.stopPropagation();
                              const query = permit.siren_demandeur || permit.demandeur;
                              if (query) {
                                setActiveModule('entreprises');
                                setCompanies([query]);
                                handleSearch(undefined, [query]);
                              }
                            }}
                            className="hover:text-blue-600 hover:underline inline-flex items-center gap-1 transition-colors text-left focus:outline-none"
                            title="Rechercher ce promoteur dans Entreprises et Participations"
                          >
                            {permit.demandeur || 'Non renseigné'}
                            {permit.demandeur && <Search className="w-3 h-3 text-slate-400" />}
                          </button>
                        </div>
                        
                        <div className="space-y-2 text-sm">
                          <div className="flex items-start gap-2 text-slate-600">
                            <MapPin className="w-4 h-4 mt-0.5 text-slate-400 shrink-0" />
                            <span className="line-clamp-2">{permit.adresse || permit.commune || 'Adresse non renseignée'}</span>
                          </div>
                          
                          {(permit.date_depot || permit.date_obtention) && (
                            <div className="flex items-center gap-2 text-slate-600">
                              <Calendar className="w-4 h-4 text-slate-400 shrink-0" />
                              <span>
                                {permit.date_depot && `Dépôt: ${formatDate(permit.date_depot)}`}
                                {permit.date_depot && permit.date_obtention && ' • '}
                                {permit.date_obtention && `Décision: ${formatDate(permit.date_obtention)}`}
                              </span>
                            </div>
                          )}
                        </div>
                        
                        {permit.description && (
                          <div className="mt-4 pt-4 border-t border-slate-100">
                            <p className="text-xs text-slate-500 line-clamp-3">{permit.description}</p>
                          </div>
                        )}
                      </motion.div>
                    ))}
                  </div>
                  )}
                </div>
              )}
            </motion.div>
          )}
          {activeModule === 'admin' && (
            <motion.div
              key="admin-module"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
            >
              <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden max-w-3xl mx-auto mb-8">
                <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-purple-100 text-purple-600 rounded-lg">
                      <LogOut className="w-5 h-5 -scale-x-100" />
                    </div>
                    <div>
                      <h2 className="text-xl font-bold text-slate-800">Administration</h2>
                      <p className="text-slate-500 text-sm">Gestion des utilisateurs</p>
                    </div>
                  </div>
                  <button 
                    onClick={fetchAdminUsers}
                    className="p-2 text-slate-400 hover:text-purple-600 hover:bg-purple-50 rounded-lg transition-colors"
                    title="Rafraîchir"
                  >
                    <RefreshCw className={`w-5 h-5 ${isAdminLoading ? 'animate-spin' : ''}`} />
                  </button>
                </div>
                
                <div className="p-6">
                  <h3 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-4">
                    Comptes en attente de validation ({unapprovedUsers.length})
                  </h3>
                  
                  {isAdminLoading ? (
                    <div className="flex justify-center py-8">
                      <Loader2 className="w-8 h-8 text-purple-500 animate-spin" />
                    </div>
                  ) : unapprovedUsers.length === 0 ? (
                    <div className="text-center py-12 px-4 rounded-xl border-2 border-dashed border-slate-200 bg-slate-50">
                      <p className="text-slate-500 font-medium">Aucun compte en attente d'approbation pour le moment.</p>
                    </div>
                  ) : (
                    <div className="space-y-3 mb-8">
                      {unapprovedUsers.map(user => (
                        <div key={user.id} className="flex flex-col sm:flex-row items-start sm:items-center justify-between p-4 bg-slate-50 rounded-xl border border-slate-100 hover:border-slate-200 transition-colors gap-4">
                          <div className="flex flex-col">
                            <span className="font-bold text-slate-800 break-all">{user.email}</span>
                            <span className="text-sm text-slate-500">
                              Inscription : {new Date(user.createdAt).toLocaleDateString()}
                            </span>
                          </div>
                          <button
                            onClick={() => approveUser(user.id)}
                            className="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg font-bold shadow-md shadow-purple-600/20 transition-all min-w-[120px]"
                          >
                            Approuver
                          </button>
                        </div>
                      ))}
                    </div>
                  )}

                  <h3 className="text-sm font-bold text-slate-500 uppercase tracking-wider mt-12 mb-4 pt-8 border-t border-slate-200">
                    Comptes approuvés ({approvedUsers.length})
                  </h3>
                  
                  {isAdminLoading ? (
                    <div className="flex justify-center py-8">
                      <Loader2 className="w-8 h-8 text-purple-500 animate-spin" />
                    </div>
                  ) : approvedUsers.length === 0 ? (
                    <div className="text-center py-12 px-4 rounded-xl border-2 border-dashed border-slate-200 bg-slate-50">
                      <p className="text-slate-500 font-medium">Aucun compte approuvé.</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {approvedUsers.map(user => (
                        <div key={user.id} className="flex flex-col sm:flex-row items-start sm:items-center justify-between p-4 bg-white rounded-xl border border-slate-200 hover:border-slate-300 transition-colors gap-4 shadow-sm">
                          <div className="flex flex-col">
                            <span className="font-bold text-slate-800 break-all">{user.email} {user.role === 'admin' && <span className="ml-2 px-2 py-0.5 bg-purple-100 text-purple-700 text-xs rounded-full">Admin</span>} {user.id === auth.currentUser?.uid && <span className="ml-2 px-2 py-0.5 bg-blue-100 text-blue-700 text-xs rounded-full">Vous</span>}</span>
                            <span className="text-sm text-slate-500">
                              Inscription : {new Date(user.createdAt).toLocaleDateString()}
                            </span>
                          </div>
                          {user.id !== auth.currentUser?.uid && (
                            <button
                              onClick={() => revokeUser(user.id)}
                              className="bg-white border border-red-200 hover:bg-red-50 text-red-600 px-4 py-2 rounded-lg font-bold transition-all min-w-[120px]"
                            >
                              Retirer l'accès
                            </button>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
      
      <footer className="max-w-5xl mx-auto px-4 py-8 text-center border-t border-slate-200 mt-12 bg-white flex flex-col sm:flex-row items-center justify-between">
        <p className="text-slate-500 text-sm">&copy; {new Date().getFullYear()} Développement commercial.</p>
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

