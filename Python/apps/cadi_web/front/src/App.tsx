import { useState } from 'react';
import Index from './pages/Index';
import Login from './pages/Login';
import Register from './pages/Register';
import ForgotPassword from './pages/ForgotPassword';
import Sidebar, { Screen, SIDEBAR_W, SIDEBAR_W_COLLAPSED } from './components/Sidebar';
import Header from './components/Header';
import Footer from './components/Footer';
import Accueil from './pages/Accueil';
import IRIS from './pages/IRIS';
import EvalStat from './pages/EvalStat';
import Bilans from './pages/Bilans';
import Formations from './pages/Formations';
import Parametres from './pages/Parametres';
import Messagerie from './pages/Messagerie';
import { ContactPage, FAQPage, InfoPage, LegalPage, PrivacyPage, TermsPage } from './pages/InfoPages';
import { LogOut } from 'lucide-react';

type AuthScreen = 'index' | 'login' | 'register' | 'forgot';

const screenTitles: Record<Screen, string> = {
  accueil:    'Tableau de bord',
  iris:       'IRIS',
  evalstat:   'EvalStat',
  bilans:     'Bilans',
  formations: 'Formations',
  parametres: 'Paramètres',
  messagerie: 'Messagerie',
};

const infoTitles: Record<InfoPage, string> = {
  terms: 'Conditions',
  privacy: 'Confidentialite',
  legal: 'Mentions legales',
  faq: 'FAQ',
  contact: 'Contact',
};

export default function App() {
  const [auth, setAuth]           = useState<AuthScreen>('index');
  const [loggedIn, setLoggedIn]   = useState(false);
  const [screen, setScreen]       = useState<Screen>('accueil');
  const [infoScreen, setInfoScreen] = useState<InfoPage | null>(null);
  const [collapsed, setCollapsed] = useState(false);
  const [logoutToast, setLogoutToast] = useState(false);
  const [currentUser, setCurrentUser] = useState<{ prenom: string; nom: string; email: string; unite?: string } | null>(null);

  const sideW = collapsed ? SIDEBAR_W_COLLAPSED : SIDEBAR_W;
  const navigate = (s: string) => {
    setInfoScreen(null);
    setScreen(s as Screen);
  };
  const navigateInfo = (page: InfoPage) => setInfoScreen(page);

  const handleLogout = () => {
    setLoggedIn(false);
    setAuth('index');
    setLogoutToast(true);
    setTimeout(() => setLogoutToast(false), 2500);
  };

  // ── Auth screens ──
  if (!loggedIn) {
    if (auth === 'index') {
      return <Index onLogin={() => setAuth('login')} onRegister={() => setAuth('register')} />;
    }
    if (auth === 'register') {
      return <Register onBack={() => setAuth('login')} onSuccess={() => setAuth('login')} />;
    }
    if (auth === 'forgot') {
      return <ForgotPassword onBack={() => setAuth('login')} />;
    }
    return (
      <Login
        onLogin={(user) => { setLoggedIn(true); setCurrentUser(user); setScreen('accueil'); }}
        onRegister={() => setAuth('register')}
        onForgot={() => setAuth('forgot')}
      />
    );
  }

  const renderScreen = () => {
    if (infoScreen) {
      switch (infoScreen) {
        case 'terms': return <TermsPage />;
        case 'privacy': return <PrivacyPage />;
        case 'legal': return <LegalPage />;
        case 'faq': return <FAQPage onNavigate={navigateInfo} />;
        case 'contact': return <ContactPage />;
      }
    }

    switch (screen) {
      case 'accueil':    return <Accueil onNavigate={navigate} />;
      case 'iris':       return <IRIS />;
      case 'evalstat':   return <EvalStat />;
      case 'bilans':     return <Bilans currentUser={currentUser} />;
      case 'formations': return <Formations />;
      case 'parametres': return <Parametres currentUser={currentUser} onUpdateUser={setCurrentUser} />;
      case 'messagerie': return <Messagerie currentUser={currentUser} />;
    }
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#f0f3f8' }}>

      {logoutToast && (
        <div className="fixed top-4 right-4 z-50 flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium"
          style={{ background: '#1e2a4a', color: '#fff', boxShadow: '0 8px 24px rgba(0,0,0,0.2)' }}>
          <LogOut size={15} /> Déconnexion effectuée
        </div>
      )}

      <Sidebar
        active={screen}
        onNavigate={s => navigate(s)}
        collapsed={collapsed}
        onToggle={() => setCollapsed(c => !c)}
        onLogout={handleLogout}
        currentUser={currentUser}
      />

      <Header
        sidebarWidth={sideW}
        title={infoScreen ? infoTitles[infoScreen] : screenTitles[screen]}
        onNavigate={navigate}
      />

      <main
        className="flex-1 transition-all duration-200"
        style={{ marginLeft: sideW, paddingTop: 76 }}
      >
        <div className="p-7">
          {renderScreen()}
        </div>
      </main>

      <Footer
        sidebarWidth={sideW}
        onNavigateInfo={navigateInfo}
      />
    </div>
  );
}
