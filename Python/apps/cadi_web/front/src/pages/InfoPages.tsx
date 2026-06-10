import { useState, type FormEvent, type ReactNode } from 'react';
import {
  AlertTriangle,
  CheckCircle,
  FileText,
  HelpCircle,
  Mail,
  MapPin,
  MessageCircle,
  Phone,
  Scale,
  Send,
  ShieldCheck,
} from 'lucide-react';

export type InfoPage = 'terms' | 'privacy' | 'legal' | 'faq' | 'contact';

interface InfoPageProps {
  onNavigate: (page: InfoPage) => void;
}

function InfoHero({
  Icon,
  title,
  subtitle,
}: {
  Icon: React.ElementType;
  title: string;
  subtitle: string;
}) {
  return (
    <div className="text-center mb-10">
      <div className="inline-flex items-center justify-center w-20 h-20 rounded-full mb-6" style={{ background: '#fff7ed' }}>
        <Icon className="w-10 h-10" style={{ color: '#f97316' }} strokeWidth={1.8} />
      </div>
      <h1 className="text-4xl font-black mb-3" style={{ color: '#1e2a4a' }}>{title}</h1>
      <p className="text-base" style={{ color: '#64748b' }}>{subtitle}</p>
    </div>
  );
}

function InfoCard({
  title,
  icon,
  children,
  accent = '#f97316',
}: {
  title: string;
  icon?: ReactNode;
  children: ReactNode;
  accent?: string;
}) {
  return (
    <section
      className="rounded-xl p-6"
      style={{ background: '#fff', border: '1px solid #e8edf2', boxShadow: '0 1px 4px rgba(30,42,74,0.06)' }}
    >
      <h2 className="flex items-center gap-2 text-xl font-black mb-4" style={{ color: '#1e293b' }}>
        {icon && <span style={{ color: accent }}>{icon}</span>}
        {title}
      </h2>
      <div className="space-y-3 text-sm leading-relaxed" style={{ color: '#475569' }}>
        {children}
      </div>
    </section>
  );
}

function AccordionCard({
  title,
  questions,
}: {
  title: string;
  questions: { q: string; a: string }[];
}) {
  return (
    <InfoCard title={title}>
      <div>
        {questions.map((item, index) => (
          <details key={item.q} open={index === 0} className="group" style={{ borderTop: index ? '1px solid #e8edf2' : 0 }}>
            <summary className="flex min-h-12 cursor-pointer list-none items-center justify-between gap-4 py-3 text-left text-sm font-black" style={{ color: '#1e293b' }}>
              {item.q}
              <span className="text-lg font-black" style={{ color: '#f97316' }}>+</span>
            </summary>
            <p className="pb-4 pr-8 text-sm leading-relaxed" style={{ color: '#64748b' }}>{item.a}</p>
          </details>
        ))}
      </div>
    </InfoCard>
  );
}

export function FAQPage({ onNavigate }: InfoPageProps) {
  const categories = [
    {
      title: 'Compte et acces',
      questions: [
        { q: 'Qui peut utiliser CADI Web ?', a: 'CADI Web est prevu pour les agents ou utilisateurs autorises qui traitent des donnees de formation INSTN / CEA. L application n est pas un outil public.' },
        { q: 'Pourquoi dois-je me connecter ?', a: 'La connexion permet de limiter l acces aux modules sensibles, de personnaliser certaines informations utilisateur et d eviter une utilisation non autorisee des donnees.' },
        { q: 'Puis-je modifier mes informations ?', a: 'Oui. Les informations utilisateur modifiables se trouvent dans Parametres. Les donnees institutionnelles ou droits d acces doivent rester geres par le referent applicatif.' },
        { q: 'Que faire si je ne peux pas me connecter ?', a: 'Verifiez d abord vos identifiants et l etat du serveur. Si le probleme persiste, contactez le referent CADI Web avec votre nom, unite, email et une capture du message affiche.' },
      ],
    },
    {
      title: 'IRIS et selection des fichiers',
      questions: [
        { q: 'Quels fichiers dois-je choisir pour IRIS ?', a: 'Choisissez uniquement les fichiers Excel correspondant au code de la ligne selectionnee : R04110 pour Sessions, R0304 pour Formations, R04301 pour Ventes et R04500 pour Inscriptions.' },
        { q: 'Pourquoi un fichier est refuse ?', a: 'Le nom du fichier ne correspond pas au code attendu. Exemple : si vous traitez Sessions, le fichier doit commencer par R04110. Un fichier R04301 ou R0304 doit etre refuse pour eviter un traitement faux.' },
        { q: 'Puis-je selectionner plusieurs fichiers ?', a: 'Oui si le traitement l attend, par exemple plusieurs annees de sessions. Tous les fichiers selectionnes doivent appartenir au meme code et au meme type de donnees.' },
        { q: 'Pourquoi la fenetre de selection ne s ouvre pas ?', a: 'La fenetre de selection est declenchee par le backend FastAPI local. Verifiez que le backend est lance, que le port configure est correct et que le navigateur peut joindre l API.' },
      ],
    },
    {
      title: 'Backend, exports et erreurs',
      questions: [
        { q: 'Que signifie backend indisponible ?', a: 'Le front React ne parvient pas a joindre FastAPI. Relancez le serveur backend, rechargez la page, puis recommencez l action. Si le port a change, alignez la configuration front et backend.' },
        { q: 'Que faire si un traitement echoue ?', a: 'Ne relancez pas au hasard. Notez le module, le code, le fichier source, le message d erreur et l heure. Verifiez ensuite que le fichier n est pas ouvert dans Excel ou Word.' },
        { q: 'Pourquoi Word ou Excel peut bloquer un export ?', a: 'Si un fichier de sortie est deja ouvert, Office peut verrouiller le document. Fermez le fichier concerne, puis relancez le traitement.' },
        { q: 'Comment savoir si le resultat est correct ?', a: 'Controlez le statut affiche, le nom du fichier genere, la date de mise a jour et un echantillon du contenu avant diffusion.' },
      ],
    },
  ];

  return (
    <div className="max-w-4xl mx-auto py-8">
      <InfoHero Icon={HelpCircle} title="Foire aux questions" subtitle="Trouvez rapidement les reponses aux questions frequentes sur CADI Web." />
      <div className="space-y-6">
        {categories.map(category => <AccordionCard key={category.title} title={category.title} questions={category.questions} />)}
        <InfoCard title="Vous ne trouvez pas votre reponse ?" icon={<MessageCircle size={21} />}>
          <p>Contactez le support interne avec le module, le code IRIS, le nom exact du fichier, le message d erreur et une capture d ecran.</p>
          <button
            type="button"
            onClick={() => onNavigate('contact')}
            className="mt-2 rounded-lg px-4 py-2 text-sm font-bold text-white"
            style={{ background: '#f97316' }}
          >
            Nous contacter
          </button>
        </InfoCard>
      </div>
    </div>
  );
}

export function TermsPage() {
  return (
    <div className="max-w-4xl mx-auto py-8">
      <InfoHero Icon={FileText} title="Conditions d'utilisation" subtitle="Regles d'utilisation interne de CADI Web - derniere mise a jour : 8 juin 2026." />
      <div className="space-y-6">
        <InfoCard title="Acceptation des conditions" icon={<CheckCircle size={21} />} accent="#16a34a">
          <p>En utilisant CADI Web, vous acceptez de respecter les regles internes INSTN / CEA, les consignes de securite applicables et les usages prevus par les modules de l application.</p>
          <p>Si vous n etes pas habilite a manipuler les donnees concernees, vous ne devez pas lancer de traitement ni consulter les exports.</p>
        </InfoCard>
        <InfoCard title="Utilisation autorisee" icon={<ShieldCheck size={21} />}>
          <ul className="list-disc pl-5 space-y-2">
            <li>Utiliser CADI Web uniquement pour les activites de gestion, controle et production de documents de formation.</li>
            <li>Selectionner uniquement les fichiers necessaires a la mission en cours.</li>
            <li>Verifier le code, le type de fichier et l annee avant chaque traitement.</li>
            <li>Conserver les exports dans les emplacements valides par l organisation.</li>
            <li>Ne pas contourner les controles de fichiers ou les messages de refus.</li>
          </ul>
        </InfoCard>
        <InfoCard title="Responsabilites" icon={<AlertTriangle size={21} />} accent="#eab308">
          <p>L utilisateur reste responsable des fichiers selectionnes, des traitements lances, de la verification des resultats et de la diffusion des documents generes.</p>
          <p>Avant diffusion, il faut controler au minimum le nom du fichier, le module utilise, la date de generation et un echantillon du contenu.</p>
        </InfoCard>
        <InfoCard title="Cadre legal" icon={<Scale size={21} />} accent="#1e2a4a">
          <p>CADI Web est fourni pour un usage professionnel interne. Toute utilisation hors du cadre autorise, toute extraction massive non justifiee ou toute diffusion externe doit etre validee par les responsables concernes.</p>
        </InfoCard>
      </div>
    </div>
  );
}

export function PrivacyPage() {
  return (
    <div className="max-w-4xl mx-auto py-8">
      <InfoHero Icon={ShieldCheck} title="Confidentialite" subtitle="Protection des donnees de formation, fichiers sources et exports generes." />
      <div className="space-y-6">
        <InfoCard title="Donnees sensibles">
          <p>CADI Web peut manipuler des donnees liees aux sessions, formations, inscriptions, ventes, evaluations et documents de bilan. Ces informations doivent rester dans le cadre interne autorise.</p>
          <p>Les fichiers sources et les exports generes ne doivent etre accessibles qu aux personnes ayant un besoin professionnel legitime.</p>
        </InfoCard>
        <InfoCard title="Regles de manipulation">
          <ul className="list-disc pl-5 space-y-2">
            <li>Ne pas envoyer les exports par canal non valide.</li>
            <li>Ne pas deposer de fichiers sensibles dans un dossier personnel non partage ou non sauvegarde.</li>
            <li>Fermer les fichiers Excel ou Word apres controle pour eviter les verrous Office.</li>
            <li>Supprimer ou archiver les copies temporaires quand elles ne sont plus necessaires.</li>
            <li>Signaler rapidement toute erreur de destinataire, de fichier ou de diffusion.</li>
          </ul>
        </InfoCard>
        <InfoCard title="Ce que CADI Web ne doit pas remplacer">
          <p>CADI Web aide a produire et controler les documents, mais ne remplace pas la validation metier. Une verification humaine reste necessaire avant l envoi officiel d un bilan, d un export ou d un fichier consolide.</p>
        </InfoCard>
      </div>
    </div>
  );
}

export function LegalPage() {
  return (
    <div className="max-w-4xl mx-auto py-8">
      <InfoHero Icon={Scale} title="Mentions legales" subtitle="Informations institutionnelles relatives a CADI Web." />
      <div className="space-y-6">
        <InfoCard title="Identification de l'application">
          <ul className="list-disc pl-5 space-y-2">
            <li>Nom : CADI Web.</li>
            <li>Organisation : INSTN / CEA.</li>
            <li>Perimetre : gestion et traitement de donnees de formation.</li>
            <li>Modules : IRIS, EvalStat, Bilans, Formations, Parametres et Messagerie.</li>
            <li>Usage : professionnel interne uniquement.</li>
          </ul>
        </InfoCard>
        <InfoCard title="Responsabilite des contenus">
          <p>Les donnees affichees ou generees dependent des fichiers sources selectionnes, des parametrages locaux et des controles metier effectues par l utilisateur.</p>
          <p>Une erreur de fichier source, de code ou de version peut produire un resultat incomplet ou incorrect.</p>
        </InfoCard>
      </div>
    </div>
  );
}

export function ContactPage() {
  const [sent, setSent] = useState(false);

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setSent(true);
  };

  return (
    <div className="max-w-4xl mx-auto py-8">
      <InfoHero Icon={Mail} title="Contact" subtitle="Une question ou un probleme ? Envoyez les informations utiles au support CADI Web." />
      <InfoCard title="Avant de contacter le support" icon={<HelpCircle size={21} />}>
        <ul className="list-disc pl-5 space-y-2">
          <li>Reproduisez le probleme une fois si possible.</li>
          <li>Notez le module concerne : IRIS, EvalStat, Bilans, Formations, Parametres ou Messagerie.</li>
          <li>Copiez le message d erreur exact.</li>
          <li>Indiquez le nom complet du fichier source et son emplacement si c est pertinent.</li>
          <li>Precisez si le backend FastAPI et le front Vite etaient bien lances.</li>
        </ul>
      </InfoCard>
      <div className="mt-6 grid md:grid-cols-[1.2fr_0.8fr] gap-6">
        <InfoCard title="Envoyer une demande" icon={<Send size={21} />}>
          <form onSubmit={handleSubmit} className="space-y-4">
            <input className="w-full rounded-lg border px-3 py-2 text-sm outline-none" style={{ borderColor: '#e2e8f0' }} placeholder="Nom" required />
            <input className="w-full rounded-lg border px-3 py-2 text-sm outline-none" style={{ borderColor: '#e2e8f0' }} placeholder="Email" type="email" required />
            <input className="w-full rounded-lg border px-3 py-2 text-sm outline-none" style={{ borderColor: '#e2e8f0' }} placeholder="Sujet : exemple IRIS - fichier R04110 refuse" required />
            <textarea className="min-h-32 w-full rounded-lg border px-3 py-2 text-sm outline-none" style={{ borderColor: '#e2e8f0' }} placeholder="Module, action realisee, fichier utilise, message d erreur, resultat attendu..." required />
            <button type="submit" className="w-full rounded-lg px-4 py-2.5 text-sm font-bold text-white" style={{ background: '#f97316' }}>
              Preparer la demande
            </button>
            {sent && <p className="text-sm font-semibold" style={{ color: '#16a34a' }}>Demande preparee. Verifiez les informations avant transmission au referent ou support interne.</p>}
          </form>
        </InfoCard>
        <div className="space-y-6">
          <InfoCard title="Canal principal" icon={<Mail size={21} />}>
            <p>Referent applicatif CADI Web ou support interne de votre unite.</p>
          </InfoCard>
          <InfoCard title="Urgence bloquante" icon={<Phone size={21} />}>
            <p>Si un livrable doit partir rapidement, prevenez directement le referent metier avec le fichier concerne et l heure limite.</p>
          </InfoCard>
          <InfoCard title="Perimetre" icon={<MapPin size={21} />}>
            <p>CADI Web - usages internes INSTN / CEA .</p>
          </InfoCard>
        </div>
      </div>
    </div>
  );
}
