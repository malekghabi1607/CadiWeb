const modules = [
  {
    initial: "I",
    title: "IRIS",
    description: "Exports et donnees de formation.",
    href: "#iris",
  },
  {
    initial: "E",
    title: "EvalStat",
    description: "Evaluations des stagiaires.",
    href: "#evalstat",
  },
  {
    initial: "B",
    title: "Bilans",
    description: "Syntheses et documents.",
    href: "#bilans",
  },
  {
    initial: "F",
    title: "Fiches de couts",
    description: "Lecture depuis sa page dediee.",
    href: "#fdc",
  },
  {
    initial: "P",
    title: "Parametres",
    description: "Configuration et chemins.",
    href: "#parametres",
  },
];

export default function Accueil() {
  return (
    <div className="browser">
      <div className="browser-bar" aria-hidden="true">
        <div className="dots">
          <span className="dot" />
          <span className="dot" />
          <span className="dot" />
        </div>
        <div className="address" />
        <div className="visitor">C</div>
      </div>

      <header className="site-header">
        <button className="menu-button" type="button" aria-label="Ouvrir le menu">
          =
        </button>

        <a className="brand" href="/" aria-label="Accueil CADI">
          <span className="mark">C</span>
          <span className="brand-title">
            <strong>CADI</strong>
            <span>Le logiciel qui simplifie</span>
          </span>
        </a>

        <nav className="header-actions" aria-label="Navigation visiteur">
          <a className="link" href="#modules">
            Modules
          </a>
          <a className="link" href="#presentation">
            Presentation
          </a>
          <a className="login" href="#connexion">
            Connexion
          </a>
        </nav>
      </header>

      <main>
        <section className="hero" id="presentation">
          <h1>Bienvenue sur CADI</h1>
          <p>
            Une interface simple pour presenter les outils CADI et guider les
            utilisateurs vers les bons modules de traitement.
          </p>
        </section>

        <section className="module-grid" id="modules" aria-label="Modules CADI">
          {modules.map((module) => (
            <a className="module-card" href={module.href} key={module.title}>
              <div className="module-icon">{module.initial}</div>
              <h2>{module.title}</h2>
              <p>{module.description}</p>
            </a>
          ))}
        </section>

        <section className="visitor-info" aria-label="Information visiteur">
          <p />
          <p />
          <p />
          <div className="cta-row" id="connexion">
            <a className="cta primary" href="#modules">
              Voir les modules
            </a>
            <a className="cta" href="#connexion">
              Se connecter
            </a>
          </div>
        </section>
      </main>

      <footer className="site-footer">
        <div className="footer-links">
          <a href="#conditions">Conditions</a>
          <span className="footer-dot" />
          <a href="#contact">Contact</a>
        </div>
        <div className="copyright">CADI - Page d'accueil</div>
        <div className="social-links" aria-label="Reseaux sociaux">
          <span className="footer-dot" />
          <span className="footer-dot" />
          <span className="footer-dot" />
        </div>
      </footer>
    </div>
  );
}
