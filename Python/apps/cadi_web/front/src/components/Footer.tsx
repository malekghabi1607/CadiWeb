/**
 * ======================================================================================
 * CADI Web — COMPOSANT FOOTER (PIED DE PAGE)
 * ======================================================================================
 * 
 * Ce composant représente la barre de pied de page globale et persistante de l'application.
 * Il affiche l'année en cours, la version officielle de CADI Web et les mentions
 * de confidentialité institutionnelles de l'INSTN et du CEA Saclay.
 * 
 * Rôle structurel :
 * - S'adapte dynamiquement à la largeur de la barre latérale (Sidebar) pour conserver l'alignement.
 * - S'ancre au bas du viewport utilisateur avec un design épuré.
 * 
 * @author INSTN / CEA Saclay
 * @version 2.4.1
 */

interface FooterProps {
  /** Largeur courante de la barre latérale pour ajuster le décalage à gauche (margin-left) */
  sidebarWidth: number;
}

/**
 * Composant fonctionnel principal pour le Pied de page.
 * 
 * @param props Propriétés du composant s'appuyant sur l'interface FooterProps.
 * @returns Rendu HTML5 sémantique du pied de page.
 */
export default function Footer({ sidebarWidth }: FooterProps) {
  return (
    <footer
      className="flex items-center justify-between px-6 transition-all duration-200"
      style={{
        marginLeft: sidebarWidth,
        borderTop: '1px solid #e8edf2',
        background: '#ffffff',
        height: 40,
        flexShrink: 0,
      }}
    >
      {/* Mention de copyright dynamique INSTN / CEA */}
      <span className="text-xs" style={{ color: '#94a3b8' }}>
        © {new Date().getFullYear()} INSTN / CEA — CADI Web
      </span>
      {/* Avertissement de sécurité sur le traitement des données sensibles */}
      <span className="text-xs" style={{ color: '#94a3b8' }}>
        Usage interne — données sensibles
      </span>
    </footer>
  );
}
