# Gestion des intrants (v2)

Application web installable (PWA) de gestion des stocks de produits de santé, avec les données dans un Google Sheet (via Google Apps Script).

## Fonctions
- Comptes : un gestionnaire (tous les droits) et des utilisateurs (saisie et consultation). Le premier écran crée le compte gestionnaire.
- Tableau de bord : alertes, statistiques du mois, produits les plus utilisés, mouvements sur 6 mois.
- Stock par lot avec date de péremption. Les sorties prennent d'abord les lots qui expirent en premier et jamais les lots périmés.
- Stock de sécurité par produit, alerte (rupture, seuil atteint, péremption sous 3 mois, périmé), notifications sur l'appareil, e-mail quotidien aux gestionnaires.
- Consommation moyenne mensuelle (3 derniers mois) et calcul des besoins selon la couverture choisie.
- Bordereaux d'entrée et de sortie à plusieurs produits, historique avec l'utilisateur, exports CSV.

## Fichiers
- `index.html` : l'application. L'URL de l'API est en haut du script (`API`).
- `Code.gs` : backend Apps Script, à coller dans le Google Sheet.
- `manifest.webmanifest`, `sw.js`, `icons/` : installation sur l'écran d'accueil.

## Installation
1. Google Sheet > Extensions > Apps Script : remplacez le contenu de `Code.gs` par celui de ce dépôt. Déployer > Gérer les déploiements > Modifier > Nouvelle version (accès : Tout le monde).
2. Pour les e-mails d'alerte : dans l'éditeur, choisissez la fonction `installerAlerteQuotidienne` et cliquez sur Exécuter une fois (autorisez l'envoi d'e-mails). Renseignez l'e-mail du gestionnaire.
3. Poussez le dossier sur GitHub, puis Settings > Pages > branche `main`, racine.
4. Ouvrez l'adresse GitHub Pages, créez le compte gestionnaire, puis ajoutez les produits (stock de départ et stock de sécurité) et les utilisateurs.

## Feuilles du Google Sheet
`Produits`, `Mouvements` et `Utilisateurs` sont créées ou mises à jour automatiquement. Les mots de passe y sont stockés hachés (SHA-256 avec sel). Limitez l'accès à ce Sheet aux administrateurs.

## Limites
Les notifications sur l'appareil s'affichent à l'ouverture de l'application. Hors application, les alertes arrivent par e-mail. L'URL de l'API est publique, mais toutes les actions exigent une connexion.
