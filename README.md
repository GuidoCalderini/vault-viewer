# vault-viewer

Petit chargeur (PWA) pour ouvrir mon Vault Viewer sur le téléphone. Il se connecte à mon OneDrive avec mon compte Microsoft et lit les notes là-bas.

**Ce dépôt ne contient que du code, jamais de notes ni de données.** L'identifiant client dans `config.js` n'est pas un secret.

État : étape 0, page d'essai de connexion (lecture seule).

- `index.html`, `app.js` : page d'essai
- `config.js` : réglages (identifiant client de l'application Microsoft)
- `lib/msal-browser.min.js` : MSAL.js 4.30.0 (Microsoft, licence MIT), copie fixe
- `sw.js`, `manifest.webmanifest`, `icons/` : installation sur l'écran d'accueil
