# vault-viewer

Petit chargeur (PWA) pour ouvrir mon Vault Viewer sur le téléphone. Il se connecte à mon OneDrive avec mon compte Microsoft et lit les notes là-bas.

**Ce dépôt ne contient que du code, jamais de notes ni de données.** L'identifiant client dans `config.js` n'est pas un secret.

État : étape 2 — le chargeur ouvre la vraie page du vault, en lecture seule.

Principe : `index.html` connecte à Microsoft, `graphfs.js` imite sur OneDrive la petite partie de l'API
File System Access dont la page se sert (inventaire par `/delta`, cache IndexedDB par eTag, écriture
avec `If-Match` sur une liste blanche), puis `Vault Viewer.html` est lu dans OneDrive et ouvert dans un cadre.
La page du vault n'est pas modifiée : elle utilise sa porte `window.__testRootPromise`.

- `index.html`, `loader.js` : le chargeur
- `graphfs.js` : OneDrive vu comme un dossier
- `essai.html`, `essai.js` : page d'essai de connexion (diagnostic)
- `config.js` : réglages (identifiant client de l'application Microsoft)
- `lib/msal-browser.min.js` : MSAL.js 4.30.0 (Microsoft, licence MIT), copie fixe
- `sw.js`, `manifest.webmanifest`, `icons/` : installation sur l'écran d'accueil
