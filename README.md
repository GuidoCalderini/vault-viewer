# vault-viewer

Petit chargeur (PWA) pour ouvrir mon Vault Viewer sur le téléphone. Il se connecte à mon OneDrive avec mon compte Microsoft et lit les notes là-bas.

**Ce dépôt ne contient que du code, jamais de notes ni de données.** L'identifiant client dans `config.js` n'est pas un secret.

État : étape 3 + hors ligne (2026-11-07). Le chargeur ouvre la vraie page du vault et peut écrire partout dans le vault, sauf la page elle-même, `.obsidian/` et `Calibre Library/`.

- **Connexion par défaut** ; si elle est impossible (pas de réseau, connexion expirée, trop lent), l'écran propose « Travailler hors ligne » (`?horsligne=1`) : la page s'ouvre avec la version gardée dans le téléphone, et les changements attendent la prochaine connexion.
- **Gardé dans le téléphone** : l'inventaire, toutes les notes texte (gardées d'avance après chaque synchro) et les 3 derniers gros fichiers ouverts (livres).
- **Conflits** : si seul le téléphone a changé un fichier depuis la dernière synchro, sa version est gardée. Si les deux ont changé : fusion ligne par ligne ; là où les deux ont touché les mêmes lignes, l'ordinateur gagne si les deux changements ont moins de 5 jours d'écart (`ecartJours`), sinon le plus récent gagne. La version perdante est gardée dans `Conflits (téléphone)/` à la racine du vault.

Principe : `index.html` connecte à Microsoft, `graphfs.js` imite sur OneDrive la petite partie de l'API
File System Access dont la page se sert (inventaire par `/delta`, cache IndexedDB par eTag, écriture
avec `If-Match`, sauf les chemins interdits), puis `Vault Viewer.html` est lu dans OneDrive et ouvert dans un cadre.
La page du vault n'est pas modifiée : elle utilise sa porte `window.__testRootPromise`.

- `index.html`, `loader.js` : le chargeur
- `graphfs.js` : OneDrive vu comme un dossier
- `essai.html`, `essai.js` : page d'essai de connexion (diagnostic)
- `config.js` : réglages (identifiant client de l'application Microsoft)
- `lib/msal-browser.min.js` : MSAL.js 4.30.0 (Microsoft, licence MIT), copie fixe
- `sw.js`, `manifest.webmanifest`, `icons/` : installation sur l'écran d'accueil
