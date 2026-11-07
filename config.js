// Réglages du chargeur. L'identifiant client n'est pas un secret.
window.VV_CONFIG = {
  clientId: "2e0247c8-11d9-468b-ae12-75c85f175c7d",
  // Comptes Microsoft personnels seulement.
  authority: "https://login.microsoftonline.com/consumers",
  scopes: ["User.Read", "Files.ReadWrite"],
  // Dossier du vault dans OneDrive (premier trouvé).
  vaultFolders: ["Guido's Vault", "Professor Claude"],
  // La page du vault, à la racine du dossier.
  pageFile: "Vault Viewer.html",
  // Dossiers jamais inventoriés d'avance (chargés à la demande seulement).
  skipDirs: [".obsidian", ".trash", "Calibre Library", ".git", "node_modules"],
  // Page ouverte au lancement.
  startHash: "#/langues",
  // Note lue par la page d'essai (essai.html).
  testNote: "00 - Carte du vault (pour Claude).md",

  // Écriture depuis le téléphone (étape 3, 2026-11-06).
  ecriture: true,
  // Conflit (2026-11-07, décision de Guido) : si seul le téléphone a changé un fichier depuis la dernière
  // synchro, sa version est gardée. Si les deux ont changé : fusion ligne par ligne ; là où les deux ont
  // touché les mêmes lignes, l'ordinateur gagne si les deux changements ont moins de ecartJours d'écart,
  // sinon le plus récent gagne. La version perdante est gardée dans « Conflits (téléphone)/ ».
  ecartJours: 5,
  // Depuis le 2026-11-06 (décision de Guido) : le téléphone peut écrire partout dans le vault,
  // sauf ici (chemins relatifs au vault). Rien hors du dossier du vault n'est jamais touché.
  interdits: [
    /^Vault Viewer[^/]*\.html$/i,   // la page elle-même : un bogue la casserait sur les deux appareils
    /^\.obsidian(\/|$)/,             // réglages d'Obsidian
    /^Calibre Library(\/|$)/        // la bibliothèque Calibre (gérée par Calibre)
  ]
};
