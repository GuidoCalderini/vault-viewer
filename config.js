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

  // Écriture depuis le téléphone (étape 3, 2026-11-06). Conflit avec l'ordinateur : fusion ligne par ligne,
  // et là où les deux ont touché les mêmes lignes, la version de OneDrive (l'ordinateur) gagne.
  ecriture: true,
  // Liste blanche (chemins relatifs au vault) : rien d'autre ne peut être écrit, même si « ecriture » est vrai.
  listeBlanche: [
    /^Langues\/[^/]+\/Vocabulaire - [^/]+\.md$/,
    /^Langues\/[^/]+\/Textes\/.+\.md$/,
    /^Langues\/[^/]+\/Leçons\/.+\.md$/,
    /^Langues\/[^/]+\/Évaluations\/.+\.md$/,
    /^Langues\/Journal - Langues\.md$/,
    /^Anki\/À ajouter dans Anki\.md$/
  ]
};
