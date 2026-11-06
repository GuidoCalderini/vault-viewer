// Réglages de la page d'essai. L'identifiant client n'est pas un secret.
// À remplir après l'inscription de l'application dans Microsoft Entra.
window.VV_CONFIG = {
  clientId: "A-REMPLIR",
  // Comptes Microsoft personnels seulement.
  authority: "https://login.microsoftonline.com/consumers",
  scopes: ["User.Read", "Files.ReadWrite"],
  // Dossier du vault dans OneDrive (premier trouvé).
  vaultFolders: ["Guido's Vault", "Professor Claude"],
  // Note lue pour le test de lecture (à la racine du vault).
  testNote: "00 - Carte du vault (pour Claude).md"
};
