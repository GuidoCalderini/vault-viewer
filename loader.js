/* Chargeur : connexion Microsoft, inventaire du vault (GraphFS), puis la page du vault
   (« Vault Viewer.html », lue dans OneDrive) est ouverte dans le cadre, branchée sur GraphFS
   par la porte prévue dans la page (window.__testRootPromise). Aucune donnée du vault n'est ici. */
(function () {
  "use strict";
  const cfg = window.VV_CONFIG;
  const $ = (id) => document.getElementById(id);
  const montre = (id, oui) => $(id).classList.toggle("cache", !oui);
  const rapport = { version: "chargeur-1", mode: (matchMedia("(display-mode: standalone)").matches || navigator.standalone) ? "PWA" : "navigateur", etapes: [] };
  const t0 = performance.now();
  function etat(t, d) {
    $("etat").textContent = t; $("detail").textContent = d || "";
    rapport.etapes.push(`${Math.round(performance.now() - t0)} ms — ${t}${d ? " (" + d + ")" : ""}`);
  }
  function echec(t, e) {
    etat(t, (e && (e.errorCode ? e.errorCode + " — " : "") + (e.message || String(e)) || "").slice(0, 300));
    $("accueil").style.display = ""; $("cadre").style.display = "none";
    montre("btn-reessayer", true); montre("lien-essai", true); montre("btn-deconnexion", true);
    rapport.stats = VV.fs && VV.fs.stats;
    $("rapport").textContent = "Pour Claude :\n" + JSON.stringify(rapport, null, 2); montre("rapport", true);
  }

  const redirectUri = new URL(".", location.href).href;
  const app = new msal.PublicClientApplication({
    auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri, navigateToLoginRequestUrl: true },
    cache: { cacheLocation: "localStorage" },
    system: { loggerOptions: { piiLoggingEnabled: false, logLevel: msal.LogLevel.Error, loggerCallback: () => {} } }
  });

  let reconnexion = false;
  async function getToken() {
    const compte = app.getActiveAccount() || app.getAllAccounts()[0];
    if (!compte) throw new Error("Pas connecté.");
    try { return (await app.acquireTokenSilent({ scopes: cfg.scopes, account: compte })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        if (!reconnexion) {
          reconnexion = true;
          $("bandeau-texte").textContent = "La connexion à OneDrive a expiré.";
          $("bandeau").style.display = "block";
        }
        const err = new Error("Connexion expirée : touche « Se reconnecter »."); err.reconnexion = true; throw err;
      }
      throw e;
    }
  }
  $("bandeau-btn").addEventListener("click", () => app.acquireTokenRedirect({ scopes: cfg.scopes, account: app.getActiveAccount() || undefined }));

  const VV = window.VV = {
    fs: new GraphFS({
      getToken, vaultFolders: cfg.vaultFolders, skipDirs: cfg.skipDirs,
      canWrite: (p) => !!cfg.ecriture && cfg.listeBlanche.some((re) => re.test(p))
    }),
    rapport
  };

  async function ouvrirPage() {
    const fs = VV.fs;
    const it = fs.child(fs.rootId, cfg.pageFile);
    if (!it) throw new Error(`« ${cfg.pageFile} » introuvable à la racine du vault.`);
    etat("Ouverture de la page…", cfg.pageFile);
    const html = await (await fs.content(it)).text();
    // La porte de la page : un dossier racine fourni d'avance remplace le sélecteur de dossier.
    const porte = "<script>window.__testRootPromise = Promise.resolve(parent.VV.fs.root); window.__VV_TELEPHONE = true;<\/script>";
    const avecPorte = /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, (m) => m + porte) : porte + html;
    const url = URL.createObjectURL(new Blob([avecPorte], { type: "text/html" }));
    const cadre = $("cadre");
    cadre.addEventListener("load", () => {
      $("accueil").style.display = "none"; cadre.style.display = "block";
      try { cadre.contentDocument && (document.title = cadre.contentDocument.title || "Vault"); } catch (e) {}
    }, { once: true });
    cadre.src = url + (cfg.startHash || "");
  }

  async function demarrer() {
    montre("btn-reessayer", false); montre("lien-essai", false); montre("btn-deconnexion", false); montre("rapport", false);
    try {
      etat("Inventaire du vault…", "");
      const r = await VV.fs.sync((d) => etat("Inventaire du vault…", d));
      rapport.inventaire = r;
      etat("Inventaire fait", `${r.elements} éléments, ${r.changements} changements`);
      await ouvrirPage();
    } catch (e) { echec("Le vault n'a pas pu s'ouvrir", e); }
  }

  // Quand on revient dans l'application : rafraîchir l'inventaire (seulement les changements).
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && VV.fs.rootId && Date.now() - VV.fs.lastSync > 60000) VV.fs.sync().catch(() => {});
  });

  $("btn-connexion").addEventListener("click", () => app.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" }));
  $("btn-reessayer").addEventListener("click", demarrer);
  $("btn-deconnexion").addEventListener("click", async () => {
    try { await VV.fs.signOutClear(); } catch (e) {}
    try { await app.clearCache(); } catch (e) {}
    try { localStorage.clear(); sessionStorage.clear(); } catch (e) {}
    location.replace(redirectUri);
  });

  (async () => {
    try {
      await app.initialize();
      const retour = await app.handleRedirectPromise();
      if (retour && retour.account) app.setActiveAccount(retour.account);
      const comptes = app.getAllAccounts();
      if (!comptes.length) { etat("Bienvenue", "Connecte-toi avec ton compte Microsoft (Hotmail)."); montre("btn-connexion", true); return; }
      app.setActiveAccount(app.getActiveAccount() || comptes[0]);
      await demarrer();
    } catch (e) { echec("Connexion impossible", e); montre("btn-connexion", true); }
  })();

  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
