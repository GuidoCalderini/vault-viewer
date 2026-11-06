/* Chargeur : connexion Microsoft, inventaire du vault (GraphFS), puis la page du vault
   (« Vault Viewer.html », lue dans OneDrive) est ouverte dans le cadre, branchée sur GraphFS
   par la porte prévue dans la page (window.__testRootPromise). Aucune donnée du vault n'est ici. */
(function () {
  "use strict";
  const cfg = window.VV_CONFIG;
  const $ = (id) => document.getElementById(id);
  const montre = (id, oui) => $(id).classList.toggle("cache", !oui);
  const rapport = { version: "chargeur-2", mode: (matchMedia("(display-mode: standalone)").matches || navigator.standalone) ? "PWA" : "navigateur", etapes: [] };
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
      canWrite: (p) => { const n = String(p).normalize("NFC"); return !!cfg.ecriture && !(cfg.interdits || []).some((re) => re.test(n)); },
      onStatus: ({ attente }) => montrerAttente(attente),
      onConflict: ({ nom, conflits }) => dire(`« ${nom} » avait changé sur l'ordinateur : sa version a gagné (${conflits} passage${conflits > 1 ? "s" : ""}).`, true)
    }),
    rapport
  };

  // Message dans la page du vault (sa fonction toast), sinon dans le bandeau du chargeur.
  function dire(msg, erreur) {
    try { const w = $("cadre").contentWindow; if (w && typeof w.toast === "function") { w.toast(msg, erreur ? { error: true } : undefined); return; } } catch (e) {}
    $("bandeau-texte").textContent = msg; $("bandeau-btn").style.display = "none"; $("bandeau").style.display = "block";
    setTimeout(() => { $("bandeau").style.display = "none"; $("bandeau-btn").style.display = ""; }, 6000);
  }
  // Pastille « n en attente » (changements gardés dans le téléphone, pas encore dans OneDrive).
  let attenteAvant = 0;
  function montrerAttente(n) {
    const p = $("attente"); if (!p) return;
    p.textContent = `⏳ ${n} en attente`; p.style.display = n ? "block" : "none";
    if (attenteAvant && !n) dire("Changements en attente envoyés à OneDrive ✓");
    else if (!attenteAvant && n) dire("Pas de connexion : le changement est gardé dans le téléphone et partira tout seul.");
    attenteAvant = n;
  }
  $("attente").addEventListener("click", () => { VV.fs.flush().catch(() => {}); });

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

  // Quand on revient dans l'application : envoyer ce qui attend, rafraîchir l'inventaire (seulement les changements),
  // et si quelque chose a changé ailleurs, prévenir la page (une leçon ouverte se met à jour).
  let syncEnCours = null;
  async function rafraichir() {
    if (!VV.fs.rootId || syncEnCours) return syncEnCours;
    syncEnCours = (async () => {
      try {
        await VV.fs.flush();
        const r = await VV.fs.sync();
        if (r && r.changements) { try { $("cadre").contentWindow.__vvRefresh && $("cadre").contentWindow.__vvRefresh(); } catch (e) {} }
      } catch (e) {} finally { syncEnCours = null; }
    })();
    return syncEnCours;
  }
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && Date.now() - VV.fs.lastSync > 10000) rafraichir();
  });
  addEventListener("online", () => { VV.fs.flush().catch(() => {}); });
  // Pendant qu'on lit : un coup d'œil aux changements toutes les 2 minutes (seulement les différences).
  setInterval(() => { if (document.visibilityState === "visible" && Date.now() - VV.fs.lastSync > 110000) rafraichir(); }, 30000);

  $("btn-connexion").addEventListener("click", () => app.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" }));
  $("btn-reessayer").addEventListener("click", demarrer);
  $("btn-deconnexion").addEventListener("click", async () => {
    const n = VV.fs.pending.size;
    if (n && !confirm(`${n} changement(s) pas encore envoyé(s) à OneDrive seront perdus. Se déconnecter quand même?`)) return;
    try { await VV.fs.signOutClear(true); } catch (e) {}
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
