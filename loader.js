/* Chargeur : connexion Microsoft, inventaire du vault (GraphFS), puis la page du vault
   (« Vault Viewer.html », lue dans OneDrive) est ouverte dans le cadre, branchée sur GraphFS
   par la porte prévue dans la page (window.__testRootPromise). Aucune donnée du vault n'est ici.
   Par défaut on se connecte. Si la connexion est impossible (pas de réseau, connexion expirée, trop lent),
   l'écran propose « Travailler hors ligne » : la page s'ouvre avec la version gardée dans le téléphone
   (adresse ?horsligne=1), les changements attendent la prochaine connexion. */
(function () {
  "use strict";
  const cfg = window.VV_CONFIG;
  const $ = (id) => document.getElementById(id);
  const montre = (id, oui) => $(id).classList.toggle("cache", !oui);
  const PARAMS = new URLSearchParams(location.search);
  const HORS_LIGNE = PARAMS.has("horsligne");
  // ?page=… : ouvrir une autre page du vault (ex. une sauvegarde) pour le diagnostic.
  const PAGE = PARAMS.get("page") || cfg.pageFile;
  const rapport = { version: "chargeur-7", horsLigne: HORS_LIGNE, mode: (matchMedia("(display-mode: standalone)").matches || navigator.standalone) ? "PWA" : "navigateur", etapes: [] };
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
    offrirHorsLigne();
  }
  const quand = (t) => new Date(t).toLocaleString("fr-CA", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
  // Bouton « Travailler hors ligne », seulement s'il y a une version gardée dans le téléphone.
  async function offrirHorsLigne() {
    if (HORS_LIGNE) return;
    const at = await VV.fs.peekSaved().catch(() => 0);
    if (!at) return;
    $("btn-horsligne").textContent = `Travailler hors ligne (version du ${quand(at)})`;
    montre("btn-horsligne", true);
  }
  // Écran de choix quand la connexion n'a pas marché : se connecter (par défaut) ou travailler hors ligne.
  function ecranChoix(titre, detail) {
    etat(titre, detail);
    $("accueil").style.display = ""; $("cadre").style.display = "none";
    montre("btn-connexion", true); montre("btn-reessayer", true);
    offrirHorsLigne();
  }
  const avecDelai = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => { const e = new Error("délai dépassé"); e.delai = true; rej(e); }, ms))]);

  const redirectUri = new URL(".", location.href).href;   // sans ?horsligne
  const app = new msal.PublicClientApplication({
    auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri, navigateToLoginRequestUrl: true },
    cache: { cacheLocation: "localStorage" },
    system: { loggerOptions: { piiLoggingEnabled: false, logLevel: msal.LogLevel.Error, loggerCallback: () => {} } }
  });

  let reconnexion = false, pageOuverte = false;
  async function getToken() {
    if (HORS_LIGNE) { const e = new Error("Mode hors ligne : gardé dans le téléphone."); e.horsLigne = true; throw e; }
    const compte = app.getActiveAccount() || app.getAllAccounts()[0];
    if (!compte) throw new Error("Pas connecté.");
    try { return (await app.acquireTokenSilent({ scopes: cfg.scopes, account: compte })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        if (!reconnexion && pageOuverte) {
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
      getToken, vaultFolders: cfg.vaultFolders, skipDirs: cfg.skipDirs, ecartJours: cfg.ecartJours,
      canWrite: (p) => { const n = String(p).normalize("NFC"); return !!cfg.ecriture && !(cfg.interdits || []).some((re) => re.test(n)); },
      onStatus: ({ attente }) => montrerAttente(attente),
      onConflict: ({ nom, conflits, gagnant, copie, dossier }) => {
        const pl = conflits > 1 ? "s" : "";
        const garde = copie ? ` L'autre version est gardée dans « ${dossier} ».` : "";
        dire(gagnant === "téléphone"
          ? `« ${nom} » avait aussi changé sur l'ordinateur, il y a longtemps : la version du téléphone a gagné (${conflits} passage${pl}).${garde}`
          : `« ${nom} » avait aussi changé sur l'ordinateur : sa version a gagné (${conflits} passage${pl}).${garde}`, true);
      }
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
    else if (!attenteAvant && n && pageOuverte) dire(HORS_LIGNE
      ? "Hors ligne : le changement est gardé dans le téléphone et partira à la prochaine connexion."
      : "Pas de connexion : le changement est gardé dans le téléphone et partira tout seul.");
    attenteAvant = n;
  }
  $("attente").addEventListener("click", () => { if (HORS_LIGNE) seConnecterDepuisHorsLigne(); else VV.fs.flush().catch(() => {}); });

  async function ouvrirPage() {
    const fs = VV.fs;
    const it = fs.child(fs.rootId, PAGE);
    if (!it) throw new Error(`« ${PAGE} » introuvable à la racine du vault.`);
    const quoi = PAGE + (it.size ? ` · ${Math.round(it.size/1024)} Ko` : "");
    etat("Ouverture de la page…", quoi);
    fs.onEssai = (d) => { $("etat").textContent = "Ouverture de la page… (Microsoft est lent, on insiste)"; $("detail").textContent = quoi + " · " + d; };
    const html = await (await fs.content(it)).text();
    fs.onEssai = null;
    // La porte de la page : un dossier racine fourni d'avance remplace le sélecteur de dossier.
    const porte = "<script>window.__testRootPromise = Promise.resolve(parent.VV.fs.root); window.__VV_TELEPHONE = true;<\/script>";
    const avecPorte = /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, (m) => m + porte) : porte + html;
    const url = URL.createObjectURL(new Blob([avecPorte], { type: "text/html" }));
    const cadre = $("cadre");
    cadre.addEventListener("load", () => {
      $("accueil").style.display = "none"; cadre.style.display = "block"; pageOuverte = true;
      try { cadre.contentDocument && (document.title = cadre.contentDocument.title || "Vault"); } catch (e) {}
    }, { once: true });
    cadre.src = url + (cfg.startHash || "");
  }

  function cacherBoutons() {
    for (const id of ["btn-connexion", "btn-reessayer", "lien-essai", "btn-deconnexion", "rapport", "btn-horsligne"]) montre(id, false);
  }
  // Erreurs MSAL qui veulent dire « pas de réseau » plutôt que « reconnecte-toi ».
  const RESEAU = /no_network_connectivity|post_request_failed|get_request_failed|endpoints_resolution_error|network|timed_out|failed to fetch|load failed/i;

  async function demarrer() {
    cacherBoutons();
    // Si ça traîne (réseau faible), on propose déjà le hors ligne.
    const lent = setTimeout(offrirHorsLigne, 5000);
    try {
      etat("Connexion à OneDrive…", "");
      if (navigator.onLine === false) { clearTimeout(lent); return ecranChoix("Pas de réseau", "Impossible de joindre OneDrive pour l'instant."); }
      try {
        const compte = app.getActiveAccount() || app.getAllAccounts()[0];
        await avecDelai(app.acquireTokenSilent({ scopes: cfg.scopes, account: compte }), 15000);
      } catch (e) {
        clearTimeout(lent); rapport.connexion = (e.errorCode || "") + " " + (e.message || e);
        if (e.delai || RESEAU.test((e.errorCode || "") + " " + (e.message || ""))) return ecranChoix("Impossible de joindre OneDrive", "Le réseau ne répond pas.");
        return ecranChoix("Connexion expirée", "Reconnecte-toi, ou travaille hors ligne.");
      }
      etat("Inventaire du vault…", "");
      const r = await VV.fs.sync((d) => etat("Inventaire du vault…", d));
      rapport.inventaire = r;
      etat("Inventaire fait", `${r.elements} éléments, ${r.changements} changements`);
      await ouvrirPage();
      clearTimeout(lent); montre("btn-horsligne", false);
      garderDAvance();
    } catch (e) { clearTimeout(lent); echec("Le vault n'a pas pu s'ouvrir", e); }
  }

  // Après une synchro : toutes les notes sont gardées dans le téléphone pour le hors ligne (en arrière-plan).
  function garderDAvance() {
    try { navigator.storage && navigator.storage.persist && navigator.storage.persist().then((ok) => { rapport.stockagePersistant = ok; }); } catch (e) {}
    setTimeout(() => { VV.fs.prefetch().then((r) => { if (r) rapport.gardeDAvance = r; }).catch(() => {}); }, 4000);
  }

  // Mode hors ligne : la version gardée dans le téléphone, sans connexion.
  async function demarrerHorsLigne() {
    cacherBoutons();
    try {
      etat("Ouverture hors ligne…", "");
      const at = await VV.fs.openOffline();
      if (at == null) {
        etat("Rien n'est encore gardé dans le téléphone", "Connecte-toi une première fois.");
        montre("btn-connexion", true); return;
      }
      rapport.versionDu = at && new Date(at).toISOString();
      await ouvrirPage();
      const p = $("pastille-hl");
      p.textContent = `📴 Hors ligne · ${at ? quand(at) : "?"}`; p.style.display = "block";
    } catch (e) { echec("La version hors ligne n'a pas pu s'ouvrir", e); montre("btn-connexion", true); }
  }
  function seConnecterDepuisHorsLigne() {
    const n = VV.fs.pending.size;
    const msg = (n ? `${n} changement(s) gardé(s) dans le téléphone partiront à la connexion.\n` : "") + "Se connecter à OneDrive maintenant?";
    if (confirm(msg)) location.replace(redirectUri);
  }
  $("pastille-hl").addEventListener("click", seConnecterDepuisHorsLigne);
  $("btn-horsligne").addEventListener("click", () => location.replace(redirectUri + "?horsligne=1"));

  // Quand on revient dans l'application : envoyer ce qui attend, rafraîchir l'inventaire (seulement les changements),
  // et si quelque chose a changé ailleurs, prévenir la page (une leçon ouverte se met à jour).
  let syncEnCours = null;
  async function rafraichir() {
    if (HORS_LIGNE || !VV.fs.rootId || syncEnCours) return syncEnCours;
    syncEnCours = (async () => {
      try {
        await VV.fs.flush();
        const r = await VV.fs.sync();
        if (r && r.changements) { VV.fs.prefetch().catch(() => {}); try { $("cadre").contentWindow.__vvRefresh && $("cadre").contentWindow.__vvRefresh(); } catch (e) {} }
      } catch (e) {} finally { syncEnCours = null; }
    })();
    return syncEnCours;
  }
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && Date.now() - VV.fs.lastSync > 10000) rafraichir();
  });
  addEventListener("online", () => { if (!HORS_LIGNE) VV.fs.flush().catch(() => {}); });
  // Pendant qu'on lit : un coup d'œil aux changements toutes les 2 minutes (seulement les différences).
  setInterval(() => { if (document.visibilityState === "visible" && Date.now() - VV.fs.lastSync > 110000) rafraichir(); }, 30000);

  // Se connecter : avec le compte déjà connu, Microsoft le reconnaît d'habitude en un toucher.
  $("btn-connexion").addEventListener("click", () => {
    if (HORS_LIGNE) return location.replace(redirectUri);
    const compte = app.getActiveAccount() || app.getAllAccounts()[0];
    if (compte) app.acquireTokenRedirect({ scopes: cfg.scopes, account: compte, loginHint: compte.username });
    else app.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" });
  });
  $("btn-reessayer").addEventListener("click", () => { if (HORS_LIGNE) location.replace(redirectUri); else demarrer(); });
  $("btn-deconnexion").addEventListener("click", async () => {
    const n = VV.fs.pending.size;
    if (n && !confirm(`${n} changement(s) pas encore envoyé(s) à OneDrive seront perdus. Se déconnecter quand même?`)) return;
    try { await VV.fs.signOutClear(true); } catch (e) {}
    try { await app.clearCache(); } catch (e) {}
    try { localStorage.clear(); sessionStorage.clear(); } catch (e) {}
    location.replace(redirectUri);
  });

  (async () => {
    if (HORS_LIGNE) return demarrerHorsLigne();
    try {
      await app.initialize();
      const retour = await app.handleRedirectPromise();
      if (retour && retour.account) app.setActiveAccount(retour.account);
      const comptes = app.getAllAccounts();
      if (!comptes.length) { etat("Bienvenue", "Connecte-toi avec ton compte Microsoft (Hotmail)."); montre("btn-connexion", true); offrirHorsLigne(); return; }
      app.setActiveAccount(app.getActiveAccount() || comptes[0]);
      await demarrer();
    } catch (e) { echec("Connexion impossible", e); montre("btn-connexion", true); }
  })();

  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
