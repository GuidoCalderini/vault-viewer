// Page d'essai de connexion (étape 0) — lecture seule.
(function () {
  "use strict";
  const cfg = window.VV_CONFIG;
  const $ = (id) => document.getElementById(id);
  const etapes = $("etapes");
  const rapport = {
    version: "essai-1",
    mode: (window.matchMedia("(display-mode: standalone)").matches || navigator.standalone) ? "PWA (écran d'accueil)" : "navigateur",
    agent: navigator.userAgent,
    etapes: []
  };

  function etape(titre, etat, detail) {
    const li = document.createElement("li");
    li.className = etat;
    li.textContent = (etat === "ok" ? "✓ " : etat === "erreur" ? "✗ " : "… ") + titre;
    if (detail) { const s = document.createElement("small"); s.textContent = detail; li.appendChild(s); }
    etapes.appendChild(li);
    rapport.etapes.push({ titre, etat, detail: detail || "" });
    $("rapport").textContent = JSON.stringify(rapport, null, 2);
    return li;
  }

  // Chemin OneDrive → segment d'URL Graph (apostrophes, accents, espaces).
  function enc(chemin) {
    return chemin.split("/").map((p) => encodeURIComponent(p).replace(/'/g, "%27")).join("/");
  }

  if (!cfg || !window.msal) { etape("Chargement de la page", "erreur", "config.js ou MSAL absent"); return; }
  etape("Mode d'ouverture", "ok", rapport.mode);
  if (!cfg.clientId || cfg.clientId === "A-REMPLIR") {
    etape("Identifiant client", "erreur", "Pas encore rempli dans config.js — étape Microsoft à faire.");
    return;
  }

  const redirectUri = new URL(".", location.href).href; // dossier de la page, avec « / » final
  const app = new msal.PublicClientApplication({
    auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri, navigateToLoginRequestUrl: true },
    cache: { cacheLocation: "localStorage" },
    system: { loggerOptions: { piiLoggingEnabled: false, logLevel: msal.LogLevel.Warning,
      loggerCallback: (lvl, msg) => { if (lvl <= msal.LogLevel.Warning) console.warn("[msal]", msg.replace(/[A-Za-z0-9-_]{40,}/g, "…")); } } }
  });

  async function jeton() {
    const compte = app.getActiveAccount() || app.getAllAccounts()[0];
    if (!compte) return null;
    app.setActiveAccount(compte);
    try {
      const r = await app.acquireTokenSilent({ scopes: cfg.scopes, account: compte });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        etape("Session expirée", "attente", "Nouvelle connexion par redirection…");
        await app.acquireTokenRedirect({ scopes: cfg.scopes, account: compte });
        return null;
      }
      throw e;
    }
  }

  async function graph(chemin, tok, brut) {
    const r = await fetch("https://graph.microsoft.com/v1.0" + chemin, { headers: { Authorization: "Bearer " + tok } });
    if (!r.ok) { const t = await r.text(); const err = new Error("HTTP " + r.status + " " + t.slice(0, 200)); err.status = r.status; throw err; }
    return brut ? r.text() : r.json();
  }

  async function test() {
    const t0 = performance.now();
    const tok = await jeton();
    if (!tok) { etape("Jeton", "erreur", "Pas de jeton (pas connecté)."); return; }
    etape("Jeton d'accès obtenu", "ok", "en " + Math.round(performance.now() - t0) + " ms");

    let racine = null, enfants = null;
    for (const nom of cfg.vaultFolders) {
      try {
        const t1 = performance.now();
        enfants = await graph("/me/drive/root:/" + enc(nom) + ":/children?$select=name,folder,file,size,eTag,lastModifiedDateTime&$top=200", tok);
        racine = nom;
        etape("Dossier « " + nom + " » trouvé", "ok", enfants.value.length + " éléments, " + Math.round(performance.now() - t1) + " ms");
        break;
      } catch (e) {
        if (e.status !== 404) { etape("Lecture du dossier « " + nom + " »", "erreur", e.message); return; }
      }
    }
    if (!racine) { etape("Dossier du vault", "erreur", "Introuvable : " + cfg.vaultFolders.join(" / ")); return; }
    $("liste").textContent = enfants.value
      .map((x) => (x.folder ? "📁 " : "📄 ") + x.name)
      .sort((a, b) => a.localeCompare(b, "fr")).join("\n");

    try {
      const t2 = performance.now();
      const meta = await graph("/me/drive/root:/" + enc(racine + "/" + cfg.testNote) + "?$select=name,eTag,size", tok);
      const texte = await graph("/me/drive/root:/" + enc(racine + "/" + cfg.testNote) + ":/content", tok, true);
      $("note").textContent = texte.slice(0, 800) + (texte.length > 800 ? "\n…" : "");
      etape("Note lue (accents compris)", "ok", meta.size + " octets, eTag présent : " + (meta.eTag ? "oui" : "non") + ", " + Math.round(performance.now() - t2) + " ms");
    } catch (e) {
      etape("Lecture de la note de test", "erreur", e.message);
      return;
    }
    etape("Tout fonctionne", "ok", rapport.mode === "navigateur"
      ? "Essaie maintenant depuis l'icône de l'écran d'accueil."
      : "Connexion réussie dans la PWA.");
  }

  function majBoutons() {
    const connecte = app.getAllAccounts().length > 0;
    $("btn-connexion").disabled = connecte;
    $("btn-test").disabled = !connecte;
    $("btn-deconnexion").disabled = !connecte;
  }

  $("btn-connexion").addEventListener("click", () => {
    etape("Redirection vers Microsoft", "attente", "");
    app.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" });
  });
  $("btn-test").addEventListener("click", () => { etapes.innerHTML = ""; rapport.etapes = []; etape("Mode d'ouverture", "ok", rapport.mode); test().catch((e) => etape("Erreur", "erreur", e.message)); });
  $("btn-deconnexion").addEventListener("click", async () => {
    const compte = app.getActiveAccount() || app.getAllAccounts()[0];
    // Déconnexion locale seulement (pas de détour par Microsoft).
    await app.clearCache({ account: compte });
    localStorage.clear(); sessionStorage.clear();
    etapes.innerHTML = ""; $("liste").textContent = "—"; $("note").textContent = "—";
    etape("Déconnecté", "ok", "Cache effacé sur ce téléphone.");
    majBoutons();
  });

  (async () => {
    try {
      await app.initialize();
      const retour = await app.handleRedirectPromise();
      if (retour) {
        app.setActiveAccount(retour.account);
        etape("Retour de Microsoft", "ok", "Revenu dans : " + rapport.mode);
      }
      majBoutons();
      if (app.getAllAccounts().length) await test();
      else etape("Pas connecté", "attente", "Touche « Se connecter ».");
    } catch (e) {
      etape("Connexion", "erreur", (e.errorCode ? e.errorCode + " — " : "") + (e.message || String(e)).slice(0, 300));
      majBoutons();
    }
  })();

  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
