/* GraphFS — imitation, sur OneDrive (Microsoft Graph), de la petite partie de l'API
   « File System Access » dont la page du vault se sert :
     dossier : name, kind, entries(), values(), keys(), getDirectoryHandle(n,{create}), getFileHandle(n,{create})
     fichier : name, kind, getFile() → File (text(), lastModified, type), createWritable() → write(), close()
   L'inventaire vient de /delta (un par dossier de premier niveau), gardé dans IndexedDB : au lancement
   suivant, seuls les changements sont téléchargés. Le contenu des fichiers est mis en cache par eTag.
   Écriture : PUT avec If-Match, sauf les chemins interdits. Si seul le téléphone a changé le fichier,
   sa version est gardée. Si le fichier a aussi changé ailleurs (412) : fusion ligne par ligne ; là où
   les deux ont touché les mêmes lignes, l'ordinateur gagne si les deux changements ont moins de 5 jours
   d'écart, sinon le plus récent gagne ; la version perdante est gardée dans « Conflits (téléphone)/ ».
   Sans réseau, connexion expirée ou mode hors ligne : file d'attente (IndexedDB), envoyée dès que
   possible avec la même règle. Hors ligne, tout est lu dans le cache (inventaire + contenu). */
(function () {
  "use strict";
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SEL = "id,name,parentReference,file,folder,deleted,eTag,size,lastModifiedDateTime,fileSystemInfo";
  const CACHE_MAX = 80 * 1024 * 1024;  // on ne garde pas en cache les fichiers plus gros que ça
  const GROS = 4 * 1024 * 1024;        // au-delà (livres, PDF) : seulement les GROS_N derniers ouverts
  const GROS_N = 3;
  const PREFETCH_MAX = 1.5 * 1024 * 1024; // gardés d'avance pour le hors ligne : textes jusqu'à cette taille
  const PREFETCH_EXT = /\.(md|txt|json|csv|tsv|canvas|base|ya?ml|css|svg)$/i;
  const JOUR = 86400000;
  const DOSSIER_CONFLITS = "Conflits (téléphone)";
  const TREE_VERSION = 1;

  /* ---------- IndexedDB ---------- */
  function idbOpen() {
    return new Promise((res, rej) => {
      const r = indexedDB.open("vv-graphfs", 2);
      r.onupgradeneeded = () => { const db = r.result; for (const n of ["kv", "content", "queue"]) if (!db.objectStoreNames.contains(n)) db.createObjectStore(n); };
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  }
  let dbp = null;
  function idb(store, mode, fn) {
    dbp ||= idbOpen();
    return dbp.then((db) => new Promise((res, rej) => {
      const tx = db.transaction(store, mode); const st = tx.objectStore(store);
      const r = fn(st); tx.oncomplete = () => res(r && r.result); tx.onerror = () => rej(tx.error); tx.onabort = () => rej(tx.error);
    }));
  }
  const idbGet = (store, k) => idb(store, "readonly", (s) => s.get(k));
  const idbPut = (store, k, v) => idb(store, "readwrite", (s) => s.put(v, k));
  const idbClear = (store) => idb(store, "readwrite", (s) => s.clear());
  const idbDel = (store, k) => idb(store, "readwrite", (s) => s.delete(k));
  const idbAll = (store) => idb(store, "readonly", (s) => s.getAll());

  /* ---------- utilitaires ---------- */
  const nfc = (s) => (s || "").normalize("NFC");
  const enc = (s) => encodeURIComponent(s).replace(/'/g, "%27");
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  function notFound(name) { const e = new Error(`« ${name} » introuvable`); e.name = "NotFoundError"; return e; }

  // Limiteur : au plus N requêtes à la fois.
  function limiter(n) {
    let active = 0; const q = [];
    const next = () => { if (active >= n || !q.length) return; active++; const [fn, res, rej] = q.shift(); fn().then(res, rej).finally(() => { active--; next(); }); };
    return (fn) => new Promise((res, rej) => { q.push([fn, res, rej]); next(); });
  }

  // Erreur de réseau ou connexion expirée → on garde le changement pour plus tard.
  const horsLigne = (e) => !!e && (e.horsLigne || e.reconnexion || e.name === "TypeError" || /failed to fetch|load failed|network/i.test(e.message || "") || navigator.onLine === false);

  /* ---------- fusion à trois, ligne par ligne ---------- */
  // Pour chaque ligne de a : l'indice de la ligne appariée dans b (plus longue sous-suite commune), ou -1.
  function apparier(a, b) {
    const n = a.length, m = b.length, map = new Int32Array(n).fill(-1);
    let s = 0; while (s < n && s < m && a[s] === b[s]) { map[s] = s; s++; }
    let ea = n, eb = m; while (ea > s && eb > s && a[ea - 1] === b[eb - 1]) { ea--; eb--; map[ea] = eb; }
    const N = ea - s, M = eb - s;
    if (!N || !M || N * M > 1.2e7) return map;          // milieu trop gros : traité comme un seul bloc
    const W = M + 1, L = new Uint16Array((N + 1) * W);
    for (let i = N - 1; i >= 0; i--) for (let j = M - 1; j >= 0; j--)
      L[i * W + j] = a[s + i] === b[s + j] ? L[(i + 1) * W + j + 1] + 1 : Math.max(L[(i + 1) * W + j], L[i * W + j + 1]);
    let i = 0, j = 0;
    while (i < N && j < M) {
      if (a[s + i] === b[s + j]) { map[s + i] = s + j; i++; j++; }
      else if (L[(i + 1) * W + j] >= L[i * W + j + 1]) i++; else j++;
    }
    return map;
  }
  // base = ce que le téléphone avait lu ; moi = ce que le téléphone veut écrire ; eux = OneDrive maintenant.
  // Bloc touché des deux côtés de façon différente → « eux » gagne (l'ordinateur), sauf gagnant = "moi".
  function fusion3(baseTxt, moiTxt, euxTxt, gagnant) {
    const moiGagne = gagnant === "moi";
    const o = baseTxt.split("\n"), m = moiTxt.split("\n"), t = euxTxt.split("\n");
    const mM = apparier(o, m), mT = apparier(o, t);
    const egal = (x, y) => x.length === y.length && x.every((v, q) => v === y[q]);
    let conflits = 0; const out = [];
    const choisir = (O, M, T) => {
      if (egal(M, O)) return T; if (egal(T, O)) return M; if (egal(M, T)) return M;
      // Lignes modifiées sur place des deux côtés (ex. statuts de mots voisins) : on règle ligne par ligne.
      if (O.length === M.length && O.length === T.length)
        return O.map((x, q) => (M[q] === x ? T[q] : T[q] === x || T[q] === M[q] ? M[q] : (conflits++, moiGagne ? M[q] : T[q])));
      conflits++; return moiGagne ? M : T;
    };
    let i = 0, j = 0, k = 0;
    for (let a = 0; a <= o.length; a++) {
      if (a < o.length && (mM[a] < 0 || mT[a] < 0)) continue;
      const ja = a < o.length ? mM[a] : m.length, ka = a < o.length ? mT[a] : t.length;
      if (ja < j || ka < k) continue;
      out.push(...choisir(o.slice(i, a), m.slice(j, ja), t.slice(k, ka)));
      if (a < o.length) out.push(o[a]);
      i = a + 1; j = ja + 1; k = ka + 1;
    }
    return { texte: out.join("\n"), conflits };
  }

  class GraphFS {
    constructor({ getToken, vaultFolders, skipDirs, canWrite, log, onStatus, onConflict, ecartJours }) {
      this.getToken = getToken;
      this.vaultFolders = vaultFolders;
      this.skip = new Set(skipDirs || []);
      this.canWrite = canWrite || (() => false);
      this.log = log || (() => {});
      this.items = new Map();          // id -> {id,name,parent,dir,eTag,size,mtime,mime,kids:Set|null}
      this.deltaLinks = {};            // id dossier de 1er niveau -> deltaLink
      this.rootId = null; this.rootName = null;
      this.lastSync = 0;
      this.limit = limiter(6);
      this.stats = { requetes: 0, telecharges: 0, depuisCache: 0 };
      this.onStatus = onStatus || (() => {});      // ({attente}) : nombre de changements pas encore envoyés
      this.onConflict = onConflict || (() => {});  // ({chemin, nom, conflits}) : l'ordinateur a gagné quelque part
      this.pending = new Map();                    // id (réel ou « local:… ») -> changement en attente
      this.alias = new Map();                      // id « local:… » -> vrai id, une fois le nouveau fichier créé
      this.queueLoaded = false; this.flushing = null; this.flushTimer = null;
      this.ecart = (ecartJours == null ? 5 : ecartJours) * JOUR;  // « à peu près en même temps »
      this.offline = false;          // mode hors ligne choisi : on ne lit que le cache
      this.syncAt = 0;               // moment de la dernière synchronisation complète gardée
      this.grosChain = Promise.resolve();
    }
    real(id) { let x = id; for (let n = 0; x && this.alias.has(x) && n < 5; n++) x = this.alias.get(x); return x; }

    /* ---- requêtes Graph ---- */
    async req(url, opts = {}, tries = 0) {
      const tok = await this.getToken();
      const headers = Object.assign({ Authorization: "Bearer " + tok }, opts.headers || {});
      this.stats.requetes++;
      const r = await this.limit(() => fetch(url.startsWith("http") ? url : GRAPH + url, Object.assign({}, opts, { headers })));
      if ((r.status === 429 || r.status === 503 || r.status === 504) && tries < 4) {
        const wait = (parseInt(r.headers.get("Retry-After") || "0", 10) || 2 ** tries) * 1000;
        await sleep(Math.min(wait, 30000)); return this.req(url, opts, tries + 1);
      }
      return r;
    }
    async json(url, opts) {
      const r = await this.req(url, opts);
      if (!r.ok) { const e = new Error(`Graph ${r.status} : ${(await r.text()).slice(0, 200)}`); e.status = r.status; throw e; }
      return r.json();
    }

    /* ---- arbre ---- */
    put(raw, parentOverride) {
      const id = raw.id; const old = this.items.get(id);
      const parent = parentOverride !== undefined ? parentOverride : (raw.parentReference && raw.parentReference.id) || null;
      const it = {
        id, name: nfc(raw.name), parent, dir: !!raw.folder,
        eTag: raw.eTag || null, size: raw.size || 0,
        mtime: raw.lastModifiedDateTime ? Date.parse(raw.lastModifiedDateTime) : 0,
        mime: (raw.file && raw.file.mimeType) || "",
        kids: old ? old.kids : (raw.folder ? null : undefined)
      };
      if (old && old.parent !== it.parent) this.items.get(old.parent)?.kids?.delete(id);
      this.items.set(id, it);
      const p = this.items.get(it.parent); if (p && p.kids) p.kids.add(id);
      return it;
    }
    drop(id) {
      const it = this.items.get(id); if (!it) return;
      if (it.kids) for (const k of [...it.kids]) this.drop(k);
      this.items.get(it.parent)?.kids?.delete(id);
      this.items.delete(id);
    }
    child(dirId, name) {
      const d = this.items.get(dirId); if (!d || !d.kids) return null;
      const n = nfc(name), low = n.toLowerCase(); let ci = null;
      for (const k of d.kids) { const c = this.items.get(k); if (!c) continue; if (c.name === n) return c; if (!ci && c.name.toLowerCase() === low) ci = c; }
      return ci;
    }
    pathOf(id) {
      const parts = []; let it = this.items.get(id);
      while (it && it.id !== this.rootId) { parts.unshift(it.name); it = this.items.get(it.parent); }
      return parts.join("/");
    }

    // Enfants d'un dossier pas encore chargé (dossiers sautés comme Calibre Library) : à la demande.
    async loadKids(dirId) {
      const d = this.items.get(dirId);
      if (!d || d.kids) return;
      const kids = new Set(); d.kids = kids;
      let url = `/me/drive/items/${dirId}/children?$select=${SEL}&$top=999`;
      try {
        while (url) {
          const j = await this.json(url);
          for (const raw of j.value) { const it = this.put(raw, dirId); if (it.dir && !this.items.get(it.id).kids) it.kids = null; }
          url = j["@odata.nextLink"];
        }
      } catch (e) { d.kids = null; throw e; }
    }

    // Dossier de premier niveau qui contient cet élément (null pour la racine).
    topOf(id) {
      let it = this.items.get(id);
      while (it && it.parent && it.parent !== this.rootId) it = this.items.get(it.parent);
      return it && it.id !== this.rootId ? it : null;
    }
    async save() {
      // On ne garde que ce que /delta tient à jour : la racine, ses enfants directs et les dossiers non sautés.
      const keep = (it) => {
        if (it.id === this.rootId || it.parent === this.rootId) return true;
        const t = this.topOf(it.id); return !!t && !this.skip.has(t.name) && !t.name.startsWith(".");
      };
      const items = [...this.items.values()].filter(keep).map(({ kids, ...rest }) => {
        const t = rest.parent === this.rootId && rest.dir && (this.skip.has(rest.name) || rest.name.startsWith("."));
        return Object.assign(rest, { loaded: !!kids && !t });
      });
      const syncAt = this.lastSync || this.syncAt || 0;
      await idbPut("kv", "tree", { v: TREE_VERSION, rootId: this.rootId, rootName: this.rootName, deltaLinks: this.deltaLinks, items, at: Date.now(), syncAt });
      await idbPut("kv", "syncAt", syncAt);
    }
    async restore() {
      const t = await idbGet("kv", "tree");
      if (!t || t.v !== TREE_VERSION) return false;
      this.rootId = t.rootId; this.rootName = t.rootName; this.deltaLinks = t.deltaLinks || {};
      this.syncAt = t.syncAt || t.at || 0;
      for (const it of t.items) this.items.set(it.id, Object.assign(it, { kids: it.dir ? (it.loaded ? new Set() : null) : undefined }));
      for (const it of this.items.values()) { delete it.loaded; const p = this.items.get(it.parent); if (p && p.kids) p.kids.add(it.id); }
      return true;
    }

    // Une passe /delta pour un dossier de premier niveau (complète la 1re fois, puis seulement les changements).
    async deltaFolder(folderId) {
      let url = this.deltaLinks[folderId] || `/me/drive/items/${folderId}/delta?$select=${SEL}`;
      const upd = [], del = [];
      while (url) {
        let j;
        try { j = await this.json(url); }
        catch (e) {
          if (e.status === 410 && this.deltaLinks[folderId]) { // jeton expiré : on recommence ce dossier
            delete this.deltaLinks[folderId];
            const f = this.items.get(folderId); if (f && f.kids) for (const k of [...f.kids]) this.drop(k);
            return this.deltaFolder(folderId);
          }
          throw e;
        }
        for (const raw of j.value) (raw.deleted ? del : upd).push(raw);
        url = j["@odata.nextLink"];
        if (j["@odata.deltaLink"]) this.deltaLinks[folderId] = j["@odata.deltaLink"];
      }
      // Les dossiers d'abord, pour que les fichiers trouvent leur parent.
      upd.sort((a, b) => (b.folder ? 1 : 0) - (a.folder ? 1 : 0));
      const seen = new Set();
      for (const raw of upd) {
        if (raw.id === folderId) { const f = this.items.get(folderId); if (f) { f.eTag = raw.eTag; f.mtime = Date.parse(raw.lastModifiedDateTime || 0) || f.mtime; } continue; }
        const it = this.put(raw); if (it.dir && !it.kids) it.kids = new Set(); seen.add(raw.id);
      }
      // Au cas où un dossier serait arrivé après son contenu : on refait les liens.
      for (const raw of upd) { const it = this.items.get(raw.id); const p = it && this.items.get(it.parent); if (p && p.kids) p.kids.add(it.id); }
      for (const raw of del) if (!seen.has(raw.id)) this.drop(raw.id);
      return upd.length + del.length;
    }

    // Date de la dernière synchronisation gardée dans le téléphone (0 si rien), sans tout charger.
    async peekSaved() {
      const a = await idbGet("kv", "syncAt").catch(() => null);
      if (a) return a;
      const t = await idbGet("kv", "tree").catch(() => null);
      return (t && t.v === TREE_VERSION && (t.syncAt || t.at)) || 0;
    }
    // Mode hors ligne : inventaire et contenus pris dans le téléphone seulement. null si rien n'est gardé.
    async openOffline() {
      this.offline = true;
      await this.loadQueue();
      if (!this.rootId && !(await this.restore())) return null;
      return this.syncAt || 0;
    }

    // Lancement : cache, puis racine (1 requête), puis /delta de chaque dossier de 1er niveau.
    async sync(progress) {
      progress = progress || (() => {});
      await this.loadQueue();
      if (!this.rootId && !(await this.restore())) {
        let found = null;
        for (const n of this.vaultFolders) {
          try { found = await this.json(`/me/drive/root:/${enc(n)}?$select=${SEL}`); break; } catch (e) { if (e.status !== 404) throw e; }
        }
        if (!found) throw new Error("Dossier du vault introuvable dans OneDrive : " + this.vaultFolders.join(" / "));
        this.rootId = found.id; this.rootName = nfc(found.name);
        this.put(found, null).kids = new Set();
      }
      const root = this.items.get(this.rootId);
      // Racine : liste complète à chaque fois (peu d'éléments).
      const fresh = new Set();
      let url = `/me/drive/items/${this.rootId}/children?$select=${SEL}&$top=999`;
      while (url) { const j = await this.json(url); for (const raw of j.value) { fresh.add(raw.id); this.put(raw, this.rootId); } url = j["@odata.nextLink"]; }
      for (const k of [...root.kids]) if (!fresh.has(k)) { delete this.deltaLinks[k]; this.drop(k); }
      const tops = [...root.kids].map((k) => this.items.get(k)).filter((x) => x && x.dir && !this.skip.has(x.name) && !x.name.startsWith("."));
      let done = 0, changes = 0;
      progress(`Inventaire : 0 / ${tops.length} dossiers`);
      await Promise.all(tops.map(async (t) => {
        if (!t.kids) t.kids = new Set();
        changes += await this.deltaFolder(t.id);
        progress(`Inventaire : ${++done} / ${tops.length} dossiers`);
      }));
      for (const k of Object.keys(this.deltaLinks)) if (!this.items.has(k)) delete this.deltaLinks[k];
      this.lastSync = Date.now();
      await this.save();
      if (this.pending.size) this.flushSoon(300);
      return { elements: this.items.size, changements: changes };
    }

    /* ---- contenu ---- */
    async content(it) {
      const p = this.pending.get(it.id);
      if (p) return new Blob([p.text], { type: "text/plain;charset=utf-8" });   // changement pas encore envoyé : c'est lui qu'on montre
      const c = await idbGet("content", it.id).catch(() => null);
      if (c && c.eTag && c.eTag === it.eTag) {
        this.stats.depuisCache++;
        if (c.blob && c.blob.size > GROS) this.garderGros(it.id);
        return c.blob;
      }
      if (this.offline) {
        // Hors ligne : une version plus ancienne vaut mieux que rien.
        if (c && c.blob) { this.stats.depuisCache++; return c.blob; }
        const e = new Error(`« ${it.name} » n'est pas gardé dans le téléphone (mode hors ligne).`); e.horsLigne = true; throw e;
      }
      const r = await this.req(`/me/drive/items/${it.id}/content`);
      if (!r.ok) { const e = new Error(`Lecture impossible (${r.status}) : ${it.name}`); e.status = r.status; throw e; }
      const blob = await r.blob(); this.stats.telecharges++;
      if (blob.size <= CACHE_MAX) {
        await idbPut("content", it.id, { eTag: it.eTag, blob }).catch(() => {});
        if (blob.size > GROS) this.garderGros(it.id);
      }
      return blob;
    }
    // Gros fichiers (livres) : on ne garde que les GROS_N derniers ouverts, pour ménager la place du téléphone.
    garderGros(id) {
      this.grosChain = this.grosChain.then(async () => {
        let l = ((await idbGet("kv", "gros").catch(() => null)) || []).filter((x) => x !== id);
        l.unshift(id);
        for (const vieux of l.slice(GROS_N)) await idbDel("content", vieux).catch(() => {});
        await idbPut("kv", "gros", l.slice(0, GROS_N));
      }).catch(() => {});
      return this.grosChain;
    }
    // Après une synchronisation : garder d'avance dans le téléphone tous les textes (notes, etc.)
    // pas encore à jour dans le cache, pour pouvoir travailler hors ligne. Quelques-uns à la fois.
    prefetch() {
      if (this.prefetching) return this.prefetching;
      if (this.offline) return Promise.resolve(null);
      const liste = [...this.items.values()].filter((it) => !it.dir && it.eTag && it.size <= PREFETCH_MAX
        && PREFETCH_EXT.test(it.name) && !String(it.id).startsWith("local:"));
      this.prefetching = (async () => {
        let faits = 0;
        const un = async (it) => {
          if (this.pending.has(it.id)) return;
          const c = await idbGet("content", it.id).catch(() => null);
          if (c && c.eTag === it.eTag) return;
          try { await this.content(it); faits++; } catch (e) { if (horsLigne(e)) throw e; }
        };
        for (let i = 0; i < liste.length && !this.offline; i += 4) await Promise.all(liste.slice(i, i + 4).map(un));
        return { textes: liste.length, telecharges: faits };
      })().finally(() => { this.prefetching = null; });
      return this.prefetching;
    }

    /* ---- file d'attente ---- */
    async loadQueue() {
      if (this.queueLoaded) return; this.queueLoaded = true;
      try { for (const j of (await idbAll("queue")) || []) if (j && j.id) this.pending.set(j.id, j); } catch (e) {}
      this.status();
    }
    persist(job) { const { ...copie } = job; return idbPut("queue", job.id, copie).catch(() => {}); }
    status() { try { this.onStatus({ attente: this.pending.size }); } catch (e) {} }
    flushSoon(ms) { clearTimeout(this.flushTimer); this.flushTimer = setTimeout(() => this.flush().catch(() => {}), ms == null ? 1500 : ms); }
    flush() {
      if (this.flushing) return this.flushing;
      if (!this.pending.size) return Promise.resolve();
      this.flushing = (async () => {
        for (const job of [...this.pending.values()].sort((x, y) => x.at - y.at)) {
          try {
            const envoye = await this.send(job);
            if (job.text === envoye) { this.pending.delete(job.id); await idbDel("queue", job.id).catch(() => {}); }
            else await this.persist(job);                    // modifié pendant l'envoi : on renverra la suite
          } catch (e) {
            if (horsLigne(e)) break;
            job.erreur = String(e.message || e).slice(0, 200); await this.persist(job);
            this.log("file d'attente : " + job.erreur);
          }
        }
      })().finally(() => { this.flushing = null; this.status(); if (this.pending.size) this.flushSoon(30000); });
      return this.flushing;
    }
    // Un fichier créé hors ligne a reçu son vrai id.
    adopt(job, raw) {
      const ancien = job.id;
      if (ancien && ancien !== raw.id) {
        if (String(ancien).startsWith("local:")) { this.alias.set(ancien, raw.id); this.drop(ancien); }
        if (this.pending.get(ancien) === job) { this.pending.delete(ancien); idbDel("queue", ancien).catch(() => {}); this.pending.set(raw.id, job); }
      }
      job.id = raw.id; job.isNew = false;
    }
    // Version actuelle dans OneDrive (métadonnées + texte, sans le cache).
    async fresh(job) {
      const url = job.isNew ? `/me/drive/items/${job.parentId}:/${enc(job.name)}?$select=${SEL}` : `/me/drive/items/${job.id}?$select=${SEL}`;
      const raw = await this.json(url);
      const r = await this.req(`/me/drive/items/${raw.id}/content`);
      if (!r.ok) { const e = new Error(`Lecture impossible (${r.status}) : ${job.name}`); e.status = r.status; throw e; }
      return { raw, text: await r.text() };
    }
    // Envoie un changement ; renvoie le texte effectivement envoyé (pour savoir si la file est vide).
    async send(job) {
      const corps = (t) => new Blob([t], { type: "text/plain;charset=utf-8" });
      const envoye = job.text; let texte = envoye, perdus = 0, gagnant = "eux", perdant = null;
      for (let tour = 0; tour < 4; tour++) {
        let url; const headers = { "Content-Type": "text/plain; charset=utf-8" };
        if (!job.isNew) { url = `/me/drive/items/${job.id}/content`; if (job.baseETag) headers["If-Match"] = job.baseETag; }
        else url = `/me/drive/items/${job.parentId}:/${enc(job.name)}:/content?@microsoft.graph.conflictBehavior=fail`;
        const r = await this.req(url, { method: "PUT", headers, body: corps(texte) });
        if (r.status === 412 || r.status === 409) {
          // Changé des deux côtés : fusion ligne par ligne. Là où les deux ont touché les mêmes lignes :
          // l'ordinateur gagne si les deux changements ont moins de 5 jours d'écart ; sinon le plus récent gagne.
          const eux = await this.fresh(job);
          if (job.isNew || String(job.id).startsWith("local:")) this.adopt(job, eux.raw);
          const fsi = eux.raw.fileSystemInfo || {};
          const tEux = Date.parse(fsi.lastModifiedDateTime || eux.raw.lastModifiedDateTime || "") || 0;
          gagnant = (job.at || 0) - tEux > this.ecart ? "moi" : "eux";
          let f;
          if (job.baseText == null) f = eux.text === texte ? { texte, conflits: 0 } : { texte: gagnant === "moi" ? texte : eux.text, conflits: 1 };
          else f = fusion3(job.baseText, texte, eux.text, gagnant);
          if (f.conflits) perdant = gagnant === "moi" ? { qui: "ordinateur", texte: eux.text } : { qui: "téléphone", texte };
          perdus += f.conflits;
          job.baseText = eux.text; job.baseETag = eux.raw.eTag; texte = f.texte;
          if (texte === eux.text) {                              // rien à ajouter : la version de OneDrive suffit
            this.put(eux.raw, job.parentId);
            if (eux.text.length <= CACHE_MAX) idbPut("content", eux.raw.id, { eTag: eux.raw.eTag, blob: corps(eux.text) }).catch(() => {});
            this.save().catch(() => {});
            await this.finConflit(job, perdus, gagnant, perdant);
            return envoye;
          }
          continue;
        }
        if (!r.ok) { const e = new Error(`Écriture impossible (${r.status}) : ${(await r.text()).slice(0, 160)}`); e.status = r.status; throw e; }
        const raw = await r.json();
        if (job.isNew || String(job.id).startsWith("local:")) this.adopt(job, raw);
        this.put(raw, job.parentId);
        job.baseETag = raw.eTag; job.baseText = texte;
        if (texte.length <= CACHE_MAX) idbPut("content", raw.id, { eTag: raw.eTag, blob: corps(texte) }).catch(() => {});
        this.save().catch(() => {});
        await this.finConflit(job, perdus, gagnant, perdant);
        return envoye;
      }
      throw new Error(`« ${job.name} » change sans arrêt ailleurs; réessaie dans un moment.`);
    }

    // Après un vrai conflit : la version perdante entière est gardée dans « Conflits (téléphone)/ », puis on prévient.
    async finConflit(job, perdus, gagnant, perdant) {
      if (!perdus) return;
      let copie = null;
      if (perdant && !String(job.path || "").startsWith(DOSSIER_CONFLITS + "/")) {
        try {
          let d = this.child(this.rootId, DOSSIER_CONFLITS);
          if (!d) d = await this.mkdir(this.rootId, DOSSIER_CONFLITS);
          const t = new Date(), p = (n) => String(n).padStart(2, "0");
          const quand = `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())} ${p(t.getHours())}h${p(t.getMinutes())}${p(t.getSeconds())}`;
          copie = `${quand} — ${perdant.qui} — ${job.name}`;
          const md = /\.md$/i.test(job.name);
          const entete = md ? `> [!warning] Copie de conflit — version du ${perdant.qui} qui n'a pas gagné\n> Fichier : [[${String(job.path).replace(/\.md$/i, "")}]] — gardée le ${quand}. Comparer, garder ce qu'il faut, puis supprimer cette copie.\n\n` : "";
          await this.write(d.id, copie, null, entete + perdant.texte);
        } catch (e) { copie = null; this.log("copie de conflit : " + (e.message || e)); }
      }
      try { this.onConflict({ chemin: job.path, nom: job.name, conflits: perdus, gagnant: gagnant === "moi" ? "téléphone" : "ordinateur", copie, dossier: DOSSIER_CONFLITS }); } catch (e) {}
    }

    /* ---- écriture ---- */
    async write(parentId, name, existingId, data) {
      parentId = this.real(parentId); existingId = this.real(existingId);
      const path = (this.pathOf(parentId) ? this.pathOf(parentId) + "/" : "") + nfc(name);
      if (!this.canWrite(path)) throw new Error(`Écriture désactivée sur le téléphone pour « ${path} ».`);
      const text = data instanceof Blob ? await data.text() : String(data);
      // Déjà un changement en attente pour ce fichier : on le remplace (même point de départ).
      const enAttente = existingId && this.pending.get(existingId);
      if (enAttente) {
        enAttente.text = text; enAttente.at = Date.now(); enAttente.erreur = null; await this.persist(enAttente);
        const it = this.items.get(existingId); if (it) it.mtime = Date.now();
        this.flushSoon(800); return it;
      }
      const job = { id: existingId || null, parentId, name: nfc(name), path, isNew: !existingId, text, baseETag: null, baseText: null, at: Date.now() };
      if (existingId) {
        const it = this.items.get(existingId); job.baseETag = (it && it.eTag) || null;
        const c = await idbGet("content", existingId).catch(() => null);
        if (c && c.eTag && c.eTag === job.baseETag) job.baseText = await c.blob.text();
      }
      try { await this.send(job); return this.items.get(job.id); }
      catch (e) { if (!horsLigne(e)) throw e; }
      // Pas de réseau / connexion expirée : on garde le changement dans le téléphone.
      if (!job.id) {
        job.id = "local:" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
        this.put({ id: job.id, name: job.name, file: { mimeType: "text/markdown" }, size: text.length, lastModifiedDateTime: new Date().toISOString() }, parentId);
      }
      const it = this.items.get(job.id); if (it) it.mtime = Date.now();
      this.pending.set(job.id, job); await this.persist(job); this.status();
      this.flushSoon(15000);
      return it;
    }
    async mkdir(parentId, name) {
      parentId = this.real(parentId);
      const path = (this.pathOf(parentId) ? this.pathOf(parentId) + "/" : "") + nfc(name) + "/";
      if (!this.canWrite(path)) throw new Error(`Création de dossier désactivée sur le téléphone : « ${path} ».`);
      const raw = await this.json(`/me/drive/items/${parentId}/children`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: nfc(name), folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
      });
      const it = this.put(raw, parentId); it.kids = new Set(); return it;
    }

    // Efface tout ce qui est dans le téléphone. La file d'attente n'est effacée que si on le demande.
    async signOutClear(aussiFile) {
      this.items.clear(); this.deltaLinks = {}; this.rootId = null; this.syncAt = 0;
      await idbClear("kv"); await idbClear("content");
      if (aussiFile) { this.pending.clear(); await idbClear("queue"); }
    }

    get root() { return new GDir(this, this.rootId); }
  }

  /* ---------- poignées ---------- */
  class GDir {
    constructor(fs, id) { this.fs = fs; this.id = id; this.kind = "directory"; }
    get name() { return this.fs.items.get(this.id)?.name || ""; }
    async *entries() {
      await this.fs.loadKids(this.id);
      const d = this.fs.items.get(this.id);
      for (const k of [...(d?.kids || [])]) {
        const c = this.fs.items.get(k); if (!c) continue;
        yield [c.name, c.dir ? new GDir(this.fs, c.id) : new GFile(this.fs, c.id, this.id, c.name)];
      }
    }
    async *values() { for await (const [, h] of this.entries()) yield h; }
    async *keys() { for await (const [n] of this.entries()) yield n; }
    [Symbol.asyncIterator]() { return this.entries(); }
    async getDirectoryHandle(name, opts) {
      await this.fs.loadKids(this.id);
      const c = this.fs.child(this.id, name);
      if (c && c.dir) return new GDir(this.fs, c.id);
      if (c) { const e = new Error(`« ${name} » est un fichier`); e.name = "TypeMismatchError"; throw e; }
      if (opts && opts.create) return new GDir(this.fs, (await this.fs.mkdir(this.id, name)).id);
      throw notFound(name);
    }
    async getFileHandle(name, opts) {
      await this.fs.loadKids(this.id);
      const c = this.fs.child(this.id, name);
      if (c && !c.dir) return new GFile(this.fs, c.id, this.id, c.name);
      if (c) { const e = new Error(`« ${name} » est un dossier`); e.name = "TypeMismatchError"; throw e; }
      if (opts && opts.create) return new GFile(this.fs, null, this.id, nfc(name));
      throw notFound(name);
    }
    async queryPermission() { return "granted"; }
    async requestPermission() { return "granted"; }
  }

  class GFile {
    constructor(fs, id, parentId, name) { this.fs = fs; this.id = id; this.parentId = parentId; this._name = name; this.kind = "file"; }
    get name() { return (this.id && this.fs.items.get(this.fs.real(this.id))?.name) || this._name; }
    async getFile() {
      if (!this.id) return new File([""], this._name, { lastModified: Date.now() });
      this.id = this.fs.real(this.id);
      const it = this.fs.items.get(this.id); if (!it) throw notFound(this._name);
      const blob = await this.fs.content(it);
      return new File([blob], it.name, { type: it.mime || blob.type || "", lastModified: it.mtime || Date.now() });
    }
    async createWritable() {
      const parts = []; const self = this;
      return {
        async write(d) { if (d && typeof d === "object" && "type" in d && d.type === "write") d = d.data; parts.push(d); },
        async truncate() { parts.length = 0; },
        async abort() { parts.length = 0; },
        async close() {
          const blob = new Blob(parts.map((p) => (p instanceof Blob ? p : typeof p === "string" ? p : new Uint8Array(p.buffer || p))), { type: "text/plain;charset=utf-8" });
          const it = await self.fs.write(self.parentId, self.name, self.id, blob);
          if (it) self.id = it.id;
        }
      };
    }
  }

  window.GraphFS = GraphFS;
  window.GraphFS.fusion3 = fusion3;   // exposé pour les tests
})();
