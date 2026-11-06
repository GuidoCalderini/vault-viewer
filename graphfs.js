/* GraphFS — imitation, sur OneDrive (Microsoft Graph), de la petite partie de l'API
   « File System Access » dont la page du vault se sert :
     dossier : name, kind, entries(), values(), keys(), getDirectoryHandle(n,{create}), getFileHandle(n,{create})
     fichier : name, kind, getFile() → File (text(), lastModified, type), createWritable() → write(), close()
   L'inventaire vient de /delta (un par dossier de premier niveau), gardé dans IndexedDB : au lancement
   suivant, seuls les changements sont téléchargés. Le contenu des fichiers est mis en cache par eTag.
   Écriture : PUT avec If-Match (jamais d'écrasement silencieux), seulement sur la liste blanche. */
(function () {
  "use strict";
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SEL = "id,name,parentReference,file,folder,deleted,eTag,size,lastModifiedDateTime";
  const CACHE_MAX = 8 * 1024 * 1024;   // on ne garde pas en cache les fichiers plus gros que ça
  const TREE_VERSION = 1;

  /* ---------- IndexedDB ---------- */
  function idbOpen() {
    return new Promise((res, rej) => {
      const r = indexedDB.open("vv-graphfs", 1);
      r.onupgradeneeded = () => { r.result.createObjectStore("kv"); r.result.createObjectStore("content"); };
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

  class GraphFS {
    constructor({ getToken, vaultFolders, skipDirs, canWrite, log }) {
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
    }

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
      await idbPut("kv", "tree", { v: TREE_VERSION, rootId: this.rootId, rootName: this.rootName, deltaLinks: this.deltaLinks, items, at: Date.now() });
    }
    async restore() {
      const t = await idbGet("kv", "tree");
      if (!t || t.v !== TREE_VERSION) return false;
      this.rootId = t.rootId; this.rootName = t.rootName; this.deltaLinks = t.deltaLinks || {};
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

    // Lancement : cache, puis racine (1 requête), puis /delta de chaque dossier de 1er niveau.
    async sync(progress) {
      progress = progress || (() => {});
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
      return { elements: this.items.size, changements: changes };
    }

    /* ---- contenu ---- */
    async content(it) {
      const c = await idbGet("content", it.id).catch(() => null);
      if (c && c.eTag && c.eTag === it.eTag) { this.stats.depuisCache++; return c.blob; }
      const r = await this.req(`/me/drive/items/${it.id}/content`);
      if (!r.ok) { const e = new Error(`Lecture impossible (${r.status}) : ${it.name}`); e.status = r.status; throw e; }
      const blob = await r.blob(); this.stats.telecharges++;
      if (blob.size <= CACHE_MAX) idbPut("content", it.id, { eTag: it.eTag, blob }).catch(() => {});
      return blob;
    }

    /* ---- écriture ---- */
    async write(parentId, name, existingId, data) {
      const path = (this.pathOf(parentId) ? this.pathOf(parentId) + "/" : "") + nfc(name);
      if (!this.canWrite(path)) throw new Error(`Écriture désactivée sur le téléphone pour « ${path} ».`);
      const body = data instanceof Blob ? data : new Blob([data], { type: "text/plain;charset=utf-8" });
      let url, headers = { "Content-Type": "text/plain; charset=utf-8" };
      if (existingId) { url = `/me/drive/items/${existingId}/content`; const it = this.items.get(existingId); if (it && it.eTag) headers["If-Match"] = it.eTag; }
      else url = `/me/drive/items/${parentId}:/${enc(nfc(name))}:/content?@microsoft.graph.conflictBehavior=fail`;
      const r = await this.req(url, { method: "PUT", headers, body });
      if (r.status === 412 || r.status === 409) {
        this.sync().catch(() => {});
        throw new Error(`« ${nfc(name)} » a changé ailleurs (sur l'ordinateur?). Rien n'a été écrasé : recharge la page, puis refais le geste.`);
      }
      if (!r.ok) throw new Error(`Écriture impossible (${r.status}) : ${(await r.text()).slice(0, 160)}`);
      const raw = await r.json();
      const it = this.put(raw, parentId);
      if (body.size <= CACHE_MAX) idbPut("content", it.id, { eTag: it.eTag, blob: body }).catch(() => {});
      this.save().catch(() => {});
      return it;
    }
    async mkdir(parentId, name) {
      const path = (this.pathOf(parentId) ? this.pathOf(parentId) + "/" : "") + nfc(name) + "/";
      if (!this.canWrite(path)) throw new Error(`Création de dossier désactivée sur le téléphone : « ${path} ».`);
      const raw = await this.json(`/me/drive/items/${parentId}/children`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: nfc(name), folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
      });
      const it = this.put(raw, parentId); it.kids = new Set(); return it;
    }

    async signOutClear() { this.items.clear(); this.deltaLinks = {}; this.rootId = null; await idbClear("kv"); await idbClear("content"); }

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
    get name() { return (this.id && this.fs.items.get(this.id)?.name) || this._name; }
    async getFile() {
      if (!this.id) return new File([""], this._name, { lastModified: Date.now() });
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
          self.id = it.id;
        }
      };
    }
  }

  window.GraphFS = GraphFS;
})();
