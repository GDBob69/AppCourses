const APP = {
  title: 'Courses & Menus',
  sheets: {
    settings: ['CLE', 'VALEUR'],
    catalogue: ['ID', 'TYPOLOGIE', 'PRODUIT', 'MAGASIN_DEFAUT', 'RAYON', 'CONDITIONNEMENT', 'ACTIF', 'DERNIER_ACHAT', 'MAJ', 'QUANTITE_DEFAUT', 'ICONE'],
    courses: ['ID', 'CATALOGUE_ID', 'PRODUIT', 'TYPOLOGIE', 'MAGASIN_DEFAUT', 'MAGASIN_ACTUEL', 'RAYON', 'QUANTITE', 'CONDITIONNEMENT', 'COMMENTAIRE', 'STATUT', 'AJOUTE_LE', 'ACHETE_LE', 'MAJ', 'MENUS_IDS'],
    menus: ['ID', 'DATE', 'REPAS', 'PERSONNES', 'MENU', 'URL', 'COMMENTAIRE', 'CHOUCHOU', 'LOULOU', 'DEPLACEMENT', 'STATUT', 'MAJ', 'EVALUATION', 'ETAT_MENU'],
    restes: ['ID', 'NOM', 'QUANTITE', 'UNITE', 'DATE_LIMITE', 'URGENCE', 'DATE_UTILISATION', 'REPAS', 'STATUT', 'COMMENTAIRE', 'MAJ', 'CATEGORIE'],
    recettes: ['ID', 'NOM', 'URL', 'COMMENTAIRE', 'MAJ'],
    historiqueMenus: ['ID', 'DATE', 'REPAS', 'PERSONNES', 'MENU', 'URL', 'COMMENTAIRE', 'CHOUCHOU', 'LOULOU', 'DEPLACEMENT', 'STATUT', 'MAJ', 'EVALUATION', 'ETAT_MENU'],
    references: ['ID', 'TYPE', 'VALEUR', 'ACTIF', 'ORDRE', 'MAJ'],
    ideesPlats: ['ID', 'PLAT', 'URL', 'COMMENTAIRE', 'STATUT', 'DATE_PLANIFIEE', 'REPAS_PLANIFIE', 'MAJ', 'EVALUATION'],
    legumesSaison: ['ID', 'ITEM', 'MOIS', 'MAJ'],
    menuIngredients: ['ID', 'MENU_ID', 'CATALOGUE_ID', 'PRODUIT', 'QUANTITE', 'UNITE', 'COMMENTAIRE', 'STATUT', 'COURSE_ID', 'RESTE_ID', 'MAJ']
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupApp() {
  return locked_(function () {
    return setupApp_();
  });
}

function setupApp_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(APP.sheets).forEach(function (key) {
    const name = sheetName_(key);
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = APP.sheets[key];
    ensureSheetSize_(sh, 1, headers.length);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#184E47').setFontColor('#FFFFFF');
    if (sh.getMaxColumns() > headers.length) {
      sh.deleteColumns(headers.length + 1, sh.getMaxColumns() - headers.length);
    }
  });
  setSetting_('REVISION', String(Date.now()));
  seedReferences_();
  seedSeasonalVegetables_();
  setSetting_('APP_VERSION', '2.4.0');
  setSetting_('WEEK_START', '6');
  return { ok: true, spreadsheetUrl: ss.getUrl() };
}

function getBootstrap() {
  ensureSetup_();
  const today = isoDate_(new Date());
  const minMenuDate = new Date();
  minMenuDate.setDate(minMenuDate.getDate() - 31);
  const maxMenuDate = new Date();
  maxMenuDate.setDate(maxMenuDate.getDate() + 180);
  const minMenuIso = isoDate_(minMenuDate);
  const maxMenuIso = isoDate_(maxMenuDate);
  const allCourses = rows_('courses');
  return {
    revision: getRevision_(),
    today: today,
    settings: { weekStart: Number(getSetting_('WEEK_START') || 6) },
    catalogue: rows_('catalogue').filter(function (x) { return x.ACTIF !== false && x.ACTIF !== 'FALSE'; }),
    courses: allCourses.filter(isActiveCourse_),
    preparation: allCourses.filter(function (x) {
      return ['A_DECIDER', 'A_VERIFIER', 'IGNORE'].indexOf(x.STATUT) >= 0;
    }),
    menus: rows_('menus').filter(function (x) {
      return x.STATUT !== 'ARCHIVE' && x.DATE >= minMenuIso && x.DATE <= maxMenuIso;
    }),
    restes: rows_('restes').filter(isVisibleLeftover_),
    recettes: [],
    references: getReferences_(),
    ideesPlats: rows_('ideesPlats').filter(function (x) { return x.STATUT !== 'SUPPRIME'; }),
    legumesSaison: rows_('legumesSaison'),
    menuIngredients: rows_('menuIngredients')
  };
}

function getBootstrapLight() {
  ensureSetup_();
  const today = isoDate_(new Date());
  const minMenuDate = new Date();
  minMenuDate.setDate(minMenuDate.getDate() - 31);
  const maxMenuDate = new Date();
  maxMenuDate.setDate(maxMenuDate.getDate() + 180);
  const minMenuIso = isoDate_(minMenuDate);
  const maxMenuIso = isoDate_(maxMenuDate);
  return {
    revision: getRevision_(),
    today: today,
    settings: { weekStart: Number(getSetting_('WEEK_START') || 6) },
    catalogue: [],
    courses: [],
    preparation: [],
    menus: rows_('menus').filter(function (x) {
      return x.STATUT !== 'ARCHIVE' && x.DATE >= minMenuIso && x.DATE <= maxMenuIso;
    }),
    restes: rows_('restes').filter(isVisibleLeftover_),
    recettes: [],
    references: getReferences_(),
    ideesPlats: rows_('ideesPlats').filter(function (x) { return x.STATUT !== 'SUPPRIME'; }),
    legumesSaison: rows_('legumesSaison'),
    menuIngredients: rows_('menuIngredients')
  };
}

function getBootstrapHome() {
  ensureSetup_();
  const today = isoDate_(new Date());
  const minMenuDate = new Date();
  minMenuDate.setDate(minMenuDate.getDate() - 31);
  const maxMenuDate = new Date();
  maxMenuDate.setDate(maxMenuDate.getDate() + 60);
  const minMenuIso = isoDate_(minMenuDate);
  const maxMenuIso = isoDate_(maxMenuDate);
  return {
    revision: getRevision_(),
    today: today,
    settings: { weekStart: Number(getSetting_('WEEK_START') || 6) },
    menus: rows_('menus').filter(function (x) {
      return x.STATUT !== 'ARCHIVE' && x.DATE >= minMenuIso && x.DATE <= maxMenuIso;
    }),
    restes: rows_('restes').filter(isVisibleLeftover_),
    menuIngredients: rows_('menuIngredients')
  };
}

function getBootstrapHomeIfChanged(clientRevision) {
  ensureSetup_();
  const revision = getRevision_();
  if (String(clientRevision || '') === String(revision || '')) {
    return { unchanged: true, revision: revision };
  }
  const data = getBootstrapHome();
  data.unchanged = false;
  return data;
}

function getShoppingData() {
  ensureSetup_();
  const allCourses = rows_('courses');
  return {
    revision: getRevision_(),
    courses: allCourses.filter(isActiveCourse_),
    preparation: allCourses.filter(function (x) {
      return ['A_DECIDER', 'A_VERIFIER', 'IGNORE'].indexOf(x.STATUT) >= 0;
    }),
    references: getReferences_()
  };
}

function getCourseListData() {
  ensureSetup_();
  return {
    revision: getRevision_(),
    courses: rows_('courses').filter(isActiveCourse_)
  };
}

function getCatalogData() {
  ensureSetup_();
  const shopping = getShoppingData();
  return {
    revision: shopping.revision,
    catalogue: rows_('catalogue').filter(function (x) { return x.ACTIF !== false && x.ACTIF !== 'FALSE'; }),
    courses: shopping.courses,
    preparation: shopping.preparation,
    references: getReferences_()
  };
}

function getSeasonData() {
  ensureSetup_();
  return {
    revision: getRevision_(),
    legumesSaison: rows_('legumesSaison')
  };
}

function startNewPreparation() {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const candidates = rows_('catalogue').filter(function (x) {
      return x.ACTIF !== false && x.ACTIF !== 'FALSE';
    }).map(function (x) {
      return {
        ID: Utilities.getUuid(), CATALOGUE_ID: x.ID, PRODUIT: x.PRODUIT,
        TYPOLOGIE: x.TYPOLOGIE, MAGASIN_DEFAUT: x.MAGASIN_DEFAUT,
        MAGASIN_ACTUEL: x.MAGASIN_DEFAUT, RAYON: x.RAYON,
        QUANTITE: x.QUANTITE_DEFAUT || '1', CONDITIONNEMENT: x.CONDITIONNEMENT,
        COMMENTAIRE: '', STATUT: 'A_DECIDER', AJOUTE_LE: now,
        ACHETE_LE: '', MAJ: now
      };
    });
    clearData_('courses');
    appendRecords_('courses', candidates);
    setSetting_('PREPARATION_STARTED', isoDate_(now));
    bumpRevision_();
    return { ok: true, candidates: candidates.map(serializeRecord_) };
  });
}

function serializeRecord_(record) {
  const out = {};
  Object.keys(record).forEach(function (key) { out[key] = serialize_(record[key]); });
  return out;
}

function setPreparationDecision(id, status) {
  return locked_(function () {
    const allowed = ['A_DECIDER', 'A_VERIFIER', 'A_ACHETER', 'A_VERIFIER_STOCK', 'IGNORE'];
    if (allowed.indexOf(status) < 0) throw new Error('Décision invalide.');
    const record = findById_('courses', id);
    if (!record) return { ok: false, missing: true, id: id };
    record.STATUT = status;
    record.MAJ = new Date();
    upsert_('courses', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function setPreparationDecisions(items) {
  return locked_(function () {
    ensureSetup_();
    const allowed = ['A_DECIDER', 'A_VERIFIER', 'A_ACHETER', 'A_VERIFIER_STOCK', 'IGNORE'];
    const now = new Date();
    const results = [];
    let changed = false;
    (items || []).forEach(function (item) {
      const status = item && item.status;
      if (allowed.indexOf(status) < 0) return;
      const record = findById_('courses', item.id);
      if (!record) {
        results.push({ ok: false, missing: true, id: item.id });
        return;
      }
      record.STATUT = status;
      if ((status === 'A_ACHETER' || status === 'A_VERIFIER_STOCK') && item.note !== undefined) record.COMMENTAIRE = clean_(item.note);
      record.MAJ = now;
      upsert_('courses', record);
      results.push(record);
      changed = true;
    });
    if (changed) bumpRevision_();
    return { ok: true, results: results.map(serializeRecord_) };
  });
}

function getRevision() {
  ensureSetup_();
  return { revision: getRevision_() };
}

function searchMenuHistory(query, limit, month, minRating) {
  ensureSetup_();
  const q = normalizeSearch_(query);
  const m = clean_(month);
  const rating = Number(minRating) || 0;
  const max = Math.min(Math.max(Number(limit) || 80, 1), 300);
  return getMenuHistorySearchRows_()
    .filter(function (x) {
      if (m && String(x.DATE || '').slice(5, 7) !== m) return false;
      if (rating && Number(x.EVALUATION || 0) < rating) return false;
      if (!q) return true;
      return normalizeSearch_([x.MENU, x.COMMENTAIRE, x.PERSONNES, x.DATE, x.REPAS].join(' ')).indexOf(q) >= 0;
    })
    .sort(function (a, b) { return String(b.DATE).localeCompare(String(a.DATE)); })
    .slice(0, max);
}

function getMenuHistoryIndex() {
  ensureSetup_();
  return {
    revision: getRevision_(),
    rows: getMenuHistorySearchRows_().map(function (x) {
      x.SEARCH_TEXT = normalizeSearch_([
        x.MENU, x.COMMENTAIRE, x.PERSONNES, x.DATE, x.REPAS,
        x.CHOUCHOU, x.LOULOU, x.DEPLACEMENT, x.URL, x.EVALUATION
      ].join(' '));
      return x;
    })
  };
}

function normalizeSearch_(value) {
  return clean_(value).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function defaultProductIcon_(product, aisle) {
  const text = normalizeSearch_([product, aisle].join(' '));
  if (/pomme|poire|banane|orange|citron|fruit|fraise|raisin|kiwi|melon|peche|abricot/.test(text)) return '🍎';
  if (/legume|salade|tomate|carotte|courgette|poireau|oignon|ail|pomme de terre|patate/.test(text)) return '🥕';
  if (/lait|yaourt|fromage|creme|beurre|oeuf|œuf|cremerie/.test(text)) return '🥛';
  if (/viande|poulet|boeuf|bœuf|porc|jambon|poisson|saumon|thon|crevette/.test(text)) return '🥩';
  if (/riz|pate|pates|farine|sucre|huile|conserve|epicerie|cereal/.test(text)) return '🛒';
  if (/surgel|glace|congele/.test(text)) return '❄️';
  if (/chat|chien|animal|litiere|croquette/.test(text)) return '🐾';
  if (/lessive|nettoy|eponge|entretien|sac poubelle|papier toilette/.test(text)) return '🧽';
  if (/savon|shampo|dentifrice|hygiene|deodorant/.test(text)) return '🧴';
  if (/pain|brioche|boulanger/.test(text)) return '🥖';
  if (/boisson|eau|jus|soda|vin|biere/.test(text)) return '🥤';
  return '🛒';
}

function getMenuHistorySearchRows_() {
  const cache = CacheService.getScriptCache();
  const key = 'MENU_HISTORY_SEARCH_' + getRevision_();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const rows = readMenuHistorySheet_('historiqueMenus')
    .concat(readMenuHistorySheet_('menus').filter(function (x) { return x.STATUT !== 'ARCHIVE'; }));
  try {
    cache.put(key, JSON.stringify(rows), 1800);
  } catch (e) {
    // CacheService est limité en taille : si l'historique devient trop volumineux,
    // on garde simplement la recherche directe.
  }
  return rows;
}

function readMenuHistorySheet_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_(key));
  if (!sh || sh.getLastRow() < 2) return [];
  const headers = APP.sheets[key];
  const wanted = ['ID', 'DATE', 'REPAS', 'PERSONNES', 'MENU', 'URL', 'COMMENTAIRE', 'CHOUCHOU', 'LOULOU', 'DEPLACEMENT', 'STATUT', 'EVALUATION'];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  const indexes = {};
  headers.forEach(function (h, i) { indexes[h] = i; });
  return values.map(function (row) {
    const obj = {};
    wanted.forEach(function (h) { obj[h] = serialize_(row[indexes[h]]); });
    return obj;
  }).filter(function (x) { return x.MENU || x.COMMENTAIRE; });
}

function saveCourse(item) {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const existing = item.id ? findById_('courses', item.id) : null;
    const catalogueId = item.catalogueId !== undefined ? item.catalogueId : (existing && existing.CATALOGUE_ID) || '';
    const catalogue = catalogueId ? findById_('catalogue', catalogueId) : null;
    const record = {
      ID: item.id || Utilities.getUuid(),
      CATALOGUE_ID: catalogueId || '',
      PRODUIT: clean_(item.product || (existing && existing.PRODUIT) || (catalogue && catalogue.PRODUIT)),
      TYPOLOGIE: clean_(item.aisle || item.category || (existing && (existing.RAYON || existing.TYPOLOGIE)) || (catalogue && (catalogue.RAYON || catalogue.TYPOLOGIE)) || 'Autres'),
      MAGASIN_DEFAUT: clean_(item.defaultStore || (existing && existing.MAGASIN_DEFAUT) || (catalogue && catalogue.MAGASIN_DEFAUT)),
      MAGASIN_ACTUEL: clean_(item.store || (existing && existing.MAGASIN_ACTUEL) || (catalogue && catalogue.MAGASIN_DEFAUT)),
      RAYON: clean_(item.aisle || (existing && existing.RAYON) || (catalogue && catalogue.RAYON)),
      QUANTITE: clean_(item.quantity || (existing && existing.QUANTITE) || (catalogue && catalogue.QUANTITE_DEFAUT) || '1'),
      CONDITIONNEMENT: clean_(item.package || (existing && existing.CONDITIONNEMENT) || (catalogue && catalogue.CONDITIONNEMENT)),
      COMMENTAIRE: clean_(item.note !== undefined ? item.note : (existing && existing.COMMENTAIRE)),
      STATUT: normalizeCourseStatus_(item.status || (existing && existing.STATUT) || 'A_ACHETER'),
      AJOUTE_LE: existing ? existing.AJOUTE_LE : now,
      ACHETE_LE: existing ? existing.ACHETE_LE : '',
      MAJ: now,
      MENUS_IDS: clean_(item.menuIds !== undefined ? item.menuIds : (existing && existing.MENUS_IDS))
    };
    if (!record.PRODUIT) throw new Error('Le nom du produit est obligatoire.');
    upsert_('courses', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function completeCourse(id) {
  return locked_(function () {
    const record = findById_('courses', id);
    if (!record) return { ok: true };
    const now = new Date();
    record.STATUT = 'ACHETE';
    record.ACHETE_LE = now;
    record.MAJ = now;
    upsert_('courses', record);
    if (record.CATALOGUE_ID) {
      const cat = findById_('catalogue', record.CATALOGUE_ID);
      if (cat) {
        cat.DERNIER_ACHAT = now;
        cat.MAJ = now;
        upsert_('catalogue', cat);
      }
    }
    bumpRevision_();
    return { ok: true };
  });
}

function completeCourses(ids) {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const results = [];
    let changed = false;
    (ids || []).forEach(function (id) {
      const record = findById_('courses', id);
      if (!record) {
        results.push({ ok: true, missing: true, id: id });
        return;
      }
      record.STATUT = 'ACHETE';
      record.ACHETE_LE = now;
      record.MAJ = now;
      upsert_('courses', record);
      if (record.CATALOGUE_ID) {
        const cat = findById_('catalogue', record.CATALOGUE_ID);
        if (cat) {
          cat.DERNIER_ACHAT = now;
          cat.MAJ = now;
          upsert_('catalogue', cat);
        }
      }
      results.push({ ok: true, id: id });
      changed = true;
    });
    if (changed) bumpRevision_();
    return { ok: true, results: results };
  });
}

function reopenCourse(id) {
  return locked_(function () {
    ensureSetup_();
    const record = findById_('courses', id);
    if (!record) return { ok: false, missing: true, id: id };
    record.STATUT = 'A_ACHETER';
    record.ACHETE_LE = '';
    record.MAJ = new Date();
    upsert_('courses', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function updateCourseStore(id, store) {
  return locked_(function () {
    const record = findById_('courses', id);
    if (!record) return { ok: false, missing: true, id: id };
    record.MAGASIN_ACTUEL = clean_(store);
    record.MAJ = new Date();
    upsert_('courses', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function updateCourseStores(items) {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const results = [];
    let changed = false;
    (items || []).forEach(function (item) {
      const record = findById_('courses', item && item.id);
      if (!record) {
        results.push({ ok: false, missing: true, id: item && item.id });
        return;
      }
      record.MAGASIN_ACTUEL = clean_(item.store);
      record.MAJ = now;
      upsert_('courses', record);
      results.push(record);
      changed = true;
    });
    if (changed) bumpRevision_();
    return { ok: true, results: results.map(serializeRecord_) };
  });
}

function deleteCourse(id) {
  return locked_(function () {
    deleteById_('courses', id);
    bumpRevision_();
    return { ok: true };
  });
}

function saveCatalogue(item) {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const aisle = clean_(item.aisle);
    const record = {
      ID: item.id || Utilities.getUuid(),
      TYPOLOGIE: aisle || clean_(item.category || 'Autres'),
      PRODUIT: clean_(item.product),
      MAGASIN_DEFAUT: clean_(item.defaultStore),
      RAYON: aisle,
      CONDITIONNEMENT: clean_(item.package),
      ACTIF: true,
      DERNIER_ACHAT: item.lastPurchased || '',
      MAJ: now,
      QUANTITE_DEFAUT: clean_(item.defaultQuantity || '1'),
      ICONE: clean_(item.icon)
    };
    if (!record.PRODUIT) throw new Error('Le nom du produit est obligatoire.');
    record.ICONE = record.ICONE || defaultProductIcon_(record.PRODUIT, record.RAYON || record.TYPOLOGIE);
    upsert_('catalogue', record);
    syncOpenCoursesFromCatalogue_(record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function createPreparationItem(catalogueId) {
  return locked_(function () {
    ensureSetup_();
    const catalogue = findById_('catalogue', catalogueId);
    if (!catalogue) throw new Error('Produit de référence introuvable.');
    const existing = rows_('courses').filter(function (x) {
      return String(x.CATALOGUE_ID) === String(catalogueId) &&
        ['A_DECIDER', 'A_VERIFIER', 'A_ACHETER', 'A_VERIFIER_STOCK', 'IGNORE'].indexOf(x.STATUT) >= 0;
    })[0];
    if (existing) return serializeRecord_(existing);
    const now = new Date();
    const record = {
      ID: Utilities.getUuid(),
      CATALOGUE_ID: catalogue.ID,
      PRODUIT: catalogue.PRODUIT,
      TYPOLOGIE: catalogue.TYPOLOGIE,
      MAGASIN_DEFAUT: catalogue.MAGASIN_DEFAUT,
      MAGASIN_ACTUEL: catalogue.MAGASIN_DEFAUT,
      RAYON: catalogue.RAYON,
      QUANTITE: catalogue.QUANTITE_DEFAUT || '1',
      CONDITIONNEMENT: catalogue.CONDITIONNEMENT,
      COMMENTAIRE: '',
      STATUT: 'A_DECIDER',
      AJOUTE_LE: now,
      ACHETE_LE: '',
      MAJ: now
    };
    upsert_('courses', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function deleteCatalogue(id) {
  return locked_(function () {
    ensureSetup_();
    const record = findById_('catalogue', id);
    if (!record) return { ok: true };
    record.ACTIF = false;
    record.MAJ = new Date();
    upsert_('catalogue', record);
    rows_('courses').forEach(function (x) {
      if (String(x.CATALOGUE_ID) === String(id) && ['A_DECIDER', 'A_VERIFIER', 'A_ACHETER', 'A_VERIFIER_STOCK', 'IGNORE'].indexOf(x.STATUT) >= 0) {
        deleteById_('courses', x.ID);
      }
    });
    bumpRevision_();
    return { ok: true };
  });
}

function resetLeftovers(keepIds) {
  return locked_(function () {
    ensureSetup_();
    const keep = {};
    (keepIds || []).forEach(function (id) { keep[String(id)] = true; });
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_('restes'));
    const headers = APP.sheets.restes;
    if (!sh || sh.getLastRow() < 2) return { ok: true, removed: 0, kept: 0 };
    const rowCount = sh.getLastRow() - 1;
    const values = sh.getRange(2, 1, rowCount, headers.length).getValues();
    const keptRows = values.filter(function (row) { return keep[String(row[0])]; });
    const removed = values.filter(function (row) { return row[0] && !keep[String(row[0])]; }).length;
    sh.getRange(2, 1, rowCount, headers.length).clearContent();
    if (keptRows.length) sh.getRange(2, 1, keptRows.length, headers.length).setValues(keptRows);
    bumpRevision_();
    return { ok: true, removed: removed, kept: keptRows.length };
  });
}

function saveReference(item) {
  return locked_(function () {
    ensureSetup_();
    const type = normalizeReferenceType_(item.type);
    const value = clean_(item.value);
    if (!value) throw new Error('La valeur est obligatoire.');
    const now = new Date();
    const record = {
      ID: item.id || Utilities.getUuid(),
      TYPE: type,
      VALEUR: value,
      ACTIF: true,
      ORDRE: Number(item.order) || nextReferenceOrder_(type),
      MAJ: now
    };
    rows_('references').forEach(function (other) {
      if (String(other.ID) === String(record.ID)) return;
      if (other.TYPE !== type || clean_(other.VALEUR).toLowerCase() !== value.toLowerCase()) return;
      other.ACTIF = false;
      other.MAJ = now;
      upsert_('references', other);
    });
    clearDeletedReferenceMarker_(type, value, now);
    upsert_('references', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function deleteReference(id, typeHint, valueHint) {
  return locked_(function () {
    ensureSetup_();
    const record = findById_('references', id);
    const type = record ? record.TYPE : normalizeReferenceType_(typeHint);
    const rawValue = record ? record.VALEUR : valueHint;
    if (!record && !rawValue) return { ok: true, removed: 0 };
    const now = new Date();
    const value = clean_(rawValue).toLowerCase();
    let removed = 0;
    rows_('references').forEach(function (candidate) {
      const sameReference = String(candidate.ID) === String(id) || (candidate.TYPE === type && clean_(candidate.VALEUR).toLowerCase() === value);
      const parsed = candidate.TYPE === 'RAYON_MAGASIN' ? parseStoreAisleReference_(candidate) : { store: '', aisle: '' };
      const linkedStoreAisle = type === 'MAGASIN' && candidate.TYPE === 'RAYON_MAGASIN' && clean_(parsed.store).toLowerCase() === value;
      const linkedAisleStore = type === 'RAYON' && candidate.TYPE === 'RAYON_MAGASIN' && clean_(parsed.aisle).toLowerCase() === value;
      if (!sameReference && !linkedStoreAisle && !linkedAisleStore) return;
      candidate.ACTIF = false;
      candidate.MAJ = now;
      upsert_('references', candidate);
      removed++;
    });
    addDeletedReferenceMarker_(type, rawValue, now);
    bumpRevision_();
    return { ok: true, removed: removed };
  });
}

function normalizeReferenceType_(type) {
  return type === 'RAYON' ? 'RAYON' : type === 'LIEU_ENFANT' ? 'LIEU_ENFANT' : type === 'RESTE_CATEGORIE' ? 'RESTE_CATEGORIE' : 'MAGASIN';
}

function deletedReferenceType_(type) {
  return 'SUPPRIME_' + normalizeReferenceType_(type);
}

function addDeletedReferenceMarker_(type, value, now) {
  value = clean_(value);
  if (!value) return;
  const markerType = deletedReferenceType_(type);
  const existing = rows_('references').filter(function (x) {
    return x.TYPE === markerType && clean_(x.VALEUR).toLowerCase() === value.toLowerCase();
  })[0];
  upsert_('references', {
    ID: existing ? existing.ID : Utilities.getUuid(),
    TYPE: markerType,
    VALEUR: value,
    ACTIF: true,
    ORDRE: existing ? existing.ORDRE : 1,
    MAJ: now || new Date()
  });
}

function clearDeletedReferenceMarker_(type, value, now) {
  value = clean_(value);
  if (!value) return;
  const markerType = deletedReferenceType_(type);
  rows_('references').forEach(function (x) {
    if (x.TYPE !== markerType || clean_(x.VALEUR).toLowerCase() !== value.toLowerCase()) return;
    x.ACTIF = false;
    x.MAJ = now || new Date();
    upsert_('references', x);
  });
}

function saveReferencesOrder(type, orderedIds) {
  return locked_(function () {
    ensureSetup_();
    const refType = type === 'RAYON' ? 'RAYON' : type === 'LIEU_ENFANT' ? 'LIEU_ENFANT' : type === 'RESTE_CATEGORIE' ? 'RESTE_CATEGORIE' : 'MAGASIN';
    const order = {};
    (orderedIds || []).forEach(function (id, i) { order[String(id)] = i + 1; });
    rows_('references').forEach(function (record) {
      if (record.TYPE !== refType || order[String(record.ID)] === undefined) return;
      record.ORDRE = order[String(record.ID)];
      record.MAJ = new Date();
      upsert_('references', record);
    });
    bumpRevision_();
    return { ok: true };
  });
}

function saveStoreAisleOrder(store, orderedAisles) {
  return locked_(function () {
    ensureSetup_();
    const storeName = clean_(store);
    if (!storeName) throw new Error('Le magasin est obligatoire.');
    const now = new Date();
    const delimiter = '\u001f';
    const existing = rows_('references').filter(function (x) {
      return x.TYPE === 'RAYON_MAGASIN';
    });
    const byValue = {};
    existing.forEach(function (x) { byValue[String(x.VALEUR)] = x; });
    const kept = {};
    (orderedAisles || []).map(normalizeStoreAisleEntry_).filter(function (x) { return x.aisle; }).forEach(function (entry, i) {
      const aisle = entry.aisle;
      const value = storeName + delimiter + aisle;
      const record = byValue[value] || {
        ID: Utilities.getUuid(),
        TYPE: 'RAYON_MAGASIN',
        VALEUR: value
      };
      record.ACTIF = entry.visible;
      record.ORDRE = i + 1;
      record.MAJ = now;
      kept[value] = true;
      upsert_('references', record);
    });
    existing.forEach(function (record) {
      const parsed = parseStoreAisleReference_(record);
      if (parsed.store !== storeName || kept[String(record.VALEUR)]) return;
      record.ACTIF = false;
      record.MAJ = now;
      upsert_('references', record);
    });
    bumpRevision_();
    return { ok: true };
  });
}

function normalizeStoreAisleEntry_(entry) {
  if (typeof entry === 'string') return { aisle: clean_(entry), visible: true };
  const aisle = clean_(entry && (entry.aisle || entry.AISLE || entry.value || entry.VALEUR));
  const visible = !(entry && (entry.visible === false || entry.visible === 'false' || entry.visible === 'FALSE'));
  return { aisle: aisle, visible: visible };
}

function saveDishIdea(item) {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const existing = item.id ? findById_('ideesPlats', item.id) : null;
    const record = {
      ID: item.id || Utilities.getUuid(),
      PLAT: clean_(item.dish),
      URL: clean_(item.url),
      COMMENTAIRE: clean_(item.comment),
      STATUT: clean_(item.status || (existing && existing.STATUT) || 'A_TESTER'),
      DATE_PLANIFIEE: item.plannedDate ? parseDate_(item.plannedDate) : (existing ? existing.DATE_PLANIFIEE : ''),
      REPAS_PLANIFIE: item.plannedMeal === 'MIDI' ? 'MIDI' : item.plannedMeal === 'SOIR' ? 'SOIR' : (existing ? existing.REPAS_PLANIFIE : ''),
      MAJ: now,
      EVALUATION: clean_(item.rating)
    };
    if (!record.PLAT) throw new Error('Le nom du plat est obligatoire.');
    upsert_('ideesPlats', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function deleteDishIdea(id) {
  return locked_(function () {
    ensureSetup_();
    const record = findById_('ideesPlats', id);
    if (!record) return { ok: true };
    record.STATUT = 'SUPPRIME';
    record.MAJ = new Date();
    upsert_('ideesPlats', record);
    bumpRevision_();
    return { ok: true };
  });
}

function planDishIdea(item) {
  return locked_(function () {
    ensureSetup_();
    const idea = findById_('ideesPlats', item.id);
    if (!idea) throw new Error('Plat à tester introuvable.');
    if (!item.date) throw new Error('La date est obligatoire.');
    const meal = item.meal === 'MIDI' ? 'MIDI' : 'SOIR';
    const now = new Date();
    const existing = rows_('menus').filter(function (x) {
      return isoDate_(x.DATE) === item.date && x.REPAS === meal && x.STATUT !== 'ARCHIVE';
    })[0];
    const menu = {
      ID: existing ? existing.ID : Utilities.getUuid(),
      DATE: parseDate_(item.date),
      REPAS: meal,
      PERSONNES: existing ? existing.PERSONNES : '',
      MENU: idea.PLAT,
      URL: idea.URL,
      COMMENTAIRE: clean_(item.comment || idea.COMMENTAIRE),
      CHOUCHOU: existing ? existing.CHOUCHOU || 'À définir' : 'À définir',
      LOULOU: existing ? existing.LOULOU || 'À définir' : 'À définir',
      DEPLACEMENT: existing ? existing.DEPLACEMENT : '',
      STATUT: 'PLANIFIE',
      MAJ: now,
      EVALUATION: clean_(item.rating || idea.EVALUATION || (existing && existing.EVALUATION)),
      ETAT_MENU: normalizeMenuState_(existing && existing.ETAT_MENU)
    };
    upsert_('menus', menu);
    idea.STATUT = 'PLANIFIE';
    idea.DATE_PLANIFIEE = parseDate_(item.date);
    idea.REPAS_PLANIFIE = meal;
    idea.EVALUATION = clean_(item.rating || idea.EVALUATION);
    idea.MAJ = now;
    upsert_('ideesPlats', idea);
    bumpRevision_();
    return { idea: serializeRecord_(idea), menu: serializeRecord_(menu) };
  });
}

function saveMenu(item) {
  return locked_(function () {
    ensureSetup_();
    const now = new Date();
    const meal = item.meal === 'MIDI' ? 'MIDI' : 'SOIR';
    const record = {
      ID: item.id || Utilities.getUuid(),
      DATE: parseDate_(item.date),
      REPAS: meal,
      PERSONNES: clean_(item.people),
      MENU: clean_(item.menu),
      URL: clean_(item.url),
      COMMENTAIRE: clean_(item.comment),
      CHOUCHOU: clean_(item.chouchou || 'À définir'),
      LOULOU: clean_(item.loulou || 'À définir'),
      DEPLACEMENT: '',
      STATUT: 'PLANIFIE',
      MAJ: now,
      EVALUATION: normalizeRating_(item.rating),
      ETAT_MENU: normalizeMenuState_(item.menuState)
    };
    if (!item.date) throw new Error('La date est obligatoire.');
    const activeMenus = rows_('menus').filter(function (x) {
      return x.STATUT !== 'ARCHIVE';
    });
    const existingAtTarget = activeMenus.filter(function (x) {
      return isoDate_(x.DATE) === item.date && x.REPAS === meal;
    })[0];
    if (!item.id && existingAtTarget) record.ID = existingAtTarget.ID;
    if (item.id && existingAtTarget && String(existingAtTarget.ID) !== String(item.id)) deleteById_('menus', existingAtTarget.ID);
    upsert_('menus', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function saveMenuIngredients(menuId, items) {
  return locked_(function () {
    ensureSetup_();
    const id = clean_(menuId);
    if (!id) throw new Error('Menu obligatoire pour les ingrédients.');
    const now = new Date();
    rows_('menuIngredients').forEach(function (x) {
      if (String(x.MENU_ID) === String(id)) deleteById_('menuIngredients', x.ID);
    });
    const records = (items || []).map(function (item) {
      return {
        ID: clean_(item.id) || Utilities.getUuid(),
        MENU_ID: id,
        CATALOGUE_ID: clean_(item.catalogueId),
        PRODUIT: clean_(item.product),
        QUANTITE: clean_(item.quantity),
        UNITE: clean_(item.unit),
        COMMENTAIRE: clean_(item.comment),
        STATUT: normalizeIngredientStatus_(item.status),
        COURSE_ID: clean_(item.courseId),
        RESTE_ID: clean_(item.leftoverId),
        MAJ: now
      };
    }).filter(function (x) { return x.PRODUIT; });
    appendRecords_('menuIngredients', records);
    bumpRevision_();
    return records.map(serializeRecord_);
  });
}

function addMenuIngredientsToCourses(menuId) {
  return locked_(function () {
    ensureSetup_();
    const id = clean_(menuId);
    const menu = findById_('menus', id);
    if (!menu) throw new Error('Menu introuvable.');
    const now = new Date();
    const records = rows_('menuIngredients').filter(function (x) {
      return String(x.MENU_ID) === String(id) && ['A_ACHETER', 'A_VERIFIER_STOCK'].indexOf(x.STATUT) >= 0;
    });
    const results = records.map(function (item) {
      let course = item.COURSE_ID ? findById_('courses', item.COURSE_ID) : null;
      const catalogue = item.CATALOGUE_ID ? findById_('catalogue', item.CATALOGUE_ID) : null;
      if (!course) {
        course = {
          ID: Utilities.getUuid(),
          CATALOGUE_ID: item.CATALOGUE_ID || '',
          PRODUIT: item.PRODUIT || (catalogue && catalogue.PRODUIT) || '',
          TYPOLOGIE: (catalogue && (catalogue.RAYON || catalogue.TYPOLOGIE)) || 'Autres',
          MAGASIN_DEFAUT: (catalogue && catalogue.MAGASIN_DEFAUT) || '',
          MAGASIN_ACTUEL: (catalogue && catalogue.MAGASIN_DEFAUT) || '',
          RAYON: (catalogue && catalogue.RAYON) || '',
          QUANTITE: item.QUANTITE || (catalogue && catalogue.QUANTITE_DEFAUT) || '1',
          CONDITIONNEMENT: item.UNITE || (catalogue && catalogue.CONDITIONNEMENT) || '',
          COMMENTAIRE: item.COMMENTAIRE || menu.MENU || '',
          STATUT: item.STATUT,
          AJOUTE_LE: now,
          ACHETE_LE: '',
          MAJ: now,
          MENUS_IDS: id
        };
      } else {
        course.QUANTITE = item.QUANTITE || course.QUANTITE;
        course.CONDITIONNEMENT = item.UNITE || course.CONDITIONNEMENT;
        course.COMMENTAIRE = item.COMMENTAIRE || course.COMMENTAIRE;
        course.STATUT = item.STATUT;
        course.MENUS_IDS = mergeIds_(course.MENUS_IDS, id);
        course.MAJ = now;
      }
      upsert_('courses', course);
      item.COURSE_ID = course.ID;
      item.MAJ = now;
      upsert_('menuIngredients', item);
      return serializeRecord_(course);
    });
    if (records.length) bumpRevision_();
    return { ok: true, courses: results, ingredients: rows_('menuIngredients').filter(function (x) { return String(x.MENU_ID) === String(id); }).map(serializeRecord_) };
  });
}

function updateMenuState(id, state) {
  return locked_(function () {
    ensureSetup_();
    const record = findById_('menus', id);
    if (!record) return { ok: false, missing: true, id: id };
    record.ETAT_MENU = normalizeMenuState_(state);
    record.MAJ = new Date();
    upsert_('menus', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function archiveMenu(id) {
  return locked_(function () {
    ensureSetup_();
    const record = findById_('menus', id);
    if (!record) return { ok: true };
    record.STATUT = 'ARCHIVE';
    record.MAJ = new Date();
    upsert_('menus', record);
    bumpRevision_();
    return { ok: true };
  });
}

function saveLeftover(item) {
  return locked_(function () {
    const now = new Date();
    const record = {
      ID: item.id || Utilities.getUuid(),
      NOM: clean_(item.name),
      QUANTITE: clean_(item.quantity),
      UNITE: clean_(item.unit),
      DATE_LIMITE: item.expiry ? parseDate_(item.expiry) : '',
      URGENCE: item.urgency === 'URGENT' ? 'URGENT' : 'NON_URGENT',
      DATE_UTILISATION: item.useDate ? parseDate_(item.useDate) : '',
      REPAS: clean_(item.meal),
      STATUT: leftoverStatus_(item),
      COMMENTAIRE: clean_(item.comment),
      MAJ: now,
      CATEGORIE: normalizeLeftoverCategory_(item.category)
    };
    if (!record.NOM) throw new Error('Le nom du reste est obligatoire.');
    upsert_('restes', record);
    bumpRevision_();
    return serializeRecord_(record);
  });
}

function saveLeftovers(items) {
  return locked_(function () {
    const now = new Date();
    const saved = (items || []).map(function (item) {
      const record = {
        ID: item.id || Utilities.getUuid(),
        NOM: clean_(item.name),
        QUANTITE: clean_(item.quantity),
        UNITE: clean_(item.unit),
        DATE_LIMITE: item.expiry ? parseDate_(item.expiry) : '',
        URGENCE: item.urgency === 'URGENT' ? 'URGENT' : 'NON_URGENT',
        DATE_UTILISATION: item.useDate ? parseDate_(item.useDate) : '',
        REPAS: clean_(item.meal),
        STATUT: leftoverStatus_(item),
        COMMENTAIRE: clean_(item.comment),
        MAJ: now,
        CATEGORIE: normalizeLeftoverCategory_(item.category)
      };
      if (record.NOM) upsert_('restes', record);
      return serializeRecord_(record);
    });
    bumpRevision_();
    return { ok: true, saved: saved };
  });
}

function consumeLeftover(id) {
  return locked_(function () {
    const record = findById_('restes', id);
    if (record) {
      record.STATUT = 'CONSOMME';
      record.MAJ = new Date();
      upsert_('restes', record);
    }
    bumpRevision_();
    return { ok: true };
  });
}

function deleteLeftover(id) {
  return locked_(function () {
    ensureSetup_();
    deleteById_('restes', id);
    bumpRevision_();
    return { ok: true };
  });
}

function importLegacySpreadsheet(spreadsheetId) {
  return locked_(function () {
    ensureSetup_();
    const source = SpreadsheetApp.openById(String(spreadsheetId).trim());
    const stats = { catalogue: 0, menus: 0, historiqueMenus: 0, restes: 0, recettes: 0, ideesPlats: 0 };
    const buffers = { catalogue: [], menus: [], historiqueMenus: [], restes: [], recettes: [], ideesPlats: [] };
    ['catalogue', 'menus', 'historiqueMenus', 'restes', 'recettes'].forEach(clearData_);
    importCoursesBase_(source, stats, buffers);
    importMenus_(source, 'Menus semaine', 'menus', stats, buffers);
    importMenus_(source, 'Menus - Historique', 'historiqueMenus', stats, buffers);
    importRecipes_(source, stats, buffers);
    Object.keys(buffers).forEach(function (key) {
      appendRecords_(key, buffers[key]);
    });
    bumpRevision_();
    return stats;
  });
}

function importLegacyData() {
  return importLegacySpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId());
}

function importLegacyRecipesToIdeas(spreadsheetId) {
  return locked_(function () {
    ensureSetup_();
    const source = SpreadsheetApp.openById(String(spreadsheetId || SpreadsheetApp.getActiveSpreadsheet().getId()).trim());
    const stats = { ideesPlats: 0 };
    const buffers = { ideesPlats: [] };
    importRecipes_(source, stats, buffers);
    appendRecords_('ideesPlats', buffers.ideesPlats);
    if (buffers.ideesPlats.length) bumpRevision_();
    return stats;
  });
}

function importCoursesBase_(source, stats, buffers) {
  const sh = source.getSheetByName('Courses base');
  if (!sh) return;
  const values = sh.getDataRange().getDisplayValues();
  const seen = {};
  [[0, 2, 3, 4], [7, 9, 10, 11]].forEach(function (block) {
    let category = 'Autres';
    for (let r = 2; r < values.length; r++) {
      const catCell = clean_(values[r][block[0]]);
      if (catCell) category = catCell.replace(/^\d+\.\s*/, '').trim();
      const product = clean_(values[r][block[1]]);
      if (!product || product === 'ITEM') continue;
      const key = product.toLowerCase() + '|' + category.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      buffers.catalogue.push({
        ID: Utilities.getUuid(), TYPOLOGIE: category, PRODUIT: product,
        MAGASIN_DEFAUT: clean_(values[r][block[2]]), RAYON: '',
        CONDITIONNEMENT: '', ACTIF: true,
        DERNIER_ACHAT: '', MAJ: new Date(),
        QUANTITE_DEFAUT: clean_(values[r][block[3]]) || '1',
        ICONE: defaultProductIcon_(product, category)
      });
      stats.catalogue++;
    }
  });
}

function importMenus_(source, sourceName, targetKey, stats, buffers) {
  const sh = source.getSheetByName(sourceName);
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  let currentDate = null;
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] instanceof Date) currentDate = values[r][0];
    const mealRaw = String(values[r][2] || '').toUpperCase();
    const meal = mealRaw === 'M' ? 'MIDI' : mealRaw === 'S' ? 'SOIR' : '';
    const menu = clean_(values[r][4]);
    if (!currentDate || !meal || !menu) continue;
    const comment = clean_(values[r][5]);
    const urlMatch = comment.match(/https?:\/\/\S+/);
    buffers[targetKey].push({
      ID: Utilities.getUuid(), DATE: currentDate, REPAS: meal,
      PERSONNES: clean_(values[r][3]), MENU: menu,
      URL: urlMatch ? urlMatch[0] : '', COMMENTAIRE: comment,
      CHOUCHOU: 'À définir', LOULOU: 'À définir', DEPLACEMENT: clean_(values[r][7]),
      STATUT: targetKey === 'menus' ? 'PLANIFIE' : 'ARCHIVE', MAJ: new Date(), EVALUATION: '', ETAT_MENU: 'A_FAIRE'
    });
    stats[targetKey]++;
  }
  if (sourceName === 'Menus semaine') {
    for (let r = 16; r < values.length; r++) {
      const name = clean_(values[r][5]);
      if (!name || name === 'RESTES' || name === 'PLATS PREPARES') continue;
      const urgency = /urgent/i.test(String(values[r][6])) ? 'URGENT' : 'OK';
      buffers.restes.push({
        ID: Utilities.getUuid(), NOM: name, QUANTITE: '', UNITE: '',
        DATE_LIMITE: '', URGENCE: urgency,
        DATE_UTILISATION: values[r][7] instanceof Date ? values[r][7] : '',
        REPAS: clean_(values[r][8]), STATUT: values[r][7] ? 'PLANIFIE' : 'DISPONIBLE',
        COMMENTAIRE: '', MAJ: new Date(), CATEGORIE: 'Plats préparés'
      });
      stats.restes++;
    }
  }
}

function importRecipes_(source, stats, buffers) {
  const sh = source.getSheetByName('Recettes');
  if (!sh) return;
  const values = sh.getDataRange().getDisplayValues();
  const seenIdeas = {};
  rows_('ideesPlats').forEach(function (x) {
    seenIdeas[recipeIdeaKey_(x.PLAT, x.URL)] = true;
  });
  for (let r = 0; r < values.length; r++) {
    const name = clean_(values[r][0]);
    if (!name) continue;
    const raw = clean_(values[r][1]);
    if (r === 0 && /^plat|^nom|^recette/i.test(name) && /url|lien|recette/i.test(raw)) continue;
    const match = raw.match(/https?:\/\/\S+/);
    const url = match ? match[0] : raw;
    if (buffers.recettes) {
      buffers.recettes.push({
        ID: Utilities.getUuid(), NOM: name, URL: url,
        COMMENTAIRE: match ? raw.replace(match[0], '').trim() : '', MAJ: new Date()
      });
      stats.recettes++;
    }
    const ideaKey = recipeIdeaKey_(name, url);
    if (buffers.ideesPlats && !seenIdeas[ideaKey]) {
      seenIdeas[ideaKey] = true;
      buffers.ideesPlats.push({
        ID: Utilities.getUuid(), PLAT: name, URL: url, COMMENTAIRE: '',
        STATUT: 'A_TESTER', DATE_PLANIFIEE: '', REPAS_PLANIFIE: '',
        MAJ: new Date(), EVALUATION: ''
      });
      stats.ideesPlats++;
    }
  }
}

function recipeIdeaKey_(name, url) {
  return clean_(name).toLowerCase() + '|' + clean_(url).toLowerCase();
}

function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(sheetName_('settings')) || getSetting_('APP_VERSION') !== '2.4.0') setupApp_();
}

function sheetName_(key) {
  const names = {
    settings: 'APP_PARAMETRES', catalogue: 'APP_CATALOGUE', courses: 'APP_COURSES',
    menus: 'APP_MENUS', restes: 'APP_RESTES', recettes: 'APP_RECETTES',
    historiqueMenus: 'APP_HISTORIQUE_MENUS', references: 'APP_REFERENTIELS',
    ideesPlats: 'APP_IDEES_PLATS', legumesSaison: 'APP_LEGUMES_SAISON',
    menuIngredients: 'APP_MENU_INGREDIENTS'
  };
  return names[key];
}

function ensureSheetSize_(sh, minRows, minCols) {
  const rows = sh.getMaxRows();
  if (rows < minRows) {
    if (rows <= 0) sh.insertRows(1, minRows);
    else sh.insertRowsAfter(rows, minRows - rows);
  }
  const cols = sh.getMaxColumns();
  if (cols < minCols) {
    if (cols <= 0) sh.insertColumns(1, minCols);
    else sh.insertColumnsAfter(cols, minCols - cols);
  }
}

function rows_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_(key));
  if (!sh || sh.getLastRow() < 2) return [];
  const headers = APP.sheets[key];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = serialize_(row[i]); });
    return obj;
  }).filter(function (obj) { return obj.ID || key === 'settings'; });
}

function getReferences_() {
  const allReferenceRows = rows_('references');
  const allRefs = allReferenceRows.filter(isReferenceActive_);
  const deleted = {
    stores: deletedReferenceValues_(allRefs, 'MAGASIN'),
    aisles: deletedReferenceValues_(allRefs, 'RAYON'),
    childLocations: deletedReferenceValues_(allRefs, 'LIEU_ENFANT'),
    leftoverCategories: deletedReferenceValues_(allRefs, 'RESTE_CATEGORIE')
  };
  const refs = dedupeReferences_(allRefs).filter(function (x) {
    if (String(x.TYPE || '').indexOf('SUPPRIME_') === 0) return false;
    return !isDeletedReferenceValue_(deleted, x.TYPE, x.VALEUR);
  });
  const storeAisleOrders = {};
  refs.filter(function (x) { return x.TYPE === 'RAYON_MAGASIN'; }).forEach(function (x) {
    const parsed = parseStoreAisleReference_(x);
    if (!parsed.store || !parsed.aisle) return;
    if (isDeletedReferenceValue_(deleted, 'MAGASIN', parsed.store) || isDeletedReferenceValue_(deleted, 'RAYON', parsed.aisle)) return;
    if (!storeAisleOrders[parsed.store]) storeAisleOrders[parsed.store] = [];
    storeAisleOrders[parsed.store].push({
      ID: x.ID,
      STORE: parsed.store,
      AISLE: parsed.aisle,
      ORDRE: x.ORDRE
    });
  });
  const hiddenStoreAisles = {};
  allReferenceRows.filter(function (x) {
    return x.TYPE === 'RAYON_MAGASIN' && !isReferenceActive_(x);
  }).forEach(function (x) {
    const parsed = parseStoreAisleReference_(x);
    if (!parsed.store || !parsed.aisle) return;
    if (isDeletedReferenceValue_(deleted, 'MAGASIN', parsed.store) || isDeletedReferenceValue_(deleted, 'RAYON', parsed.aisle)) return;
    if (!hiddenStoreAisles[parsed.store]) hiddenStoreAisles[parsed.store] = [];
    hiddenStoreAisles[parsed.store].push(parsed.aisle);
  });
  Object.keys(storeAisleOrders).forEach(function (store) {
    storeAisleOrders[store].sort(function (a, b) {
      return (Number(a.ORDRE) || 9999) - (Number(b.ORDRE) || 9999) || String(a.AISLE).localeCompare(String(b.AISLE));
    });
  });
  return {
    stores: refs.filter(function (x) { return x.TYPE === 'MAGASIN'; }).sort(referenceSort_),
    aisles: refs.filter(function (x) { return x.TYPE === 'RAYON'; }).sort(referenceSort_),
    childLocations: refs.filter(function (x) { return x.TYPE === 'LIEU_ENFANT'; }).sort(referenceSort_),
    leftoverCategories: refs.filter(function (x) { return x.TYPE === 'RESTE_CATEGORIE'; }).sort(referenceSort_),
    storeAisleOrders: storeAisleOrders,
    hiddenStoreAisles: hiddenStoreAisles,
    deleted: deleted
  };
}

function deletedReferenceValues_(refs, type) {
  const markerType = deletedReferenceType_(type);
  return refs.filter(function (x) { return x.TYPE === markerType; }).map(function (x) { return clean_(x.VALEUR); }).filter(Boolean);
}

function isDeletedReferenceValue_(deleted, type, value) {
  const key = type === 'RAYON' ? 'aisles' : type === 'LIEU_ENFANT' ? 'childLocations' : type === 'RESTE_CATEGORIE' ? 'leftoverCategories' : 'stores';
  const raw = clean_(value).toLowerCase();
  return (deleted[key] || []).some(function (x) { return clean_(x).toLowerCase() === raw; });
}

function isReferenceActive_(record) {
  return record && record.ACTIF !== false && record.ACTIF !== 'FALSE' && record.ACTIF !== 'False' && record.ACTIF !== 'false';
}

function dedupeReferences_(refs) {
  const sorted = refs.slice().sort(referenceSort_);
  const seen = {};
  return sorted.filter(function (x) {
    const key = String(x.TYPE || '') + '|' + clean_(x.VALEUR).toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function parseStoreAisleReference_(record) {
  const raw = String(record && record.VALEUR || '');
  const delimiter = '\u001f';
  const index = raw.indexOf(delimiter);
  if (index < 0) return { store: '', aisle: '' };
  return {
    store: raw.slice(0, index),
    aisle: raw.slice(index + delimiter.length)
  };
}

function seedReferences_() {
  const existing = rows_('references').filter(isReferenceActive_);
  const seen = {};
  existing.forEach(function (x) { seen[x.TYPE + '|' + String(x.VALEUR).toLowerCase()] = true; });
  const records = [];
  ['À définir', 'Maison', 'Extérieur'].forEach(function (value) {
    addSeedReference_(records, seen, 'LIEU_ENFANT', value);
  });
  ['Plats préparés', 'Fruits', 'Légumes', 'Crèmerie', 'Viande & Co', 'Surgelés', 'Epicerie'].forEach(function (value) {
    addSeedReference_(records, seen, 'RESTE_CATEGORIE', value);
  });
  rows_('catalogue').forEach(function (x) {
    addSeedReference_(records, seen, 'MAGASIN', x.MAGASIN_DEFAUT);
    addSeedReference_(records, seen, 'RAYON', x.RAYON);
  });
  appendRecords_('references', records);
}

function addSeedReference_(records, seen, type, value) {
  value = clean_(value);
  if (!value) return;
  const key = type + '|' + value.toLowerCase();
  if (seen[key]) return;
  seen[key] = true;
  records.push({
    ID: Utilities.getUuid(), TYPE: type, VALEUR: value, ACTIF: true,
    ORDRE: records.filter(function (x) { return x.TYPE === type; }).length + 1,
    MAJ: new Date()
  });
}

function seedSeasonalVegetables_() {
  clearData_('legumesSaison');
  const now = new Date();
  const byMonth = {
    1: ['Ail', 'Betterave', 'Carotte', 'Céleri-branche', 'Céleri-rave', 'Chou', 'Chou blanc', 'Chou de Bruxelles', 'Chou frisé', 'Chou rouge', 'Courge', 'Crosne', 'Endive', 'Épinard', 'Frisée', 'Mâche', 'Navet', 'Oignon', 'Panais', 'Poireau', 'Pomme de terre de conservation', 'Salsifi', 'Topinambour'],
    2: ['Ail', 'Artichaut', 'Asperge', 'Aubergine', 'Betterave', 'Blette', 'Carotte', 'Chou-fleur', 'Concombre', 'Courgette', 'Épinard', 'Laitue', 'Navet', 'Oignon', 'Petit pois', 'Pomme de terre primeur', 'Radis'],
    3: ['Ail', 'Asperge', 'Betterave', 'Blette', 'Carotte', 'Céleri-rave', 'Chou', 'Chou de Bruxelles', 'Chou-fleur', 'Crosne', 'Endive', 'Épinard', 'Frisée', 'Navet', 'Oignon', 'Panais', 'Poireau', 'Pomme de terre de conservation', 'Radis', 'Salsifi', 'Topinambour'],
    4: ['Ail', 'Artichaut', 'Asperge', 'Betterave', 'Blette', 'Carotte', 'Chou-fleur', 'Concombre', 'Endive', 'Épinard', 'Frisée', 'Laitue', 'Navet', 'Oignon', 'Petit pois', 'Poireau', 'Pomme de terre primeur', 'Radis'],
    5: ['Ail', 'Artichaut', 'Asperge', 'Aubergine', 'Betterave', 'Blette', 'Carotte', 'Chou-fleur', 'Concombre', 'Courgette', 'Épinard', 'Laitue', 'Navet', 'Oignon', 'Petit pois', 'Pomme de terre primeur', 'Radis'],
    6: ['Ail', 'Artichaut', 'Asperge', 'Aubergine', 'Betterave', 'Blette', 'Brocoli', 'Carotte', 'Chou romanesco', 'Concombre', 'Courgette', 'Épinard', 'Fenouil', 'Haricot vert', 'Laitue', 'Navet', 'Oignon', 'Petit pois', 'Poivron', 'Pomme de terre primeur', 'Radis'],
    7: ['Ail', 'Artichaut', 'Asperge', 'Aubergine', 'Betterave', 'Blette', 'Brocoli', 'Carotte', 'Céleri-branche', 'Chou romanesco', 'Concombre', 'Courgette', 'Épinard', 'Fenouil', 'Haricot vert', 'Laitue', 'Oignon', 'Petit pois', 'Poivron', 'Pomme de terre primeur', 'Radis'],
    8: ['Ail', 'Artichaut', 'Aubergine', 'Betterave', 'Blette', 'Brocoli', 'Carotte', 'Céleri-branche', 'Chou blanc', 'Chou romanesco', 'Chou rouge', 'Concombre', 'Courge', 'Courgette', 'Épinard', 'Fenouil', 'Frisée', 'Haricot vert', 'Laitue', 'Oignon', 'Poivron', 'Pomme de terre primeur', 'Radis'],
    9: ['Ail', 'Artichaut', 'Aubergine', 'Betterave', 'Blette', 'Brocoli', 'Carotte', 'Céleri-branche', 'Chou', 'Chou blanc', 'Chou de Bruxelles', 'Chou frisé', 'Chou romanesco', 'Chou rouge', 'Chou-fleur', 'Concombre', 'Courge', 'Courgette', 'Épinard', 'Fenouil', 'Frisée', 'Haricot vert', 'Laitue', 'Oignon', 'Panais', 'Patate douce', 'Poireau', 'Poivron', 'Pomme de terre de conservation', 'Potiron', 'Radis'],
    10: ['Ail', 'Aubergine', 'Betterave', 'Blette', 'Brocoli', 'Carotte', 'Céleri-branche', 'Céleri-rave', 'Chou', 'Chou blanc', 'Chou de Bruxelles', 'Chou frisé', 'Chou rouge', 'Chou-fleur', 'Concombre', 'Courge', 'Courgette', 'Echalote', 'Endive', 'Épinard', 'Fenouil', 'Frisée', 'Haricot vert', 'Laitue', 'Navet', 'Oignon', 'Panais', 'Patate douce', 'Poireau', 'Pomme de terre de conservation', 'Potiron', 'Radis', 'Rutabaga', 'Salsifi', 'Topinambour'],
    11: ['Ail', 'Betterave', 'Brocoli', 'Cardon', 'Carotte', 'Céleri-branche', 'Céleri-rave', 'Chou', 'Chou blanc', 'Chou de Bruxelles', 'Chou frisé', 'Chou rouge', 'Chou-fleur', 'Citrouille', 'Courge', 'Crosne', 'Echalote', 'Endive', 'Épinard', 'Fenouil', 'Frisée', 'Mâche', 'Navet', 'Oignon', 'Panais', 'Poireau', 'Pomme de terre de conservation', 'Potiron', 'Radis', 'Rutabaga', 'Salsifi', 'Topinambour'],
    12: ['Ail', 'Betterave', 'Carotte', 'Céleri-branche', 'Céleri-rave', 'Chou', 'Chou blanc', 'Chou de Bruxelles', 'Chou frisé', 'Chou rouge', 'Courge', 'Crosne', 'Echalote', 'Endive', 'Épinard', 'Frisée', 'Mâche', 'Navet', 'Oignon', 'Panais', 'Poireau', 'Pomme de terre de conservation', 'Potiron', 'Radis', 'Rutabaga', 'Salsifi', 'Topinambour']
  };
  const records = [];
  Object.keys(byMonth).forEach(function (month) {
    byMonth[month].forEach(function (item) {
      records.push({ ID: month + '|' + item, ITEM: item, MOIS: Number(month), MAJ: now });
    });
  });
  appendRecords_('legumesSaison', records);
}

function nextReferenceOrder_(type) {
  return rows_('references').filter(function (x) {
    return x.TYPE === type && x.ACTIF !== false && x.ACTIF !== 'FALSE';
  }).length + 1;
}

function referenceSort_(a, b) {
  return (Number(a.ORDRE) || 9999) - (Number(b.ORDRE) || 9999) || String(a.VALEUR).localeCompare(String(b.VALEUR));
}

function findById_(key, id) {
  return rows_(key).filter(function (x) { return String(x.ID) === String(id); })[0] || null;
}

function syncOpenCoursesFromCatalogue_(catalogue) {
  rows_('courses').forEach(function (x) {
    if (String(x.CATALOGUE_ID) !== String(catalogue.ID)) return;
    if (['A_DECIDER', 'A_VERIFIER', 'A_ACHETER', 'A_VERIFIER_STOCK', 'IGNORE'].indexOf(x.STATUT) < 0) return;
    x.PRODUIT = catalogue.PRODUIT;
    x.TYPOLOGIE = catalogue.TYPOLOGIE;
    x.MAGASIN_DEFAUT = catalogue.MAGASIN_DEFAUT;
    x.MAGASIN_ACTUEL = catalogue.MAGASIN_DEFAUT;
    x.RAYON = catalogue.RAYON;
    x.QUANTITE = catalogue.QUANTITE_DEFAUT || x.QUANTITE || '1';
    x.CONDITIONNEMENT = catalogue.CONDITIONNEMENT;
    x.MAJ = new Date();
    upsert_('courses', x);
  });
}

function upsert_(key, record) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_(key));
  const headers = APP.sheets[key];
  let rowNumber = 0;
  if (record.ID && sh.getLastRow() >= 2) {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(record.ID)) { rowNumber = i + 2; break; }
    }
  }
  const values = headers.map(function (h) {
    const v = record[h];
    if (h.indexOf('DATE') >= 0 || h.endsWith('_LE') || h === 'MAJ' || h === 'DERNIER_ACHAT') {
      return v ? parseDate_(v) : '';
    }
    return v === undefined || v === null ? '' : v;
  });
  if (rowNumber) sh.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
  else sh.appendRow(values);
}

function appendRecords_(key, records) {
  if (!records.length) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_(key));
  const headers = APP.sheets[key];
  const values = records.map(function (record) {
    return headers.map(function (h) {
      const v = record[h];
      if (h.indexOf('DATE') >= 0 || h.endsWith('_LE') || h === 'MAJ' || h === 'DERNIER_ACHAT') {
        return v ? parseDate_(v) : '';
      }
      return v === undefined || v === null ? '' : v;
    });
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

function clearData_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_(key));
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, APP.sheets[key].length).clearContent();
  }
}

function deleteById_(key, id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_(key));
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
  }
}

function setSetting_(key, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_('settings'));
  const data = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function getSetting_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_('settings'));
  if (!sh || sh.getLastRow() < 2) return '';
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  const row = data.filter(function (x) { return x[0] === key; })[0];
  return row ? row[1] : '';
}

function bumpRevision_() { setSetting_('REVISION', String(Date.now())); }
function getRevision_() { return getSetting_('REVISION') || '0'; }
function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function clean_(value) { return String(value || '').replace(/\s+/g, ' ').trim(); }
function normalizeRating_(value) {
  const rating = Number(value);
  return rating >= 1 && rating <= 5 ? rating : '';
}
function normalizeMenuState_(value) {
  const raw = clean_(value).toUpperCase();
  return raw === 'EN_COURS' || raw === 'OK' ? raw : 'A_FAIRE';
}
function normalizeCourseStatus_(value) {
  const raw = clean_(value).toUpperCase();
  return raw === 'A_VERIFIER_STOCK' ? 'A_VERIFIER_STOCK' : 'A_ACHETER';
}
function normalizeIngredientStatus_(value) {
  const raw = clean_(value).toUpperCase();
  return ['A_VERIFIER_STOCK', 'A_ACHETER', 'EN_STOCK', 'IGNORE'].indexOf(raw) >= 0 ? raw : 'A_VERIFIER_STOCK';
}
function mergeIds_(current, id) {
  const ids = String(current || '').split(',').map(function (x) { return clean_(x); }).filter(Boolean);
  if (id && ids.indexOf(String(id)) < 0) ids.push(String(id));
  return ids.join(',');
}
function isActiveCourse_(record) {
  const status = clean_(record && record.STATUT).toUpperCase();
  return status === 'A_ACHETER' || status === 'A_VERIFIER_STOCK';
}
function isVisibleLeftover_(record) {
  const status = clean_(record && record.STATUT).toUpperCase();
  return status === 'DISPONIBLE' || status === 'PLANIFIE' || status === 'CONSOMME';
}
function leftoverStatus_(item) {
  if (!item || !item.useDate) return 'DISPONIBLE';
  return clean_(item.status).toUpperCase() === 'CONSOMME' ? 'CONSOMME' : 'PLANIFIE';
}
function normalizeLeftoverCategory_(value) {
  return clean_(value);
}
function parseDate_(value) {
  if (value instanceof Date) return value;
  if (!value) return '';
  const parts = String(value).slice(0, 10).split('-');
  if (parts.length === 3) return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12);
  return new Date(value);
}
function isoDate_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(d, 'Europe/Paris', 'yyyy-MM-dd');
}
function serialize_(value) {
  return value instanceof Date ? isoDate_(value) : value;
}
