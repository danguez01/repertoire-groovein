/**
 * Groove In — Apps Script Web App
 * Pont entre l'application HTML "Répertoire" et le Google Sheet.
 *
 * Endpoints :
 *   GET  ?action=list                        → renvoie toutes les chansons (JSON)
 *   POST { action: "update_song", id, fields } → met à jour les champs d'une chanson
 *   POST { action: "add_song", song }        → ajoute une chanson, renvoie son id
 *   POST { action: "delete_song", id }       → supprime une chanson
 *   POST { action: "ping" }                  → test de connexion
 *
 * Le client envoie les POST avec Content-Type: text/plain;charset=utf-8 pour
 * éviter le préflight CORS (Apps Script ne gère pas OPTIONS proprement).
 */

const SHEET_NAME = "Répertoire";

// Mapping en-tête de colonne (Sheet) ↔ champ JSON (côté app)
const FIELD_MAP = {
  "ID":                "id",
  "Titre":             "title",
  "Doublon":           "_dup",
  "Style":             "style",
  "Genre":             "genre",
  "Rythme":            "rythme",
  "BPM":               "bpm",
  "Tonalité (auto)":   "key_detected",
  "Tonalité (jouée)":  "key_played",
  "Capo":              "capo",
  "Durée":             "duree",
  "Tags concert":      "tags",
  "Niveau":            "niveau",
  "Accords utilisés":  "chords_str",
  "Nb accords":        "n_chords",
  "Nb lignes":         "n_lines",
  "Intro":             "intros",
  "Couplets":          "verses",
  "Pré-refrain":       "pre_chorus",
  "Refrain":           "choruses",
  "Bridge":            "bridges",
  "Body complet":      "body_full",
  "Notes":             "notes"
};

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Feuille "' + SHEET_NAME + '" introuvable.');
  return sh;
}

function readHeaders_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function readAllSongs_() {
  const sh = getSheet_();
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const songs = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && !row[1]) continue; // saute les lignes vides
    const song = {};
    headers.forEach((h, j) => {
      const key = FIELD_MAP[h] || h;
      song[key] = row[j];
    });
    // String "C, G, Am, F" → array
    if (song.chords_str) {
      song.chords = String(song.chords_str)
        .split(/[,;]/).map(s => s.trim()).filter(Boolean);
    } else {
      song.chords = [];
    }
    delete song.chords_str;
    delete song._dup;
    songs.push(song);
  }
  return songs;
}

function findRowById_(sh, id) {
  if (sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === Number(id)) return i + 2;
  }
  return -1;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "list";
    if (action === "list") {
      const songs = readAllSongs_();
      return jsonOut_({ ok: true, songs: songs, n: songs.length });
    }
    if (action === "setlists") {
      return jsonOut_(Object.assign({ ok: true }, readSetlists_()));
    }
    if (action === "ping") {
      return jsonOut_({ ok: true, msg: "pong", sheet: getSheet_().getName() });
    }
    return jsonOut_({ ok: false, error: "action GET inconnue: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut_({ ok: false, error: "Corps POST manquant." });
    }
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === "ping") {
      return jsonOut_({ ok: true, msg: "pong" });
    }
    if (action === "update_song") {
      return jsonOut_(updateSong_(body.id, body.fields || {}));
    }
    if (action === "add_song") {
      return jsonOut_(addSong_(body.song || {}));
    }
    if (action === "delete_song") {
      return jsonOut_(deleteSong_(body.id));
    }
    if (action === "save_setlists") {
      return jsonOut_(saveSetlists_(body.setlists || {}, body.currentSetlistId || ""));
    }
    return jsonOut_({ ok: false, error: "action POST inconnue: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

// ===== SETLISTS (synchronisation multi-appareils) =====
const APP_DATA_SHEET = "_AppData";
const KEY_SETLISTS = "setlists";
const KEY_CURRENT_SETLIST = "currentSetlistId";

function getOrCreateAppDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(APP_DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(APP_DATA_SHEET);
    sh.getRange("A1:B1").setValues([["Clé", "Valeur"]]).setFontWeight("bold");
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 600);
    sh.hideSheet(); // optionnel : la masquer pour ne pas polluer l'interface
  }
  return sh;
}

function readKVPair_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return "";
  const data = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1] || "";
  }
  return "";
}

function writeKVPair_(sh, key, value) {
  const last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sh.getRange(i + 2, 2).setValue(value);
        return;
      }
    }
  }
  sh.appendRow([key, value]);
}

function readSetlists_() {
  const sh = getOrCreateAppDataSheet_();
  const slRaw = readKVPair_(sh, KEY_SETLISTS);
  const cur = readKVPair_(sh, KEY_CURRENT_SETLIST);
  let setlists = {};
  if (slRaw) {
    try { setlists = JSON.parse(slRaw); } catch (e) { setlists = {}; }
  }
  return { setlists: setlists, currentSetlistId: cur || "" };
}

function saveSetlists_(setlists, currentSetlistId) {
  const sh = getOrCreateAppDataSheet_();
  writeKVPair_(sh, KEY_SETLISTS, JSON.stringify(setlists));
  writeKVPair_(sh, KEY_CURRENT_SETLIST, currentSetlistId || "");
  return { ok: true, n: Object.keys(setlists).length };
}

function updateSong_(id, fields) {
  const sh = getSheet_();
  const row = findRowById_(sh, id);
  if (row < 0) return { ok: false, error: "ID introuvable: " + id };
  const headers = readHeaders_(sh);
  for (let j = 0; j < headers.length; j++) {
    const h = headers[j];
    const key = FIELD_MAP[h];
    if (!key) continue;
    let v = null, has = false;
    // "Accords utilisés" : prendre fields.chords (array) ou fields.chords_str
    if (h === "Accords utilisés") {
      if (Array.isArray(fields.chords)) { v = fields.chords.join(", "); has = true; }
      else if (typeof fields.chords_str === "string") { v = fields.chords_str; has = true; }
    } else if (key in fields) {
      v = fields[key]; has = true;
    }
    if (has) sh.getRange(row, j + 1).setValue(v == null ? "" : v);
  }
  return { ok: true, id: id, row: row };
}

function addSong_(song) {
  const sh = getSheet_();
  const headers = readHeaders_(sh);
  const lastRow = sh.getLastRow();
  let nextId = 1;
  if (lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    nextId = ids.reduce((m, v) => Math.max(m, Number(v) || 0), 0) + 1;
  }
  song.id = nextId;
  const row = headers.map(h => {
    if (h === "Accords utilisés" && Array.isArray(song.chords)) {
      return song.chords.join(", ");
    }
    const key = FIELD_MAP[h];
    if (!key) return "";
    return song[key] != null ? song[key] : "";
  });
  sh.appendRow(row);
  return { ok: true, id: nextId };
}

function deleteSong_(id) {
  const sh = getSheet_();
  const row = findRowById_(sh, id);
  if (row < 0) return { ok: false, error: "ID introuvable: " + id };
  sh.deleteRow(row);
  return { ok: true, id: id };
}

/**
 * À exécuter UNE fois (menu "Exécuter" → "smokeTest") avant de déployer,
 * pour vérifier que le script accède bien au Sheet et trouver les bonnes colonnes.
 */
function smokeTest() {
  const songs = readAllSongs_();
  Logger.log("Songs trouvées : " + songs.length);
  if (songs.length) {
    Logger.log("Première : " + JSON.stringify(songs[0]).slice(0, 300));
  }
}
