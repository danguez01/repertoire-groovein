#!/usr/bin/env python3
"""Build the single-file HTML interface, with Drive sync support."""
import json
from datetime import datetime
from pathlib import Path

SONGS_JSON = Path(__file__).parent / "songs.json"
OUT_HTML = Path(__file__).parent.parent / "index.html"  # écrit à la racine du dépôt (servi par GitHub Pages)
BUILD_VERSION = datetime.now().strftime("%Y%m%d-%H%M")

songs = json.loads(SONGS_JSON.read_text(encoding="utf-8"))
# Add empty metadata fields so the embedded fallback has the same shape as the xlsx-loaded one
for s in songs:
    s.setdefault("style", "")
    s.setdefault("genre", "")
    s.setdefault("rythme", "")
    s.setdefault("bpm", "")
    s.setdefault("duree", "")
    s.setdefault("tags", "")
    s.setdefault("niveau", "")
    s.setdefault("notes", "")
    s.setdefault("key_played", "")
    s.setdefault("artiste", "")
    s.setdefault("spotify", "")
    s.setdefault("lien_ecoute", "")
songs_json = json.dumps(songs, ensure_ascii=False)

HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<!-- Anti-cache : force le navigateur à re-vérifier la dernière version à chaque ouverture -->
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta name="build-version" content="__BUILD_VERSION__">
<title>Répertoire Groove In</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --bg: #0f1419;
    --bg-elev: #1a2129;
    --bg-elev2: #232c36;
    --border: #2d3742;
    --text: #e8edf2;
    --text-dim: #97a3b1;
    --accent: #e07a5f;
    --accent-hover: #f29280;
    --chord: #5fb3e0;
    --section: #f4a261;
    --ok: #6bc26b;
    --warn: #f4a261;
    --err: #e07a5f;
    /* Couleurs des chanteurs (jusqu'à 4 par chanson) */
    --singer-1: #ec4899;
    --singer-2: #4fc3f7;
    --singer-3: #fbbf24;
    --singer-4: #a78bfa;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: var(--bg);
    color: var(--text);
    height: 100vh; height: 100dvh;
    overflow: hidden;
  }
  .app { display: grid; grid-template-columns: 320px 1fr 300px; height: 100vh; height: 100dvh; }
  .app.stage { grid-template-columns: 1fr; }

  /* === LEFT PANEL (song list) === */
  .left { background: var(--bg-elev); border-right: 1px solid var(--border); display: flex; flex-direction: column; overflow: hidden; }
  .left-header { padding: 14px 16px; border-bottom: 1px solid var(--border); }
  .left-header h1 { margin: 0 0 4px; font-size: 17px; color: var(--accent); letter-spacing: -0.3px; display: flex; align-items: center; justify-content: space-between; }
  .left-header h1 .icons { display: inline-flex; gap: 4px; }
  .left-header h1 .gear, .left-header h1 .add-btn { font-size: 14px; cursor: pointer; color: var(--text-dim); padding: 4px 8px; border-radius: 4px; line-height: 1; }
  .left-header h1 .gear:hover, .left-header h1 .add-btn:hover { color: var(--accent); background: var(--bg-elev2); }
  .left-header h1 .add-btn { color: var(--accent); font-weight: 700; }
  .pencil { display: inline-block; margin-left: 6px; font-size: 14px; cursor: pointer; opacity: 0.55; padding: 2px 6px; border-radius: 4px; }
  .pencil:hover { opacity: 1; background: var(--bg-elev2); color: var(--accent); }
  .pencil-body { font-size: 12px; opacity: 0.7; border: 1px solid var(--border); padding: 3px 8px; }
  .pencil-body:hover { background: var(--accent); color: var(--bg); border-color: var(--accent); }
  .save-bar { padding: 10px 16px; background: rgba(244, 162, 97, 0.12); border-top: 1px solid var(--warn); display: none; flex-direction: column; gap: 8px; }
  #drop-overlay { position: fixed; inset: 0; background: rgba(13, 18, 26, 0.92); border: 4px dashed var(--accent); display: none; align-items: center; justify-content: center; z-index: 9999; color: var(--accent); font-size: 28px; font-weight: 700; pointer-events: none; letter-spacing: 0.5px; }
  #drop-overlay.show { display: flex; }
  .status a { color: inherit; text-decoration: underline; font-weight: 600; }
  .status ol, .status ul { color: inherit; }
  .save-bar.show { display: flex; }
  .save-bar .label { font-size: 12px; color: var(--warn); font-weight: 600; }
  textarea.song-input { width: 100%; min-height: 280px; padding: 10px 12px; background: var(--bg); border: 1px solid var(--border); color: var(--text); border-radius: 5px; font-size: 12.5px; outline: none; font-family: "JetBrains Mono", "Consolas", monospace; line-height: 1.5; resize: vertical; }
  textarea.song-input:focus { border-color: var(--accent); }
  .preview-card { background: var(--bg); border: 1px solid var(--border); border-radius: 6px; padding: 12px 14px; margin-top: 12px; font-size: 12.5px; }
  .preview-card .row { display: flex; gap: 16px; flex-wrap: wrap; padding: 4px 0; }
  .preview-card .row span { color: var(--text-dim); }
  .preview-card .row strong { color: var(--text); }
  .preview-card .err { color: var(--err); }
  .left-header .subtitle { color: var(--text-dim); font-size: 12px; }
  .left-header .source-info { color: var(--text-dim); font-size: 11px; margin-top: 4px; }
  .left-header .source-info .ok { color: var(--ok); }
  .left-header .source-info .warn { color: var(--warn); }
  .connect-btn { margin-top: 6px; padding: 4px 10px; border-radius: 5px; border: 1px solid var(--border); background: var(--bg-elev2); color: var(--text-dim); font-size: 11px; cursor: pointer; width: 100%; text-align: left; font-family: inherit; transition: color 0.2s, border-color 0.2s; }
  .connect-btn:hover { border-color: var(--accent); color: var(--accent); }
  .connect-btn.ok { color: var(--ok); border-color: rgba(107,194,107,0.4); }
  .connect-btn.err { color: var(--err); border-color: rgba(224,122,95,0.4); }
  .connect-btn.err:hover { border-color: var(--err); }
  .connect-btn.loading { color: var(--warn); border-color: rgba(244,162,97,0.4); cursor: wait; }
  .filters { padding: 10px 16px; border-bottom: 1px solid var(--border); }
  .filters input, .filters select {
    width: 100%; padding: 7px 9px; background: var(--bg);
    border: 1px solid var(--border); color: var(--text);
    border-radius: 5px; font-size: 12px; margin-bottom: 6px; outline: none;
  }
  .filters input:focus, .filters select:focus { border-color: var(--accent); }
  .filters .row { display: flex; gap: 6px; }
  .filters .row > * { flex: 1; }
  .song-list { flex: 1; overflow-y: auto; padding: 4px 0; }
  .song-item { padding: 9px 40px 9px 16px; cursor: pointer; border-left: 3px solid transparent; transition: background 0.1s; position: relative; }
  .song-del { position: absolute; right: 4px; top: 50%; transform: translateY(-50%); padding: 8px 10px; font-size: 14px; line-height: 1; opacity: 0.3; border-radius: 6px; cursor: pointer; }
  .song-item:hover .song-del { opacity: 0.7; }
  .song-del:hover { opacity: 1 !important; background: rgba(224, 122, 95, 0.18); }
  @media (pointer: coarse) { .song-del { opacity: 0.5; padding: 12px 12px; } }
  .song-item:hover { background: var(--bg-elev2); }
  .song-item.active { background: var(--bg-elev2); border-left-color: var(--accent); }
  .song-item .title { font-size: 13.5px; font-weight: 500; }
  .song-item .meta { font-size: 11px; color: var(--text-dim); margin-top: 2px; }
  .song-item .badge { display: inline-block; padding: 1px 6px; border-radius: 3px; background: var(--bg); margin-right: 4px; font-family: monospace; font-size: 10px; }
  .song-item .badge.style { background: rgba(95, 179, 224, 0.15); color: var(--chord); }

  /* === CENTER PANEL === */
  .center { display: flex; flex-direction: column; overflow: hidden; position: relative; min-width: 0; }
  .center-header { padding: 14px 24px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 10px; background: var(--bg-elev); flex-wrap: wrap; }
  .song-title { font-size: 22px; font-weight: 600; flex: 1; margin: 0; letter-spacing: -0.4px; }
  .meta-pill { padding: 4px 10px; background: var(--bg); border: 1px solid var(--border); border-radius: 14px; font-size: 12px; color: var(--text-dim); }
  .meta-pill strong { color: var(--text); }
  .meta-pill.style { background: rgba(95, 179, 224, 0.1); border-color: rgba(95, 179, 224, 0.3); }
  .meta-pill.bpm { background: rgba(244, 162, 97, 0.1); border-color: rgba(244, 162, 97, 0.3); }
  .controls { display: flex; gap: 5px; }
  .btn { padding: 7px 12px; background: var(--bg-elev2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; cursor: pointer; font-size: 13px; transition: all 0.1s; }
  .btn:hover { background: var(--accent); border-color: var(--accent); }
  .btn.ghost:hover { background: var(--bg); border-color: var(--accent); color: var(--accent); }
  .btn.primary { background: var(--accent); border-color: var(--accent); color: #1a1208; font-weight: 600; }
  .btn.primary:hover { background: var(--accent-hover); border-color: var(--accent-hover); }
  .btn.icon { padding: 7px 9px; }
  .center-body { flex: 1; overflow-y: auto; padding: 32px 48px 80px; background: var(--bg); }
  .empty { display: flex; align-items: center; justify-content: center; height: 100%; color: var(--text-dim); text-align: center; flex-direction: column; gap: 12px; padding: 24px; }
  .empty .hint { font-size: 13px; max-width: 400px; line-height: 1.5; }
  pre.song-body { font-family: "JetBrains Mono", "Consolas", "Menlo", monospace; font-size: var(--song-fs, 14px); line-height: 1.55; white-space: pre; overflow-x: auto; margin: 0; color: var(--text); }
  pre.song-body .chord { color: var(--chord); font-weight: 600; }
  pre.song-body .section { color: var(--section); font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 1em; display: block; }

  /* === CHANTEURS (couleurs par ligne) === */
  pre.song-body .line { display: inline; }
  .line-singer-0 { /* defaut : tous, pas de couleur */ }
  .line-singer-1 { color: var(--singer-1) !important; }
  .line-singer-2 { color: var(--singer-2) !important; }
  .line-singer-3 { color: var(--singer-3) !important; }
  .line-singer-4 { color: var(--singer-4) !important; }

  /* Mode attribution : bandeau d'info + curseur main sur les lignes */
  .singer-banner { display: none; }
  .app.singer-assign-mode .singer-banner {
    display: flex; align-items: center; gap: 12px;
    padding: 10px 16px; background: rgba(236, 72, 153, 0.15);
    border-bottom: 1px solid var(--singer-1); color: var(--text);
    font-size: 13px; font-weight: 500;
  }
  .singer-banner .palette { display: inline-flex; gap: 6px; }
  .singer-banner .palette .swatch {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 3px 8px; border-radius: 12px;
    background: var(--bg-elev2); font-size: 11px;
  }
  .singer-banner .palette .swatch::before {
    content: ""; display: inline-block; width: 10px; height: 10px; border-radius: 50%;
  }
  .singer-banner .palette .sw-1::before { background: var(--singer-1); }
  .singer-banner .palette .sw-2::before { background: var(--singer-2); }
  .singer-banner .palette .sw-3::before { background: var(--singer-3); }
  .singer-banner .palette .sw-4::before { background: var(--singer-4); }
  .singer-banner .palette .sw-0::before { background: var(--text-dim); }
  .singer-banner .close-assign { margin-left: auto; cursor: pointer; padding: 4px 10px; }
  .app.singer-assign-mode pre.song-body .line { cursor: pointer; padding: 1px 0; border-radius: 3px; }
  .app.singer-assign-mode pre.song-body .line:hover { background: rgba(255,255,255,0.06); }

  /* Modal de configuration des chanteurs */
  .singer-config { display: flex; flex-direction: column; gap: 12px; }
  .singer-config .count-row { display: flex; gap: 6px; }
  .singer-config .count-row .btn { flex: 1; }
  .singer-config .count-row .btn.active {
    background: var(--accent); border-color: var(--accent); color: var(--bg); font-weight: 700;
  }
  .singer-name-row { display: flex; align-items: center; gap: 10px; }
  .singer-name-row .swatch { width: 18px; height: 18px; border-radius: 50%; flex-shrink: 0; }
  .singer-name-row.s-1 .swatch { background: var(--singer-1); }
  .singer-name-row.s-2 .swatch { background: var(--singer-2); }
  .singer-name-row.s-3 .swatch { background: var(--singer-3); }
  .singer-name-row.s-4 .swatch { background: var(--singer-4); }
  .singer-name-row input { flex: 1; padding: 8px 10px; background: var(--bg); border: 1px solid var(--border); color: var(--text); border-radius: 5px; font-size: 13px; }

  .extra-meta { display: flex; gap: 16px; flex-wrap: wrap; padding: 0 0 24px; color: var(--text-dim); font-size: 12px; border-bottom: 1px dashed var(--border); margin-bottom: 24px; }
  .extra-meta div { display: flex; flex-direction: column; gap: 2px; }
  .extra-meta strong { color: var(--text); font-size: 13px; }

  /* === RIGHT PANEL === */
  .right { background: var(--bg-elev); border-left: 1px solid var(--border); display: flex; flex-direction: column; overflow: hidden; }
  .right-header { padding: 16px; border-bottom: 1px solid var(--border); display: flex; flex-direction: column; gap: 8px; }
  .right-header h2 { margin: 0; font-size: 13px; color: var(--accent); text-transform: uppercase; letter-spacing: 0.6px; }
  .setlist-name, .setlist-select { padding: 6px 8px; background: var(--bg); border: 1px solid var(--border); color: var(--text); border-radius: 4px; font-size: 12.5px; outline: none; width: 100%; }
  .setlist-items { flex: 1; overflow-y: auto; padding: 8px 0; }
  .setlist-item { padding: 8px 12px; border-left: 3px solid transparent; cursor: pointer; display: flex; align-items: center; gap: 8px; font-size: 13px; }
  .setlist-item:hover { background: var(--bg-elev2); }
  .setlist-item.current { border-left-color: var(--accent); background: var(--bg-elev2); }
  .setlist-item .num { color: var(--text-dim); font-family: monospace; min-width: 18px; }
  .setlist-item .title { flex: 1; }
  .setlist-item .remove { color: var(--text-dim); cursor: pointer; padding: 0 4px; opacity: 0; transition: opacity 0.15s; }
  .setlist-item:hover .remove { opacity: 1; }
  .setlist-item .remove:hover { color: var(--accent); }
  .setlist-item.dragging { opacity: 0.4; }
  .setlist-item .bpm-badge { font-size: 10px; color: var(--text-dim); background: var(--bg); border: 1px solid var(--border); border-radius: 3px; padding: 1px 5px; font-family: monospace; white-space: nowrap; }
  .setlist-empty { padding: 16px; color: var(--text-dim); font-size: 12px; font-style: italic; text-align: center; }
  .setlist-actions { padding: 12px 16px; border-top: 1px solid var(--border); display: flex; flex-direction: column; gap: 6px; }
  .setlist-sync-status { font-size: 11px; color: var(--text-dim); text-align: center; padding: 2px 0 4px; min-height: 14px; }
  .setlist-sync-status.ok { color: var(--ok); }
  .setlist-sync-status.warn { color: var(--warn); }
  .setlist-sync-status.err { color: var(--err); }

  /* === SETLIST PLAY MODE (lecture continue de toute la setlist) === */
  .app.setlist-mode .left, .app.setlist-mode .right { display: none; }
  .app.setlist-mode { grid-template-columns: 1fr; }
  .app.setlist-mode .center-header { background: var(--bg); padding: 12px 24px; border-bottom: 2px solid var(--accent); }
  .app.setlist-mode .stage-hide { display: none; }
  .app.setlist-mode .song-title { font-size: 17px; }
  .setlist-flow { padding: 12px 0 80px; }
  .setlist-flow .song-block { margin-bottom: 36px; }
  .setlist-flow .song-block:last-child { margin-bottom: 12px; }
  .setlist-flow .song-block-header {
    position: sticky; top: 0; z-index: 10;
    background: linear-gradient(to bottom, var(--bg) 0%, var(--bg) 70%, transparent 100%);
    padding: 14px 0 10px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 14px;
    display: flex; align-items: baseline; gap: 14px;
  }
  .setlist-flow .song-num { color: var(--accent); font-weight: 700; font-size: 22px; min-width: 36px; }
  .setlist-flow .song-block-title { color: var(--text); font-size: 22px; font-weight: 600; flex: 1; }
  .setlist-flow .song-meta { color: var(--text-dim); font-size: 12px; display: flex; gap: 12px; flex-wrap: wrap; }
  .setlist-flow .song-meta strong { color: var(--text); }
  .setlist-flow pre.song-body { padding: 6px 0 0; }
  .setlist-flow .song-sep {
    height: 2px; background: linear-gradient(to right, transparent, var(--border), transparent);
    margin: 24px 0;
  }
  @media (max-width: 700px) {
    .setlist-flow { padding: 0 0 100px; }
    .setlist-flow .song-num { font-size: 18px; min-width: 28px; }
    .setlist-flow .song-block-title { font-size: 18px; }
  }

  /* === NOTES PANEL === */
  .notes-panel {
    margin-top: 36px;
    border-top: 1px dashed var(--border);
    padding-top: 20px;
  }
  .notes-panel .notes-header {
    display: flex; align-items: center; gap: 8px; margin-bottom: 10px;
    font-size: 11px; color: var(--text-dim); font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.6px;
  }
  .notes-panel .notes-edit-btn {
    cursor: pointer; padding: 2px 9px; border-radius: 4px; font-size: 11px;
    background: none; border: 1px solid var(--border); color: var(--text-dim);
    font-family: inherit;
  }
  .notes-panel .notes-edit-btn:hover { background: var(--bg-elev2); color: var(--accent); border-color: var(--accent); }
  .notes-panel .notes-content {
    background: rgba(244, 162, 97, 0.07);
    border: 1px solid rgba(244, 162, 97, 0.25);
    border-left: 3px solid var(--warn);
    border-radius: 0 6px 6px 0;
    padding: 12px 16px;
    font-size: var(--song-fs, 14px);
    color: var(--text);
    white-space: pre-wrap;
    line-height: 1.6;
    min-height: 48px;
  }
  .notes-panel .notes-empty {
    color: var(--text-dim); font-style: italic; font-size: 12px; padding: 6px 0;
  }
  .notes-panel textarea.notes-textarea {
    width: 100%; min-height: 120px;
    padding: 10px 12px;
    background: rgba(244, 162, 97, 0.05);
    border: 1px solid var(--warn);
    border-left: 3px solid var(--warn);
    color: var(--text);
    border-radius: 0 6px 6px 0;
    font-size: var(--song-fs, 14px);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    line-height: 1.6;
    outline: none; resize: vertical;
  }
  .notes-panel textarea.notes-textarea:focus { border-color: var(--accent); }
  .notes-panel .notes-actions { display: flex; gap: 8px; margin-top: 8px; }

  /* === STAGE MODE === */
  .app.stage .left, .app.stage .right { display: none; }
  .app.stage .center-header { background: var(--bg); padding: 10px 24px; }
  .app.stage .song-title { font-size: 18px; }
  .app.stage .stage-hide { display: none; }
  .app.stage .notes-panel .notes-content { font-size: calc(var(--song-fs, 14px) * 1.05); border-left-width: 4px; }

  /* === MODAL === */
  .modal-bg { position: fixed; inset: 0; background: rgba(0,0,0,0.6); display: none; align-items: center; justify-content: center; z-index: 100; }
  .modal-bg.show { display: flex; }
  .modal { background: var(--bg-elev); border: 1px solid var(--border); border-radius: 10px; padding: 24px; max-width: 540px; width: 90%; max-height: 90vh; overflow-y: auto; }
  .modal h2 { margin: 0 0 8px; font-size: 18px; color: var(--accent); }
  .modal p, .modal label { font-size: 13px; line-height: 1.5; }
  .modal label { display: block; margin: 16px 0 6px; color: var(--text-dim); font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
  .modal input[type=text], .modal input[type=url], .modal textarea { width: 100%; padding: 9px 11px; background: var(--bg); border: 1px solid var(--border); color: var(--text); border-radius: 5px; font-size: 13px; outline: none; font-family: inherit; }
  .modal input:focus { border-color: var(--accent); }
  .modal .btn-row { display: flex; gap: 8px; margin-top: 16px; flex-wrap: wrap; }
  .modal .help { background: var(--bg); border-left: 3px solid var(--chord); padding: 10px 14px; border-radius: 0 5px 5px 0; font-size: 12.5px; color: var(--text-dim); line-height: 1.6; margin-top: 12px; }
  .modal .help ol { margin: 6px 0; padding-left: 20px; }
  .modal .status { padding: 8px 12px; border-radius: 5px; font-size: 12.5px; margin-top: 12px; }
  .modal .status.ok { background: rgba(107, 194, 107, 0.15); color: var(--ok); border: 1px solid rgba(107, 194, 107, 0.3); }
  .modal .status.err { background: rgba(224, 122, 95, 0.15); color: var(--err); border: 1px solid rgba(224, 122, 95, 0.3); }
  .modal .close { float: right; cursor: pointer; color: var(--text-dim); font-size: 22px; line-height: 1; }
  .modal .close:hover { color: var(--accent); }

  /* === ÉCOUTER === */
  .btn.listen-spotify { background: #1db954; border-color: #1db954; color: #06210f; font-weight: 600; }
  .btn.listen-spotify:hover { background: #3ad06e; border-color: #3ad06e; }
  .song-item .artist { color: var(--text-dim); font-size: 11px; margin-right: 6px; }
  .meta-pill.artist { background: rgba(244, 162, 97, 0.08); }

  /* === FICHE INFOS === */
  .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px; }
  .info-grid > div { min-width: 0; }
  .info-grid input, .info-grid select, #info-tags {
    width: 100%; padding: 9px 11px; background: var(--bg); border: 1px solid var(--border);
    color: var(--text); border-radius: 5px; font-size: 16px; outline: none; font-family: inherit; min-height: 42px;
  }
  .info-grid input:focus, .info-grid select:focus, #info-tags:focus { border-color: var(--accent); }
  .info-hint { color: var(--text-dim); font-size: 11.5px; margin-top: 4px; }
  .tag-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
  .tag-chip { padding: 5px 11px; border-radius: 14px; border: 1px solid var(--border); background: var(--bg); color: var(--text-dim); font-size: 12.5px; cursor: pointer; font-family: inherit; }
  .tag-chip.on { background: var(--accent); border-color: var(--accent); color: #1a1208; font-weight: 600; }
  @media (max-width: 480px) { .info-grid { grid-template-columns: 1fr; } }
  #toast { position: fixed; left: 50%; bottom: calc(24px + env(safe-area-inset-bottom)); transform: translateX(-50%) translateY(20px); opacity: 0; pointer-events: none;
    background: var(--bg-elev2); border: 1px solid var(--border); color: var(--text); padding: 10px 16px; border-radius: 8px; font-size: 13px; z-index: 9998; transition: opacity .2s, transform .2s; max-width: 90vw; text-align: center; }
  #toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
  #toast.ok { border-color: var(--ok); } #toast.err { border-color: var(--err); color: var(--err); }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 8px; height: 8px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--text-dim); }

  /* === SIDEBAR TOGGLE === */
  .sidebar-toggle { font-size: 20px !important; padding: 4px 10px !important; margin-right: 6px; }
  /* Desktop : repli en réduisant la colonne gauche */
  .app.sidebar-collapsed { grid-template-columns: 0 1fr 300px; }
  .app.sidebar-collapsed .left { display: none; }
  /* === RESPONSIVE : tablette + téléphone ===
     <= 1100px : liste (gauche) et setlist (droite) deviennent des panneaux coulissants,
     la zone de lecture prend toute la largeur. */
  .setlist-toggle { display: none; }
  @media (max-width: 1100px) {
    .app, .app.sidebar-collapsed { grid-template-columns: 1fr; }
    .app .left, .app .right {
      position: fixed; top: 0; bottom: 0;
      width: 86vw; max-width: 360px; z-index: 200;
      transition: transform 0.25s ease;
      display: flex !important; flex-direction: column;
      padding-top: env(safe-area-inset-top);
    }
    .app .left { left: 0; transform: translateX(-100%); box-shadow: 6px 0 24px rgba(0,0,0,0.6); }
    .app .right { right: 0; transform: translateX(100%); box-shadow: -6px 0 24px rgba(0,0,0,0.6); background: var(--bg-elev); overflow-y: auto; }
    .app.sidebar-open .left { transform: translateX(0); }
    .app.setlist-open .right { transform: translateX(0); }
    .app.sidebar-open::after, .app.setlist-open::after {
      content: ""; position: fixed; inset: 0;
      background: rgba(0,0,0,0.55); z-index: 150;
    }
    .setlist-toggle { display: inline-flex; }
    .app.stage .setlist-toggle { display: none; }
    .app.stage .left, .app.stage .right, .app.setlist-mode .left, .app.setlist-mode .right { display: none !important; }
  }
  @media (max-width: 700px) {
    .center-body { padding: 14px 14px calc(100px + env(safe-area-inset-bottom)); }
    .center-header { padding: 8px 10px; gap: 6px; flex-wrap: wrap; }
    .song-title { font-size: 17px !important; flex: 1 1 100%; }
    .pencil-body { white-space: nowrap; }
    .controls { flex-wrap: wrap; }
  }
  /* Zones sûres (encoche, barre d'état, barre du bas) : le bouton Scène/quitter reste toujours cliquable */
  .center-header {
    padding-top: calc(8px + env(safe-area-inset-top));
    padding-left: max(10px, env(safe-area-inset-left));
    padding-right: max(10px, env(safe-area-inset-right));
  }
  .app.stage .center-header {
    padding-top: calc(10px + env(safe-area-inset-top));
    padding-left: max(14px, env(safe-area-inset-left));
    padding-right: max(14px, env(safe-area-inset-right));
  }
  .app.stage .song-title { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

  /* Tactile : zones de toucher confortables (>= 40px) */
  @media (pointer: coarse) {
    .btn { min-height: 40px; min-width: 40px; padding: 8px 12px; }
    .pencil { padding: 8px 10px; font-size: 16px; }
    .song-item { padding-top: 12px; padding-bottom: 12px; }
    .setlist-item { padding-top: 12px; padding-bottom: 12px; }
    .left-header h1 .gear, .left-header h1 .add-btn { padding: 8px 12px; font-size: 18px; }
    .filters select, .filters input { min-height: 40px; font-size: 16px; }
  }

  /* Paroles : la taille est ajustée automatiquement à la largeur (voir fitSongBodies) */
  pre.song-body { font-size: var(--song-fs-fit, var(--song-fs, 14px)); }

  /* === AUTO-SCROLL CONTROLS (téléprompteur) === */
  .scroll-controls {
    position: absolute; bottom: 18px; right: 24px;
    display: flex; gap: 6px; align-items: center;
    background: rgba(35, 44, 54, 0.92); backdrop-filter: blur(8px);
    border: 1px solid var(--border); border-radius: 999px;
    padding: 6px 10px; z-index: 50;
    box-shadow: 0 6px 20px rgba(0,0,0,0.45);
  }
  .scroll-controls .sc-btn {
    background: transparent; border: 0; color: var(--text);
    width: 34px; height: 34px; border-radius: 50%;
    font-size: 16px; cursor: pointer;
    display: inline-flex; align-items: center; justify-content: center;
    transition: background 0.15s;
  }
  .scroll-controls .sc-btn:hover { background: var(--bg-elev2); color: var(--accent); }
  .scroll-controls.playing #scroll-toggle { background: var(--accent); color: var(--bg); }
  .scroll-controls .sc-speed {
    min-width: 56px; text-align: center;
    color: var(--text-dim); font-size: 12px;
    font-variant-numeric: tabular-nums; padding: 0 4px;
    letter-spacing: 0.3px;
  }
  .scroll-controls .sc-speed::after { content: " px/s"; opacity: 0.65; }
  .app.stage .scroll-controls { bottom: 28px; right: 32px; }
  @media (max-width: 700px) {
    .scroll-controls { bottom: 14px; right: 14px; padding: 4px 8px; }
    .scroll-controls .sc-btn { width: 38px; height: 38px; font-size: 17px; }
    .scroll-controls .sc-speed { min-width: 48px; }
  }

</style>
</head>
<body>

<div class="app" id="app">

  <!-- LEFT: SONG LIST -->
  <aside class="left">
    <div class="left-header">
      <h1>🎸 Répertoire Groove In
        <span class="icons">
          <span class="add-btn" onclick="openAddSong()" title="Ajouter une chanson (coller le texte)">＋</span>
          <span class="gear" onclick="openSettings()" title="Source des données">⚙</span>
        </span>
      </h1>
      <div class="subtitle"><span id="count-songs"></span> chansons</div>
      <div class="source-info" id="source-info">Source : interne</div>
      <button id="connect-btn" class="connect-btn" onclick="forceReconnect()" title="Forcer la reconnexion au Sheet" style="display:none">● Connecter</button>
    </div>
    <div class="filters">
      <input id="search" type="search" placeholder="Rechercher (titre, accord, parole…)" autocomplete="off">
      <div class="row">
        <select id="filter-key"><option value="">Toute tonalité</option></select>
        <select id="filter-capo"><option value="">Tout capo</option></select>
      </div>
      <div class="row">
        <select id="filter-style"><option value="">Tout style</option></select>
        <select id="filter-tag"><option value="">Tout tag</option></select>
      </div>
    </div>
    <div class="song-list" id="song-list"></div>
  </aside>

  <!-- CENTER: SONG VIEW -->
  <main class="center">
    <div class="center-header">
      <button class="btn icon ghost sidebar-toggle stage-hide" id="sidebar-toggle" onclick="toggleSidebar()" title="Afficher/masquer la liste des chansons" aria-label="Menu">☰</button>
      <h2 class="song-title"><span id="song-title">—</span><span class="pencil stage-hide" id="title-edit" onclick="editCurrentTitle()" title="Modifier le titre" style="display:none">✎</span><span class="pencil pencil-body stage-hide" id="body-edit" onclick="editCurrentBody()" title="Modifier paroles et accords" style="display:none">✎ Texte</span><span class="pencil pencil-body stage-hide" id="info-edit" onclick="openInfoModal()" title="Modifier genre, capo, tonalité, BPM, tags…" style="display:none">ⓘ Infos</span></h2>
      <span class="meta-pill artist stage-hide" id="meta-artiste"></span>
      <span class="meta-pill stage-hide" id="meta-key"></span>
      <span class="meta-pill stage-hide" id="meta-capo"></span>
      <span class="meta-pill style stage-hide" id="meta-style"></span>
      <span class="meta-pill bpm stage-hide" id="meta-bpm"></span>
      <span class="meta-pill stage-hide" id="meta-chords"></span>
      <div class="controls">
        <button class="btn icon ghost stage-hide" title="Diminuer" onclick="changeFontSize(-1)">A−</button>
        <button class="btn icon ghost stage-hide" title="Augmenter" onclick="changeFontSize(1)">A+</button>
        <button class="btn icon ghost stage-hide" id="fit-btn" title="Ajuster le texte à la largeur de l'écran (activé/désactivé)" onclick="toggleFit()">⇔</button>
        <button class="btn icon ghost" title="Transposer −1" onclick="transpose(-1)">♭</button>
        <button class="btn icon ghost" title="Transposer +1" onclick="transpose(1)">♯</button>
        <button class="btn icon ghost" title="Réinitialiser transposition" onclick="resetTranspose()">⟲</button>
        <button class="btn listen-spotify stage-hide" id="listen-spotify" onclick="openListen('spotify')" style="display:none" title="Ouvrir dans Spotify">▶ Spotify</button>
        <button class="btn ghost stage-hide" id="listen-other" onclick="openListen('other')" style="display:none" title="Ouvrir le lien d'écoute">▶ Écouter</button>
        <button class="btn ghost stage-hide" id="listen-search" onclick="openListen('search')" style="display:none" title="Chercher ce titre dans Spotify">🔎 Écouter</button>
        <button class="btn ghost stage-hide" onclick="addCurrentToSetlist()">+ Setlist</button>
        <button class="btn ghost stage-hide setlist-toggle" onclick="toggleSetlistPanel()" title="Afficher la setlist" aria-label="Setlist">📋 Set</button>
        <button class="btn ghost stage-hide" id="singers-btn" onclick="openSingersModal()" title="Configurer les chanteurs">🎤<span id="singers-btn-count" style="margin-left:3px;font-size:11px;font-weight:700"></span></button>
        <button class="btn icon ghost" id="fullscreen-btn" onclick="toggleFullscreen()" title="Plein écran (F)">⛶</button>
        <button class="btn primary" onclick="toggleStage()">🎤 Scène</button>
      </div>
    </div>
    <div class="singer-banner" id="singer-banner">
      <span><strong>Mode attribution</strong> &middot; Cliquez sur une ligne pour cycler le chanteur</span>
      <span class="palette" id="singer-palette"></span>
      <button class="btn ghost close-assign" onclick="exitSingerAssignMode()">&#x2715; Quitter</button>
    </div>
    <div class="center-body" id="song-body">
      <div class="empty">Sélectionnez une chanson dans la liste à gauche.</div>
    </div>
    <div class="scroll-controls" id="scroll-controls" style="display:none">
      <button class="sc-btn" id="scroll-toggle" onclick="toggleAutoScroll()" title="Lecture / pause (Espace)">▶</button>
      <button class="sc-btn" onclick="changeScrollSpeed(-5)" title="Plus lent (−)">−</button>
      <span class="sc-speed" id="scroll-speed-label">25</span>
      <button class="sc-btn" onclick="changeScrollSpeed(5)" title="Plus rapide (+)">+</button>
      <button class="sc-btn sc-reset" onclick="resetScroll()" title="Retour en haut">⤴</button>
    </div>
  </main>

  <!-- RIGHT: SETLIST -->
  <aside class="right">
    <div class="right-header">
      <h2>Setlist du concert</h2>
      <select class="setlist-select" id="setlist-select" onchange="loadSetlist(this.value)"></select>
      <input class="setlist-name" id="setlist-name" type="text" placeholder="Nom du concert" onchange="renameSetlist(this.value)">
    </div>
    <div class="setlist-items" id="setlist-items"></div>
    <div class="setlist-actions">
      <button class="btn primary" onclick="playSetlist()" id="play-setlist-btn">▶ Lancer le set</button>
      <button class="btn" id="save-setlist-btn" onclick="saveSetlistNow()">💾 Sauvegarder dans le Sheet</button>
      <div class="setlist-sync-status" id="setlist-sync-status">—</div>
      <button class="btn" onclick="newSetlist()">+ Nouvelle setlist</button>
      <button class="btn ghost" onclick="exportSetlist()">📋 Copier</button>
      <button class="btn ghost" onclick="deleteSetlist()" style="color:#ff6b6b;">🗑 Supprimer</button>
    </div>
    <div class="save-bar" id="save-bar">
      <div class="label">⚠ Modifications non sauvegardées</div>
      <button class="btn primary" onclick="saveAsXlsx()">💾 Télécharger le .xlsx mis à jour</button>
      <button class="btn ghost" onclick="discardChanges()" style="font-size:11px;">Annuler les modifications</button>
    </div>
  </aside>

</div>

<!-- ADD SONG MODAL -->
<div class="modal-bg" id="add-modal">
  <div class="modal" style="max-width:680px;">
    <span class="close" onclick="closeAddSong()">×</span>
    <h2 id="add-modal-title">Ajouter une chanson</h2>
    <p style="color:var(--text-dim);margin:0 0 8px;">Collez le texte de la chanson (titre en première ligne, puis paroles avec accords). La détection est automatique.</p>
    <label for="paste-area">Texte de la chanson</label>
    <textarea id="paste-area" class="song-input" placeholder="Mon Titre Capo 2

[Intro]
Am  C  G  D

[Verse 1]
Am          C
Première ligne de paroles
G                 D
Deuxième ligne…

[Chorus]
F   C   G   Am
…"></textarea>
    <div id="parse-preview"></div>
    <div class="btn-row">
      <button class="btn primary" id="add-save-btn" onclick="saveNewSong()">✓ Ajouter au répertoire</button>
      <button class="btn ghost" onclick="closeAddSong()">Annuler</button>
    </div>
  </div>
</div>

<!-- SINGERS CONFIG MODAL -->
<div id="toast"></div>
<!-- FICHE INFOS -->
<div class="modal-bg" id="info-modal">
  <div class="modal" style="max-width:600px;">
    <span class="close" onclick="closeInfoModal()">&#xd7;</span>
    <h2>ⓘ Infos — <span id="info-title"></span></h2>
    <div class="info-grid">
      <div style="grid-column:1/-1"><label for="info-artiste">Artiste</label><input id="info-artiste" type="text" autocomplete="off" placeholder="ex. Calogero"></div>
      <div><label for="info-genre">Genre</label><input id="info-genre" type="text" list="dl-genre" autocomplete="off"></div>
      <div><label for="info-style">Style</label><input id="info-style" type="text" list="dl-style" autocomplete="off"></div>
      <div><label for="info-rythme">Rythme</label><input id="info-rythme" type="text" list="dl-rythme" autocomplete="off"></div>
      <div><label for="info-bpm">BPM</label><input id="info-bpm" type="number" inputmode="numeric" min="30" max="300" placeholder="ex. 90"></div>
      <div><label for="info-capo">Capo</label><select id="info-capo"></select></div>
      <div><label for="info-key">Tonalité jouée</label><input id="info-key" type="text" list="dl-key" autocomplete="off"><div class="info-hint" id="info-key-hint"></div></div>
      <div><label for="info-duree">Durée</label><input id="info-duree" type="text" inputmode="numeric" placeholder="ex. 3:45" autocomplete="off"></div>
      <div><label for="info-niveau">Niveau</label><select id="info-niveau"></select></div>
    </div>
    <label for="info-tags">Tags concert</label>
    <input id="info-tags" type="text" placeholder="ex. mariage, cocktail" autocomplete="off" oninput="syncTagChips()">
    <div class="tag-chips" id="info-tag-chips"></div>
    <div class="info-grid" style="grid-template-columns:1fr">
      <div><label for="info-spotify">Lien Spotify</label><input id="info-spotify" type="url" inputmode="url" autocomplete="off" placeholder="https://open.spotify.com/track/…"><div class="info-hint">Dans Spotify : ⋯ → Partager → Copier le lien du titre.</div></div>
      <div><label for="info-lien">Autre lien d'écoute</label><input id="info-lien" type="url" inputmode="url" autocomplete="off" placeholder="YouTube, Deezer, Apple Music…"></div>
    </div>
    <datalist id="dl-genre"></datalist><datalist id="dl-style"></datalist><datalist id="dl-rythme"></datalist><datalist id="dl-key"></datalist>
    <div class="btn-row" style="margin-top:20px">
      <button class="btn primary" onclick="saveInfo()">✓ Enregistrer</button>
      <button class="btn ghost" onclick="closeInfoModal()">Annuler</button>
    </div>
  </div>
</div>

<div class="modal-bg" id="singers-modal">
  <div class="modal">
    <span class="close" onclick="closeSingersModal()">&#xd7;</span>
    <h2>&#x1f3a4; Chanteurs de cette chanson</h2>
    <p style="color:var(--text-dim);font-size:12px;margin:0 0 8px">
      Choisissez le nombre de chanteurs (1 a 4). Optionnel : nommez-les. Puis activez le mode <strong>Attribution</strong> pour cliquer sur chaque ligne et lui assigner un chanteur (ou laisser "tous" pour les passages communs).
    </p>
    <div class="singer-config">
      <div>
        <label style="margin-top:0">Nombre de chanteurs</label>
        <div class="count-row" id="singer-count-row">
          <button class="btn" data-n="1" onclick="setSingerCount(1)">1</button>
          <button class="btn" data-n="2" onclick="setSingerCount(2)">2</button>
          <button class="btn" data-n="3" onclick="setSingerCount(3)">3</button>
          <button class="btn" data-n="4" onclick="setSingerCount(4)">4</button>
        </div>
      </div>
      <div id="singer-names-list"></div>
    </div>
    <div class="btn-row" style="margin-top:18px">
      <button class="btn primary" onclick="enterSingerAssignMode()">&#x1f3af; Activer le mode attribution</button>
      <button class="btn ghost" onclick="resetSongAssignments()">&#x21ba; Tout remettre a "tous"</button>
    </div>
  </div>
</div>

<!-- SETTINGS MODAL -->
<div class="modal-bg" id="settings-modal">
  <div class="modal">
    <span class="close" onclick="closeSettings()">×</span>
    <h2>Source des données</h2>
    <p style="color:var(--text-dim);margin:0 0 4px;">Connectez l'interface à votre Google Sheet via Apps Script (lecture + écriture en direct), ou en mode .xlsx (lecture seule).</p>

    <h3 style="margin:14px 0 4px;font-size:14px;color:var(--accent)">⚡ Recommandé : Apps Script (sync bidirectionnel)</h3>
    <label for="gas-url">URL Apps Script (terminée par <code>/exec</code>)</label>
    <input id="gas-url" type="url" placeholder="https://script.google.com/macros/s/AKfycb.../exec">
    <div class="btn-row">
      <button class="btn primary" onclick="connectGAS()">🔌 Connecter au Sheet</button>
      <button class="btn ghost" onclick="reloadFromGAS()">↻ Recharger depuis Sheet</button>
      <button class="btn ghost" onclick="syncSetlistsFromGAS()">📋 Récupérer les setlists</button>
      <button class="btn ghost" onclick="testGAS()">🩺 Test connexion</button>
    </div>

    <details style="margin:14px 0 8px">
      <summary style="cursor:pointer;color:var(--text-dim);font-size:13px">Mode alternatif : import .xlsx (lecture seule)</summary>
      <label for="drive-url" style="margin-top:10px">Lien Drive du fichier .xlsx</label>
      <input id="drive-url" type="url" placeholder="https://drive.google.com/file/d/XXXXXXX/view?usp=sharing">
      <div class="btn-row">
        <button class="btn ghost" onclick="loadFromDriveUrl()">↓ Charger depuis Drive</button>
        <button class="btn ghost" onclick="document.getElementById('file-picker').click()">📂 Importer fichier .xlsx</button>
        <button class="btn ghost" onclick="resetToEmbedded()">↺ Base interne</button>
        <input type="file" id="file-picker" accept=".xlsx,.xls" style="display:none" onchange="loadFromFile(this.files[0])">
      </div>
    </details>

    <div class="status" id="load-status" style="display:none"></div>

    <div class="help">
      <strong>Apps Script — mise en place (~5 min, une seule fois) :</strong>
      <ol>
        <li>Ouvrez votre Sheet → <em>Extensions → Apps Script</em></li>
        <li>Collez le code de <code>AppsScript_Code.gs</code> dans <code>Code.gs</code> et sauvegardez</li>
        <li>Exécutez la fonction <code>smokeTest</code> et autorisez l'accès à votre Sheet</li>
        <li><em>Déployer → Nouveau déploiement → Application Web</em>. Choisissez "Tout le monde" pour l'accès</li>
        <li>Copiez l'URL <code>/exec</code> dans le champ ci-dessus → "Connecter au Sheet"</li>
      </ol>
      Détails complets dans le fichier <code>Instructions_AppsScript.md</code>.
    </div>
  </div>
</div>

<script>
const EMBEDDED_SONGS = __SONGS_JSON__;
const BUILD_VERSION = "__BUILD_VERSION__";

// ===== STATE =====
const STORAGE_KEY = "groove-in-setlists-v1";
const SETTINGS_KEY = "groove-in-settings-v2";
let state = {
  currentSongId: null,
  fontSize: 14,
  fitWidth: true,
  transposeSteps: 0,
  songs: [...EMBEDDED_SONGS],
  filtered: [...EMBEDDED_SONGS],
  setlists: {},
  currentSetlistId: null,
  driveUrl: "",
  gasUrl: "",
  source: "interne",  // "interne" | "drive" | "fichier" | "gas"
  lastLoaded: null,
  dirty: false,
  editingId: null,
  gasSyncing: false,
  songSingers: {},     // { [songId]: { count: 1-4, names: [...], assignments: { [lineKey]: 0-4 } } }
  singerAssignMode: false,
};

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const data = JSON.parse(raw);
      state.setlists = data.setlists || {};
      state.currentSetlistId = data.currentSetlistId || null;
    }
    const settings = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
    if (settings.fontSize) state.fontSize = settings.fontSize;
    if (typeof settings.fitWidth === "boolean") state.fitWidth = settings.fitWidth;
    if (settings.driveUrl) state.driveUrl = settings.driveUrl;
    if (settings.gasUrl) state.gasUrl = settings.gasUrl;
    if (settings.songSingers) state.songSingers = settings.songSingers;
  } catch (e) { console.warn(e); }
  if (Object.keys(state.setlists).length === 0) {
    const id = "sl_" + Date.now();
    state.setlists[id] = { id, name: "Concert sans titre", songIds: [], notes: "" };
    state.currentSetlistId = id;
  }
  if (!state.setlists[state.currentSetlistId]) state.currentSetlistId = Object.keys(state.setlists)[0];
}

function saveToStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ setlists: state.setlists, currentSetlistId: state.currentSetlistId }));
  localStorage.setItem(SETTINGS_KEY, JSON.stringify({ fontSize: state.fontSize, fitWidth: state.fitWidth, driveUrl: state.driveUrl, gasUrl: state.gasUrl, songSingers: state.songSingers }));
  // Si connecté à Apps Script, synchroniser les setlists et chanteurs vers le Sheet (avec debounce)
  if (typeof scheduleSetlistsSave === "function") scheduleSetlistsSave();
  if (typeof scheduleSingersSave === "function") scheduleSingersSave();
}

// ===== DRIVE & XLSX LOADING =====
function extractDriveFileId(url) {
  if (!url) return null;
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

async function fetchXlsxFromDrive(input) {
  const id = extractDriveFileId(input);
  const candidates = [];
  if (id) {
    // Best for Google Sheets URLs (and works for .xlsx-in-Sheets too)
    candidates.push(`https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`);
    // Fallback for files stored as raw .xlsx on Drive (not opened in Sheets)
    candidates.push(`https://drive.google.com/uc?export=download&id=${id}`);
  } else {
    candidates.push(input); // assume direct URL
  }
  let lastErr;
  for (const url of candidates) {
    try {
      const resp = await fetch(url, { redirect: "follow" });
      if (!resp.ok) { lastErr = new Error(`HTTP ${resp.status}`); continue; }
      const ab = await resp.arrayBuffer();
      // Validate it's a real xlsx (zip files start with "PK")
      const b = new Uint8Array(ab);
      if (b.length >= 2 && b[0] === 0x50 && b[1] === 0x4B) return ab;
      lastErr = new Error("réponse non-xlsx (Google a renvoyé une page HTML — vérifiez que le fichier est bien partagé en public)");
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr || new Error("téléchargement impossible");
}

async function loadFromDriveUrl() {
  const url = document.getElementById("drive-url").value.trim();
  if (!url) { showStatus("err", "Collez d'abord un lien."); return; }
  state.driveUrl = url;
  saveToStorage();
  showStatus("ok", "Téléchargement en cours…");
  try {
    const ab = await fetchXlsxFromDrive(url);
    const songs = parseXlsxArrayBuffer(ab);
    if (!songs.length) throw new Error("Aucune chanson trouvée dans la feuille « Répertoire ».");
    state.songs = songs;
    state.source = "drive";
    state.lastLoaded = new Date();
    showStatus("ok", `✓ ${songs.length} chansons chargées depuis Drive.`);
    refreshAll();
  } catch (e) {
    console.error(e);
    const id = extractDriveFileId(url);
    const dlLink = id ? `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx` : url;
    const html = `<strong>Échec : ${escapeHtml(e.message)}</strong>
      <div style="margin-top:6px;font-size:13px;line-height:1.5">
        Le navigateur bloque l'accès direct à Drive (politique <em>CORS</em> de Google). Solution rapide :
        <ol style="margin:6px 0 0 18px;padding:0">
          <li><a href="${dlLink}" target="_blank" rel="noopener">📥 Cliquez ici pour télécharger le .xlsx</a></li>
          <li>Faites-le <strong>glisser-déposer</strong> sur cette fenêtre, ou utilisez « Importer fichier .xlsx ».</li>
        </ol>
      </div>`;
    showStatusHtml("err", html);
  }
}

function loadFromFile(file) {
  if (!file) return;
  showStatus("ok", "Lecture du fichier…");
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const songs = parseXlsxArrayBuffer(e.target.result);
      if (!songs.length) throw new Error("Aucune chanson trouvée dans la feuille « Répertoire ».");
      state.songs = songs;
      state.source = "fichier";
      state.lastLoaded = new Date();
      showStatus("ok", `✓ ${songs.length} chansons chargées depuis « ${file.name} ».`);
      refreshAll();
    } catch (err) {
      console.error(err);
      showStatus("err", `Échec : ${err.message}`);
    }
  };
  reader.onerror = () => showStatus("err", "Erreur lecture du fichier.");
  reader.readAsArrayBuffer(file);
}

function resetToEmbedded() {
  state.songs = [...EMBEDDED_SONGS];
  state.source = "interne";
  state.lastLoaded = null;
  showStatus("ok", `↺ Base interne restaurée (${EMBEDDED_SONGS.length} chansons).`);
  refreshAll();
}

// ===== BOUTON CONNEXION =====
function setConnectBtn(state_) {
  const btn = document.getElementById("connect-btn");
  if (!btn) return;
  if (!state.gasUrl) { btn.style.display = "none"; return; }
  btn.style.display = "";
  btn.classList.remove("ok", "err", "loading");
  if (state_ === "ok")      { btn.classList.add("ok");      btn.textContent = "● Sheet connecté"; btn.disabled = false; }
  else if (state_ === "err"){ btn.classList.add("err");     btn.textContent = "● Reconnexion échouée — cliquer pour réessayer"; btn.disabled = false; }
  else if (state_ === "loading") { btn.classList.add("loading"); btn.textContent = "● Connexion en cours…"; btn.disabled = true; }
  else                      {                               btn.textContent = "● Connecter au Sheet"; btn.disabled = false; }
}

async function forceReconnect() {
  if (!state.gasUrl) { openSettings(); return; }
  setConnectBtn("loading");
  try {
    const [dataList, dataSetlists, dataSingers] = await Promise.all([
      gasGet("list"),
      gasGet("setlists").catch(() => null),
      gasGet("singers").catch(() => null),
    ]);
    const songs = (dataList.songs || []).map(normalizeSongFromGAS);
    if (!songs.length) throw new Error("Aucune chanson dans le Sheet.");
    state.songs = songs;
    state.source = "gas";
    state.lastLoaded = new Date();
    if (dataSetlists && dataSetlists.setlists && Object.keys(dataSetlists.setlists).length) {
      state.setlists = dataSetlists.setlists;
      state.currentSetlistId = dataSetlists.currentId || Object.keys(dataSetlists.setlists)[0];
    }
    if (dataSingers && dataSingers.singers) state.songSingers = dataSingers.singers;
    setConnectBtn("ok");
    renderSourceInfo();
    refreshAll();
    showStatus("ok", `✓ Sheet reconnecté — ${songs.length} chansons chargées.`);
  } catch (e) {
    setConnectBtn("err");
    showStatus("err", `Reconnexion échouée : ${e.message}`);
  }
}

// ===== APPS SCRIPT (Google Sheet sync bidirectionnel) =====
const GAS_TIMEOUT_MS = 15000; // 15s — Apps Script a des cold-starts lents

function fetchWithTimeout(url, options = {}) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), GAS_TIMEOUT_MS);
  return fetch(url, { ...options, signal: ctrl.signal })
    .finally(() => clearTimeout(timer));
}

async function gasGet(action) {
  if (!state.gasUrl) throw new Error("URL Apps Script non configurée.");
  const url = state.gasUrl + (state.gasUrl.includes("?") ? "&" : "?") + "action=" + encodeURIComponent(action);
  const resp = await fetchWithTimeout(url, { method: "GET", redirect: "follow" });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  if (!data.ok) throw new Error(data.error || "réponse invalide");
  return data;
}

async function gasPost(payload) {
  if (!state.gasUrl) throw new Error("URL Apps Script non configurée.");
  // Content-Type "text/plain" pour éviter le préflight CORS
  const resp = await fetchWithTimeout(state.gasUrl, {
    method: "POST",
    redirect: "follow",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const data = await resp.json();
  if (!data.ok) throw new Error(data.error || "réponse invalide");
  return data;
}

function normalizeSongFromGAS(s) {
  // Apps Script renvoie chords déjà parsé en array. Coerce types.
  return {
    id: Number(s.id) || 0,
    title: s.title || "",
    capo: s.capo === "" ? null : (Number(s.capo) || s.capo || null),
    key_detected: s.key_detected || "",
    key_played: s.key_played || "",
    chords: Array.isArray(s.chords) ? s.chords : [],
    n_chords: Number(s.n_chords) || 0,
    n_lines: Number(s.n_lines) || 0,
    intros: s.intros || "",
    verses: s.verses || "",
    pre_chorus: s.pre_chorus || "",
    choruses: s.choruses || "",
    bridges: s.bridges || "",
    body_full: s.body_full || "",
    style: s.style || "",
    genre: s.genre || "",
    rythme: s.rythme || "",
    bpm: s.bpm || "",
    duree: s.duree || "",
    tags: s.tags || "",
    niveau: s.niveau || "",
    notes: s.notes || "",
    // Tolère l'ancienne version du script (clés = en-têtes bruts du Sheet)
    artiste: s.artiste || s["Artiste"] || "",
    spotify: s.spotify || s["Lien Spotify"] || "",
    lien_ecoute: s.lien_ecoute || s["Lien écoute"] || "",
  };
}

async function connectGAS() {
  const url = (document.getElementById("gas-url").value || "").trim();
  if (!url) { showStatus("err", "Collez d'abord l'URL Apps Script."); return; }
  if (!/\/macros\/s\/[^\/]+\/exec\b/.test(url)) {
    showStatus("err", "L'URL ne ressemble pas à un déploiement Apps Script (doit finir par /exec).");
    return;
  }
  state.gasUrl = url;
  saveToStorage();
  await reloadFromGAS();
}

async function reloadFromGAS() {
  if (!state.gasUrl) { showStatus("err", "Configurez d'abord l'URL Apps Script."); return; }
  showStatus("ok", "Lecture du Sheet via Apps Script…");
  try {
    const data = await gasGet("list");
    const songs = (data.songs || []).map(normalizeSongFromGAS);
    if (!songs.length) throw new Error("Aucune chanson trouvée dans la feuille « Répertoire ».");
    state.songs = songs;
    state.source = "gas";
    state.lastLoaded = new Date();
    state.dirty = false;
    // Aussi récupérer les setlists et chanteurs synchronisés
    await gasReadSetlists();
    await gasReadSingers();
    renderDirtyBar();
    showStatus("ok", `✓ ${songs.length} chansons chargées depuis le Sheet (Apps Script).`);
    refreshAll();
  } catch (e) {
    console.error(e);
    showStatusHtml("err", `<strong>Échec : ${escapeHtml(e.message)}</strong>
      <div style="margin-top:6px;font-size:13px;line-height:1.5">
        Vérifiez que :
        <ul style="margin:6px 0 0 18px;padding:0">
          <li>L'URL se termine bien par <code>/exec</code></li>
          <li>Le déploiement est en mode <em>"Tout le monde"</em> peut accéder</li>
          <li>Vous avez exécuté <code>smokeTest</code> au moins une fois et autorisé l'accès</li>
        </ul>
      </div>`);
  }
}

async function testGAS() {
  const url = (document.getElementById("gas-url").value || state.gasUrl || "").trim();
  if (!url) { showStatus("err", "Collez d'abord l'URL Apps Script."); return; }
  state.gasUrl = url;
  showStatus("ok", "Test en cours…");
  try {
    const r = await fetch(url + (url.includes("?") ? "&" : "?") + "action=ping", { method: "GET", redirect: "follow" });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const data = await r.json();
    if (data.ok) showStatus("ok", `✓ Connexion OK. Feuille active : ${data.sheet || "(?)"}`);
    else throw new Error(data.error || "réponse invalide");
  } catch (e) {
    showStatus("err", `Échec du test : ${e.message}`);
  }
}

async function gasPushSong(song) {
  // Met à jour une chanson existante (ou l'ajoute si pas d'id)
  if (state.source !== "gas" || !state.gasUrl) return;
  state.gasSyncing = true; renderDirtyBar();
  try {
    if (song.id) {
      await gasPost({ action: "update_song", id: song.id, fields: song });
    } else {
      const r = await gasPost({ action: "add_song", song: song });
      song.id = r.id;
    }
    state.dirty = false;
  } catch (e) {
    console.error("Push GAS échoué :", e);
    showStatus("err", `Sauvegarde Sheet échouée : ${e.message}`);
    state.dirty = true;
  } finally {
    state.gasSyncing = false;
    renderDirtyBar();
  }
}

async function gasDeleteSong(id) {
  if (state.source !== "gas" || !state.gasUrl) return;
  state.gasSyncing = true; renderDirtyBar();
  try {
    await gasPost({ action: "delete_song", id: id });
    state.dirty = false;
  } catch (e) {
    console.error(e);
    showStatus("err", `Suppression Sheet échouée : ${e.message}`);
    state.dirty = true;
  } finally {
    state.gasSyncing = false;
    renderDirtyBar();
  }
}

// ===== SETLISTS sync via Apps Script =====
let setlistSaveTimer = null;
function scheduleSetlistsSave() {
  if (state.source !== "gas" || !state.gasUrl) return;
  setSetlistSyncStatus("pending");
  if (setlistSaveTimer) clearTimeout(setlistSaveTimer);
  setlistSaveTimer = setTimeout(() => { gasSaveSetlists(); }, 800);
}

async function gasSaveSetlists() {
  if (state.source !== "gas" || !state.gasUrl) {
    setSetlistSyncStatus("nogas");
    return;
  }
  state.setlistsSyncing = true; renderSourceInfo();
  setSetlistSyncStatus("saving");
  try {
    const r = await gasPost({
      action: "save_setlists",
      setlists: state.setlists,
      currentSetlistId: state.currentSetlistId || "",
    });
    setSetlistSyncStatus("saved", r.n);
  } catch (e) {
    console.error("Sync setlists échoué :", e);
    setSetlistSyncStatus("error", e.message);
  } finally {
    state.setlistsSyncing = false; renderSourceInfo();
  }
}

function setSetlistSyncStatus(kind, info) {
  const el = document.getElementById("setlist-sync-status");
  if (!el) return;
  el.classList.remove("ok", "warn", "err");
  const now = new Date().toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" });
  if (kind === "saved") { el.classList.add("ok"); el.textContent = `✓ Sauvegardé à ${now}`; }
  else if (kind === "saving") { el.textContent = "Sauvegarde en cours…"; }
  else if (kind === "pending") { el.classList.add("warn"); el.textContent = "Modifications en attente…"; }
  else if (kind === "error") { el.classList.add("err"); el.textContent = `Échec : ${info || "réessayer"}`; }
  else if (kind === "nogas") { el.textContent = "Stocké localement (pas de Sheet connecté)"; }
  else { el.textContent = "—"; }
}

async function saveSetlistNow() {
  if (state.source !== "gas" || !state.gasUrl) {
    setSetlistSyncStatus("nogas");
    showStatus("err", "Connectez d'abord l'app à votre Sheet via ⚙ Paramètres pour synchroniser les setlists.");
    return;
  }
  if (setlistSaveTimer) { clearTimeout(setlistSaveTimer); setlistSaveTimer = null; }
  await gasSaveSetlists();
}

async function gasReadSetlists() {
  if (!state.gasUrl) return;
  try {
    const data = await gasGet("setlists");
    if (data && data.setlists && Object.keys(data.setlists).length > 0) {
      state.setlists = data.setlists;
      if (data.currentSetlistId && state.setlists[data.currentSetlistId]) {
        state.currentSetlistId = data.currentSetlistId;
      } else {
        state.currentSetlistId = Object.keys(state.setlists)[0];
      }
    }
  } catch (e) {
    console.warn("Lecture setlists échouée :", e.message);
  }
}

let singersSaveTimer = null;
function scheduleSingersSave() {
  if (state.source !== "gas" || !state.gasUrl) return;
  if (singersSaveTimer) clearTimeout(singersSaveTimer);
  singersSaveTimer = setTimeout(() => { gasSaveSingers(); }, 800);
}
async function gasSaveSingers() {
  if (state.source !== "gas" || !state.gasUrl) return;
  try {
    await gasPost({ action: "save_singers", singers: state.songSingers });
  } catch (e) { console.warn("Sync chanteurs echoue :", e.message); }
}
async function gasReadSingers() {
  if (!state.gasUrl) return;
  try {
    const data = await gasGet("singers");
    if (data && data.singers && typeof data.singers === "object") {
      state.songSingers = data.singers;
    }
  } catch (e) { /* silent */ }
}

async function syncSetlistsFromGAS() {
  if (!state.gasUrl) { showStatus("err", "Configurez d'abord l'URL Apps Script."); return; }
  showStatus("ok", "Récupération des setlists…");
  const before = JSON.stringify(state.setlists);
  await gasReadSetlists();
  const after = JSON.stringify(state.setlists);
  refreshAll();
  if (before === after) showStatus("ok", "Aucune setlist à récupérer (le Sheet est vide ou identique).");
  else showStatus("ok", `✓ ${Object.keys(state.setlists).length} setlist(s) récupérée(s) depuis le Sheet.`);
}

function parseXlsxArrayBuffer(ab) {
  const wb = XLSX.read(ab, { type: "array" });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("répertoire") || n.toLowerCase().includes("repertoire")) || wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return rows.filter(r => r["Titre"] || r["titre"] || r["Title"]).map((r, i) => {
    const get = (...keys) => { for (const k of keys) if (r[k] !== undefined && r[k] !== "") return r[k]; return ""; };
    const chordsRaw = String(get("Accords utilisés", "Accords", "Chords") || "");
    const chords = chordsRaw.split(/,\s*/).filter(Boolean);
    return {
      id: parseInt(get("ID", "id", "#")) || (i + 1),
      title: String(get("Titre", "titre", "Title")),
      capo: parseInt(get("Capo", "capo")) || null,
      key_detected: String(get("Tonalité (auto)", "Tonalité auto", "Tonalité")),
      key_played: String(get("Tonalité (jouée)", "Tonalité jouée")),
      style: String(get("Style", "style")),
      genre: String(get("Genre", "genre")),
      rythme: String(get("Rythme", "rythme")),
      bpm: String(get("BPM", "bpm")),
      duree: String(get("Durée", "duree", "Duration")),
      tags: String(get("Tags concert", "Tags", "tags")),
      niveau: String(get("Niveau", "niveau")),
      notes: String(get("Notes", "notes")),
      artiste: String(get("Artiste", "artiste")),
      spotify: String(get("Lien Spotify", "Spotify", "spotify")),
      lien_ecoute: String(get("Lien écoute", "Lien ecoute", "lien_ecoute")),
      chords,
      n_chords: parseInt(get("Nb accords")) || chords.length,
      n_lines: parseInt(get("Nb lignes")) || 0,
      intros: String(get("Intro", "intros")),
      verses: String(get("Couplets", "verses")),
      pre_chorus: String(get("Pré-refrain", "pre_chorus")),
      choruses: String(get("Refrain", "choruses")),
      bridges: String(get("Bridge", "bridges")),
      body_full: String(get("Body complet", "body_full", "Body")),
    };
  });
}

function showStatus(kind, msg) {
  const el = document.getElementById("load-status");
  el.className = "status " + kind;
  el.textContent = msg;
  el.style.display = "block";
}

function showStatusHtml(kind, html) {
  const el = document.getElementById("load-status");
  el.className = "status " + kind;
  el.innerHTML = html;
  el.style.display = "block";
}

// ===== TRANSPOSITION =====
const NOTES_SHARP = ["C","C#","D","D#","E","F","F#","G","G#","A","A#","B"];
const FLAT_TO_SHARP = {"Db":"C#","Eb":"D#","Gb":"F#","Ab":"G#","Bb":"A#"};
const CHORD_REGEX = /\b([A-G])(b|#)?(maj|min|m|M|aug|dim|sus|add)?(\d{0,2})((?:sus\d?|add\d|\/[A-G][#b]?)*)\b/g;

function transposeChord(chord, steps) {
  return chord.replace(CHORD_REGEX, (match, root, acc, qual, num, tail) => {
    let note = root + (acc || "");
    note = FLAT_TO_SHARP[note] || note;
    const idx = NOTES_SHARP.indexOf(note);
    if (idx < 0) return match;
    const newIdx = (idx + steps + 120) % 12;
    let newTail = (tail || "").replace(/\/([A-G])(b|#)?/g, (m, br, ba) => {
      let bn = br + (ba || "");
      bn = FLAT_TO_SHARP[bn] || bn;
      const bIdx = NOTES_SHARP.indexOf(bn);
      if (bIdx < 0) return m;
      return "/" + NOTES_SHARP[(bIdx + steps + 120) % 12];
    });
    return NOTES_SHARP[newIdx] + (qual || "") + (num || "") + newTail;
  });
}

// ===== RENDERING =====
function escapeHtml(str) {
  return String(str || "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]));
}

const CHORD_TOKEN_RE = /^[A-G][#b]?(maj|min|m|M|aug|dim|sus|add)?\d*((?:sus\d?|add\d|\/[A-G][#b]?\d?))*$/;
function isChordLine(trimmed) {
  if (!trimmed) return false;
  const tokens = trimmed.split(/\s+/).filter(Boolean).map(t => t.replace(/[|()\[\]:]/g, ""));
  if (!tokens.length) return false;
  const chordCount = tokens.filter(t => t && CHORD_TOKEN_RE.test(t)).length;
  return chordCount / tokens.length >= 0.6;
}

function lineKeyOf(line) { return line.replace(/\s+/g, " ").trim(); }

function renderSongBody(text, steps, songId) {
  if (!text) return "(aucun corps disponible)";
  const singersCfg = songId ? ((state.songSingers || {})[songId] || {}) : {};
  const assignments = singersCfg.assignments || {};
  return text.split("\n").map(line => {
    const trimmed = line.trim();
    if (/^\[?\s*(intro|verse|chorus|bridge|refrain|pont|pre[- ]?chorus|interlude|outro|hook|solo|couplet)\s*\d*\s*:?\s*\]?$/i.test(trimmed)) {
      return `<span class="section">${escapeHtml(line)}</span>`;
    }
    if (isChordLine(trimmed)) {
      const transposed = steps !== 0 ? transposeChord(line, steps) : line;
      return `<span class="chord">${escapeHtml(transposed)}</span>`;
    }
    // Ligne de paroles : encadrer dans un span .line pour la couleur chanteur
    if (songId) {
      const lineKey = lineKeyOf(line);
      const singer = assignments[lineKey] || 0;
      return `<span class="line line-singer-${singer}" data-line="${escapeHtml(lineKey)}">${escapeHtml(line)}</span>`;
    }
    return escapeHtml(line);
  }).join("\n");
}

function renderSongList() {
  const el = document.getElementById("song-list");
  if (state.filtered.length === 0) {
    el.innerHTML = `<div style="padding:16px;color:var(--text-dim);font-size:13px;">Aucun résultat.</div>`;
    document.getElementById("count-songs").textContent = "0";
    return;
  }
  el.innerHTML = state.filtered.map(s => {
    const active = s.id === state.currentSongId ? "active" : "";
    const capo = s.capo ? `<span class="badge">Capo ${s.capo}</span>` : "";
    const key = songKey(s) ? `<span class="badge">${escapeHtml(songKey(s))}</span>` : "";
    const style = s.style ? `<span class="badge style">${escapeHtml(s.style)}</span>` : "";
    return `<div class="song-item ${active}" onclick="selectSong(${s.id})">
      <div class="title">${escapeHtml(s.title)}</div>
      <div class="meta">${s.artiste ? `<span class="artist">${escapeHtml(s.artiste)}</span>` : ""}${key}${capo}${style}</div>
      <span class="song-del" onclick="event.stopPropagation(); deleteSongById(${s.id})" title="Supprimer cette chanson" aria-label="Supprimer">🗑</span>
    </div>`;
  }).join("");
  document.getElementById("count-songs").textContent = state.filtered.length;
}

function renderSong() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  const titleEl = document.getElementById("song-title");
  const bodyEl = document.getElementById("song-body");
  const pencil = document.getElementById("title-edit");
  const setVisible = (id, html) => { const e = document.getElementById(id); if (html) { e.innerHTML = html; e.style.display = ""; } else { e.style.display = "none"; } };
  const bodyEditPencil = document.getElementById("body-edit");
  if (!song) {
    titleEl.textContent = "—";
    if (pencil) pencil.style.display = "none";
    if (bodyEditPencil) bodyEditPencil.style.display = "none";
    const infoEditEl0 = document.getElementById("info-edit"); if (infoEditEl0) infoEditEl0.style.display = "none";
    const singersBtnEl = document.getElementById("singers-btn"); if (singersBtnEl) singersBtnEl.style.display = "none";
    ["meta-artiste","meta-key","meta-capo","meta-style","meta-bpm","meta-chords"].forEach(id => setVisible(id, ""));
    ["listen-spotify","listen-other","listen-search"].forEach(id => { const b = document.getElementById(id); if (b) b.style.display = "none"; });
    bodyEl.innerHTML = `<div class="empty">Sélectionnez une chanson dans la liste à gauche.<div class="hint">Astuce : utilisez ＋ pour ajouter une chanson en collant son texte, ou ⚙ pour connecter votre fichier .xlsx Drive.</div></div>`;
    showScrollControls(false);
    return;
  }
  titleEl.textContent = song.title;
  if (pencil) pencil.style.display = "";
  if (bodyEditPencil) bodyEditPencil.style.display = "";
  const infoEditEl1 = document.getElementById("info-edit"); if (infoEditEl1) infoEditEl1.style.display = "";
  const singersBtnEl2 = document.getElementById("singers-btn"); if (singersBtnEl2) singersBtnEl2.style.display = "";
  if (state.singerAssignMode) renderSingerPalette();
  showScrollControls(true);
  updateScrollUI();
  const keyDisplay = songKey(song) ? (state.transposeSteps !== 0 ? transposeChord(songKey(song), state.transposeSteps) + " (transp.)" : songKey(song)) : "";
  setVisible("meta-artiste", song.artiste ? `<strong>${escapeHtml(song.artiste)}</strong>` : "");
  updateListenButtons(song);
  setVisible("meta-key", keyDisplay ? `Tonalité : <strong>${keyDisplay}</strong>` : "");
  setVisible("meta-capo", song.capo ? `Capo : <strong>${song.capo}</strong>` : "");
  setVisible("meta-style", song.style ? `<strong>${escapeHtml(song.style)}</strong>` : "");
  setVisible("meta-bpm", song.bpm ? `<strong>${escapeHtml(String(song.bpm))} BPM</strong>` : "");
  setVisible("meta-chords", song.n_chords ? `<strong>${song.n_chords}</strong> accords` : "");

  document.documentElement.style.setProperty("--song-fs", state.fontSize + "px");

  // Extra meta block (only if some fields are filled)
  const extras = [];
  if (song.genre) extras.push(`<div><span>Genre</span><strong>${escapeHtml(song.genre)}</strong></div>`);
  if (song.rythme) extras.push(`<div><span>Rythme</span><strong>${escapeHtml(song.rythme)}</strong></div>`);
  if (song.duree) extras.push(`<div><span>Durée</span><strong>${escapeHtml(String(song.duree))}</strong></div>`);
  if (song.niveau) extras.push(`<div><span>Niveau</span><strong>${escapeHtml(song.niveau)}</strong></div>`);
  if (song.tags) extras.push(`<div><span>Tags</span><strong>${escapeHtml(song.tags)}</strong></div>`);
  // notes handled by dedicated notes panel below the song body
  const extraHtml = extras.length ? `<div class="extra-meta">${extras.join("")}</div>` : "";

  bodyEl.innerHTML = extraHtml + `<pre class="song-body">${renderSongBody(song.body_full, state.transposeSteps, song.id)}</pre>` + buildNotesPanel(song);
  refreshSingersBtnLabel();
  fitSongBodies();
}

function renderSetlist() {
  const sel = document.getElementById("setlist-select");
  sel.innerHTML = Object.values(state.setlists)
    .sort((a,b) => a.name.localeCompare(b.name))
    .map(sl => `<option value="${sl.id}" ${sl.id === state.currentSetlistId ? "selected" : ""}>${escapeHtml(sl.name)}</option>`).join("");
  const sl = state.setlists[state.currentSetlistId];
  if (!sl) return;
  document.getElementById("setlist-name").value = sl.name;
  const itemsEl = document.getElementById("setlist-items");
  if (sl.songIds.length === 0) {
    itemsEl.innerHTML = `<div class="setlist-empty">Aucune chanson.<br>Cliquez sur « + Setlist » dans une chanson.</div>`;
    return;
  }
  itemsEl.innerHTML = sl.songIds.map((sid, i) => {
    const song = state.songs.find(s => s.id === sid);
    if (!song) return `<div class="setlist-item" style="color:var(--text-dim)"><span class="num">${i+1}.</span><span class="title"><em>(chanson #${sid} introuvable)</em></span><span class="remove" onclick="event.stopPropagation();removeFromSetlist(${i});">✕</span></div>`;
    const cls = (sid === state.currentSongId) ? "current" : "";
    const bpmBadge = song.bpm ? `<span class="bpm-badge">${song.bpm}</span>` : "";
    return `<div class="setlist-item ${cls}" draggable="true" data-idx="${i}" onclick="selectSong(${sid})"><span class="num">${i+1}.</span><span class="title">${escapeHtml(song.title)}</span>${bpmBadge}<span class="remove" onclick="event.stopPropagation();removeFromSetlist(${i});">✕</span></div>`;
  }).join("");
  let dragIdx = null;
  itemsEl.querySelectorAll(".setlist-item").forEach(it => {
    it.addEventListener("dragstart", () => { dragIdx = parseInt(it.dataset.idx); it.classList.add("dragging"); });
    it.addEventListener("dragend", () => it.classList.remove("dragging"));
    it.addEventListener("dragover", e => e.preventDefault());
    it.addEventListener("drop", e => {
      e.preventDefault();
      const dropIdx = parseInt(it.dataset.idx);
      if (dragIdx === null || dragIdx === dropIdx) return;
      const sl2 = state.setlists[state.currentSetlistId];
      const moved = sl2.songIds.splice(dragIdx, 1)[0];
      sl2.songIds.splice(dropIdx, 0, moved);
      saveToStorage();
      renderSetlist();
    });
  });
}

function renderSourceInfo() {
  const el = document.getElementById("source-info");
  let label = "";
  if (state.source === "interne") label = `<span class="warn">●</span> Source : base interne`;
  else if (state.source === "drive") label = `<span class="ok">●</span> Source : Google Drive · ${state.lastLoaded ? state.lastLoaded.toLocaleTimeString("fr-FR") : ""}`;
  else if (state.source === "fichier") label = `<span class="ok">●</span> Source : fichier local · ${state.lastLoaded ? state.lastLoaded.toLocaleTimeString("fr-FR") : ""}`;
  else if (state.source === "gas") label = `<span class="ok">●</span> Sheet (Apps Script) · ${state.gasSyncing ? "sync…" : "synchro ✓"}`;
  if (state.dirty) label += ` <strong style="color:var(--warn)">· modifié ✱</strong>`;
  label += ` <span style="opacity:0.5;font-size:10px">· v.${BUILD_VERSION}</span>`;
  el.innerHTML = label;
}

function refreshAll() {
  populateFilterOptions();
  applyFilters();
  renderSong();
  renderSetlist();
  renderSourceInfo();
  renderDirtyBar();
}

function markDirty() {
  state.dirty = true;
  renderSourceInfo();
  renderDirtyBar();
}

function renderDirtyBar() {
  const el = document.getElementById("save-bar");
  // En mode Apps Script les changements sont push\u00e9s en direct \u2192 pas de barre manuelle
  if (state.source === "gas") { el.classList.remove("show"); renderSourceInfo(); return; }
  if (state.dirty) el.classList.add("show"); else el.classList.remove("show");
}

// ===== NOTES PANEL =====
function buildNotesPanel(song) {
  if (!song) return "";
  const isStage = document.getElementById("app").classList.contains("stage");
  const editBtn = isStage ? "" : `<button class="notes-edit-btn stage-hide" onclick="openNotesEdit()">✏ Modifier</button>`;
  const notesContent = song.notes
    ? `<div class="notes-content">${escapeHtml(song.notes)}</div>`
    : `<div class="notes-empty">Aucune note pour cette chanson${isStage ? "." : " — cliquez sur ✏ Modifier pour en ajouter."}</div>`;
  return `<div class="notes-panel" id="notes-panel">
    <div class="notes-header"><span>✏ Notes</span>${editBtn}</div>
    <div id="notes-view">${notesContent}</div>
  </div>`;
}

function openNotesEdit() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) return;
  const viewEl = document.getElementById("notes-view");
  if (!viewEl) return;
  viewEl.innerHTML = `<textarea class="notes-textarea" id="notes-textarea" placeholder="Tes notes pour cette chanson (conseils de jeu, rappels, arrangements...)">${escapeHtml(song.notes || "")}</textarea>
    <div class="notes-actions">
      <button class="btn primary" onclick="saveNotes()">💾 Sauvegarder</button>
      <button class="btn ghost" onclick="renderSong()">Annuler</button>
    </div>`;
  const ta = document.getElementById("notes-textarea");
  if (ta) { ta.focus(); ta.selectionStart = ta.value.length; }
}

function saveNotes() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) return;
  const ta = document.getElementById("notes-textarea");
  if (!ta) return;
  song.notes = ta.value;
  if (state.source === "gas") gasPushSong(song); else markDirty();
  renderSong();
}

// ===== TITLE EDITING =====
function editCurrentTitle() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) return;
  const next = prompt("Nouveau titre :", song.title);
  if (next === null) return;
  const trimmed = next.trim();
  if (!trimmed) { alert("Le titre ne peut pas être vide."); return; }
  if (trimmed === song.title) return;
  song.title = trimmed;
  if (state.source === "gas") gasPushSong(song); else markDirty();
  renderSongList();
  renderSong();
  renderSetlist();
}

// ===== ADD SONG (parser + modal) =====
const SECTION_HEADER_JS_RE = /^\s*\[?\s*(intro|verse(?:\s*\d+)?|couplet(?:\s*\d+)?|chorus(?:\s*\d+)?|refrain(?:\s*\d+)?|pre[- ]?chorus|pre[- ]?refrain|bridge|pont|brid\.?|interlude|outro|hook|solo|fin)\s*:?\s*\]?\s*$/i;

function classifySectionType(header) {
  const h = header.toLowerCase().replace(/[\[\]:]/g, "").trim();
  if (h.includes("pre") && (h.includes("chorus") || h.includes("refrain"))) return "pre_chorus";
  if (h.includes("intro")) return "intro";
  if (h.includes("outro") || h === "fin") return "outro";
  if (h.includes("chorus") || h.includes("refrain")) return "chorus";
  if (h.includes("bridge") || h.includes("pont") || h.startsWith("brid")) return "bridge";
  if (h.includes("verse") || h.includes("couplet")) return "verse";
  if (h.includes("interlude")) return "interlude";
  if (h.includes("hook")) return "hook";
  if (h.includes("solo")) return "solo";
  return "other";
}

function parseSongText(rawText) {
  if (!rawText || !rawText.trim()) return null;
  const lines = rawText.split("\n");

  // 1. Find title: first non-empty line that is NOT a chord line and NOT a section header
  let titleIdx = null;
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (!t) continue;
    if (SECTION_HEADER_JS_RE.test(t)) continue;
    if (isChordLine(t)) continue;
    if (t.length > 80) continue;
    titleIdx = i;
    break;
  }
  if (titleIdx === null) {
    for (let i = 0; i < lines.length; i++) if (lines[i].trim()) { titleIdx = i; break; }
  }
  if (titleIdx === null) return null;

  // 2. Extract title + capo
  let rawTitle = lines[titleIdx].trim().replace(/^!+/, "");
  let capo = null;
  const capoMatch = rawTitle.match(/\bcapo\s*(\d+)\b/i);
  if (capoMatch) {
    capo = parseInt(capoMatch[1], 10);
    rawTitle = rawTitle.replace(/\bcapo\s*\d+\b/i, "");
  }
  rawTitle = rawTitle.replace(/\s+/g, " ").replace(/^[-_:\s]+|[-_:\s]+$/g, "");
  if (!rawTitle) return null;

  // 3. Body
  const body = lines.slice(titleIdx + 1).join("\n").replace(/^\n+/, "").replace(/\n{3,}/g, "\n\n");

  // 4. Extract chords (ordered + unique)
  const chords = [];
  const seen = new Set();
  for (const line of body.split("\n")) {
    const cleaned = line.replace(/[|()\[\]:]/g, " ");
    const toks = cleaned.split(/\s+/).filter(Boolean);
    if (!toks.length) continue;
    const chordTokens = toks.filter(t => CHORD_TOKEN_RE.test(t));
    if (chordTokens.length / toks.length >= 0.6) {
      for (const t of toks) if (CHORD_TOKEN_RE.test(t) && !seen.has(t)) { seen.add(t); chords.push(t); }
    }
  }

  // 5. Split sections
  const sections = { intro: [], verse: [], chorus: [], pre_chorus: [], bridge: [], outro: [], interlude: [], hook: [], solo: [], other: [] };
  let cur = "intro";
  let buf = [];
  const flush = () => {
    const txt = buf.join("\n").replace(/^\n+|\n+$/g, "");
    if (txt.trim()) (sections[cur] || sections.other).push(txt);
    buf = [];
  };
  for (const line of body.split("\n")) {
    if (SECTION_HEADER_JS_RE.test(line.trim())) {
      flush();
      cur = classifySectionType(line.trim());
    } else {
      buf.push(line);
    }
  }
  flush();
  const join = arr => arr.join("\n\n---\n\n");

  return {
    title: rawTitle,
    capo,
    key_detected: "",
    key_played: "",
    chords,
    n_chords: chords.length,
    n_lines: body.split("\n").filter(l => l.trim()).length,
    intros: join(sections.intro),
    verses: join(sections.verse),
    pre_chorus: join(sections.pre_chorus),
    choruses: join(sections.chorus),
    bridges: join(sections.bridge),
    body_full: body,
    style: "", genre: "", rythme: "", bpm: "", duree: "", tags: "", niveau: "", notes: "", artiste: "", spotify: "", lien_ecoute: "",
  };
}

function openAddSong() {
  state.editingId = null;
  document.getElementById("add-modal-title").textContent = "Ajouter une chanson";
  document.getElementById("add-save-btn").textContent = "✓ Ajouter au répertoire";
  document.getElementById("paste-area").value = "";
  document.getElementById("parse-preview").innerHTML = "";
  document.getElementById("add-modal").classList.add("show");
  setTimeout(() => document.getElementById("paste-area").focus(), 50);
}

function editCurrentBody() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) return;
  state.editingId = song.id;
  document.getElementById("add-modal-title").textContent = `Modifier « ${song.title} »`;
  document.getElementById("add-save-btn").textContent = "💾 Enregistrer les modifications";
  // Reconstruct editable text: title (with optional Capo) on first line, then body
  const titleLine = song.title + (song.capo ? ` Capo ${song.capo}` : "");
  document.getElementById("paste-area").value = titleLine + "\n\n" + (song.body_full || "");
  updateParsePreview();
  document.getElementById("add-modal").classList.add("show");
  setTimeout(() => document.getElementById("paste-area").focus(), 50);
}

function closeAddSong() { document.getElementById("add-modal").classList.remove("show"); }

function updateParsePreview() {
  const txt = document.getElementById("paste-area").value;
  const previewEl = document.getElementById("parse-preview");
  if (!txt.trim()) { previewEl.innerHTML = ""; return; }
  const parsed = parseSongText(txt);
  if (!parsed) {
    previewEl.innerHTML = `<div class="preview-card err">⚠ Aucun titre détecté. Mettez le titre en première ligne.</div>`;
    return;
  }
  const sections = [];
  if (parsed.intros) sections.push("Intro");
  if (parsed.verses) sections.push("Couplet(s)");
  if (parsed.pre_chorus) sections.push("Pré-refrain");
  if (parsed.choruses) sections.push("Refrain");
  if (parsed.bridges) sections.push("Bridge");
  previewEl.innerHTML = `
    <div class="preview-card">
      <div class="row"><span>Titre :</span><strong>${escapeHtml(parsed.title)}</strong></div>
      <div class="row"><span>Capo :</span><strong>${parsed.capo || "—"}</strong>
        <span>Accords :</span><strong>${parsed.n_chords || 0}</strong>
        <span>Lignes :</span><strong>${parsed.n_lines}</strong></div>
      <div class="row"><span>Sections :</span><strong>${sections.join(", ") || "—"}</strong></div>
      ${parsed.chords.length ? `<div class="row"><span>Suite d'accords :</span><strong style="font-family:monospace">${escapeHtml(parsed.chords.join(", "))}</strong></div>` : ""}
    </div>`;
}

function saveNewSong() {
  const txt = document.getElementById("paste-area").value;
  const parsed = parseSongText(txt);
  if (!parsed) { alert("Impossible de détecter une chanson valide. Vérifiez que le titre est en première ligne."); return; }

  let target = null;
  if (state.editingId) {
    // Edit mode: update existing song; preserve metadata not derived from raw text
    const existing = state.songs.find(s => s.id === state.editingId);
    if (!existing) { alert("Chanson introuvable."); state.editingId = null; return; }
    existing.title = parsed.title;
    existing.capo = parsed.capo || existing.capo || null;
    existing.chords = parsed.chords;
    existing.n_chords = parsed.n_chords;
    existing.n_lines = parsed.n_lines;
    existing.intros = parsed.intros;
    existing.verses = parsed.verses;
    existing.pre_chorus = parsed.pre_chorus;
    existing.choruses = parsed.choruses;
    existing.bridges = parsed.bridges;
    existing.body_full = parsed.body_full;
    existing.key_detected = parsed.key_detected || existing.key_detected || "";
    state.currentSongId = existing.id;
    target = existing;
  } else {
    // Add mode
    const nextId = (state.songs.reduce((m, s) => Math.max(m, s.id || 0), 0) || 0) + 1;
    const newSong = { ...parsed, id: state.source === "gas" ? null : nextId };
    state.songs.push(newSong);
    state.currentSongId = newSong.id || nextId;
    target = newSong;
  }

  state.editingId = null;
  if (state.source === "gas" && target) {
    gasPushSong(target).then(() => {
      // After GAS sets the new id, re-sync currentSongId for additions
      if (target.id) state.currentSongId = target.id;
      refreshAll();
    });
  } else {
    markDirty();
  }
  closeAddSong();
  refreshAll();
}

// ===== SAVE AS XLSX =====
const XLSX_HEADERS = ["ID","Titre","Doublon","Style","Genre","Rythme","BPM","Tonalité (auto)","Tonalité (jouée)","Capo","Durée","Tags concert","Niveau","Accords utilisés","Nb accords","Nb lignes","Intro","Couplets","Pré-refrain","Refrain","Bridge","Body complet","Notes","Artiste","Lien Spotify","Lien écoute"];

function songToRow(s) {
  return {
    "ID": s.id,
    "Titre": s.title || "",
    "Doublon": "",
    "Style": s.style || "",
    "Genre": s.genre || "",
    "Rythme": s.rythme || "",
    "BPM": s.bpm || "",
    "Tonalité (auto)": s.key_detected || "",
    "Tonalité (jouée)": s.key_played || "",
    "Capo": s.capo || "",
    "Durée": s.duree || "",
    "Tags concert": s.tags || "",
    "Niveau": s.niveau || "",
    "Accords utilisés": (s.chords || []).join(", "),
    "Nb accords": s.n_chords || 0,
    "Nb lignes": s.n_lines || 0,
    "Intro": s.intros || "",
    "Couplets": s.verses || "",
    "Pré-refrain": s.pre_chorus || "",
    "Refrain": s.choruses || "",
    "Bridge": s.bridges || "",
    "Body complet": s.body_full || "",
    "Notes": s.notes || "",
    "Artiste": s.artiste || "",
    "Lien Spotify": s.spotify || "",
    "Lien écoute": s.lien_ecoute || "",
  };
}

function saveAsXlsx() {
  const rows = state.songs.map(songToRow);
  const ws = XLSX.utils.json_to_sheet(rows, { header: XLSX_HEADERS });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Répertoire");
  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Repertoire_GrooveIn_${stamp}.xlsx`);
  state.dirty = false;
  renderSourceInfo();
  renderDirtyBar();
  alert("Fichier téléchargé. Pour mettre à jour Drive : remplacez le fichier d'origine par celui que vous venez de télécharger (clic droit sur Drive → Gérer les versions → Importer une nouvelle version).");
}

function discardChanges() {
  if (!confirm("Annuler toutes les modifications non sauvegardées ?")) return;
  // Re-load from Drive if available, else from embedded
  if (state.driveUrl && state.source !== "interne") {
    loadFromDriveUrl();
  } else if (state.source === "interne") {
    state.songs = [...EMBEDDED_SONGS];
  }
  state.dirty = false;
  refreshAll();
}

// ===== ACTIONS =====
function selectSong(id) {
  state.currentSongId = id;
  state.transposeSteps = 0;
  resetScroll();
  // Sur mobile : refermer la sidebar automatiquement après sélection
  if (window.innerWidth <= 1100) document.getElementById("app").classList.remove("sidebar-open", "setlist-open");
  renderSongList();
  renderSong();
  renderSetlist();
}
function changeFontSize(d) { state.fontSize = Math.max(10, Math.min(36, state.fontSize + d * 2)); saveToStorage(); renderSong(); }
function transpose(s) { state.transposeSteps += s; renderSong(); }
function resetTranspose() { state.transposeSteps = 0; renderSong(); }
function toggleStage() {
  const app = document.getElementById("app");
  app.classList.toggle("stage");
  const stageOn = app.classList.contains("stage");
  if (stageOn) {
    state.fontSize = Math.max(state.fontSize, 18);
    enterFullscreen();
  } else {
    exitFullscreen();
  }
  renderSong();
}

// ===== PLEIN ÉCRAN (Fullscreen API) =====
function enterFullscreen() {
  const el = document.documentElement;
  const req = el.requestFullscreen || el.webkitRequestFullscreen
              || el.mozRequestFullScreen || el.msRequestFullscreen;
  if (req) { try { req.call(el); } catch (e) { console.warn("Fullscreen denied:", e); } }
}
function exitFullscreen() {
  const exit = document.exitFullscreen || document.webkitExitFullscreen
               || document.mozCancelFullScreen || document.msExitFullscreen;
  if (exit && (document.fullscreenElement || document.webkitFullscreenElement)) {
    try { exit.call(document); } catch (e) {}
  }
}
function isFullscreen() {
  return !!(document.fullscreenElement || document.webkitFullscreenElement
            || document.mozFullScreenElement || document.msFullscreenElement);
}
function toggleFullscreen() {
  if (isFullscreen()) exitFullscreen(); else enterFullscreen();
}
// Si l'utilisateur sort du plein écran via le bouton du navigateur ou Échap,
// on quitte aussi le mode Scène pour rester cohérent
["fullscreenchange", "webkitfullscreenchange", "mozfullscreenchange", "MSFullscreenChange"].forEach(ev => {
  document.addEventListener(ev, () => {
    if (!isFullscreen()) {
      const app = document.getElementById("app");
      if (app.classList.contains("stage")) { app.classList.remove("stage"); renderSong(); }
      if (app.classList.contains("setlist-mode")) { app.classList.remove("setlist-mode"); renderSong(); }
    }
  });
});


// ===== AJUSTEMENT AUTOMATIQUE DES PAROLES À LA LARGEUR =====
// La taille choisie (A− / A+) est un MAXIMUM : si la ligne la plus longue dépasse la largeur
// disponible, on réduit la police (plancher 9px). Les accords restent alignés (police à chasse fixe).
const FIT_MIN_PX = 9;
function toggleFit() { state.fitWidth = !state.fitWidth; saveToStorage(); fitSongBodies(); }
function fitSongBodies() {
  const fb = document.getElementById("fit-btn");
  if (fb) fb.style.cssText = state.fitWidth ? "background:var(--accent);color:var(--bg);border-color:var(--accent)" : "";
  document.querySelectorAll("pre.song-body").forEach(pre => {
    pre.style.removeProperty("--song-fs-fit");
    if (!state.fitWidth) return;
    const host = pre.parentElement;
    if (!host) return;
    const cs = getComputedStyle(host);
    const avail = host.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    if (avail <= 0) return;
    const fs = parseFloat(getComputedStyle(pre).fontSize);
    const maxChars = (pre.textContent || "").split("\n").reduce((m, l) => Math.max(m, l.length), 0);
    if (!maxChars || !fs) return;
    // largeur d'un caractère = mesure réelle de la police à chasse fixe
    const probe = document.createElement("span");
    probe.style.cssText = "position:absolute;visibility:hidden;white-space:pre;font:" + getComputedStyle(pre).font;
    probe.textContent = "0".repeat(100);
    document.body.appendChild(probe);
    const ratio = probe.getBoundingClientRect().width / 100 / fs;
    probe.remove();
    const needed = maxChars * ratio * fs;
    if (needed > avail) {
      const fit = Math.max(FIT_MIN_PX, Math.floor((fs * avail / needed) * 10) / 10);
      pre.style.setProperty("--song-fs-fit", fit + "px");
    }
  });
}
let _fitTimer = null;
function scheduleFit() { clearTimeout(_fitTimer); _fitTimer = setTimeout(fitSongBodies, 120); }
window.addEventListener("resize", scheduleFit);
window.addEventListener("orientationchange", scheduleFit);
if (document.fonts && document.fonts.ready) document.fonts.ready.then(fitSongBodies);

// ===== SIDEBAR TOGGLE =====
function toggleSidebar() {
  const app = document.getElementById("app");
  if (window.innerWidth <= 1100) {
    app.classList.remove("setlist-open");
    app.classList.toggle("sidebar-open");
  } else {
    app.classList.toggle("sidebar-collapsed");
  }
}
function toggleSetlistPanel() {
  const app = document.getElementById("app");
  app.classList.remove("sidebar-open");
  app.classList.toggle("setlist-open");
}

// ===== AUTO-SCROLL (téléprompteur) =====
const scrollState = { playing: false, speed: 25, lastTs: null, raf: null, residual: 0 };

function showScrollControls(show) {
  const el = document.getElementById("scroll-controls");
  if (el) el.style.display = show ? "flex" : "none";
}

function updateScrollUI() {
  const wrap = document.getElementById("scroll-controls");
  const tgl = document.getElementById("scroll-toggle");
  const lbl = document.getElementById("scroll-speed-label");
  if (!wrap) return;
  if (scrollState.playing) { wrap.classList.add("playing"); tgl.textContent = "⏸"; tgl.title = "Pause (Espace)"; }
  else { wrap.classList.remove("playing"); tgl.textContent = "▶"; tgl.title = "Lecture (Espace)"; }
  if (lbl) lbl.textContent = String(scrollState.speed);
}

function scrollFrame(ts) {
  if (!scrollState.playing) { scrollState.raf = null; return; }
  const el = document.getElementById("song-body");
  if (!el) { scrollState.playing = false; updateScrollUI(); return; }
  if (scrollState.lastTs != null) {
    const dt = (ts - scrollState.lastTs) / 1000;
    const px = scrollState.speed * dt + scrollState.residual;
    const whole = Math.floor(px);
    scrollState.residual = px - whole;
    el.scrollTop += whole;
    if (el.scrollTop + el.clientHeight >= el.scrollHeight - 1) {
      scrollState.playing = false;
      updateScrollUI();
      scrollState.raf = null;
      return;
    }
  }
  scrollState.lastTs = ts;
  scrollState.raf = requestAnimationFrame(scrollFrame);
}

function toggleAutoScroll() {
  scrollState.playing = !scrollState.playing;
  scrollState.lastTs = null;
  scrollState.residual = 0;
  if (scrollState.playing) scrollState.raf = requestAnimationFrame(scrollFrame);
  updateScrollUI();
}

function changeScrollSpeed(delta) {
  scrollState.speed = Math.max(5, Math.min(200, scrollState.speed + delta));
  updateScrollUI();
}

function resetScroll() {
  scrollState.playing = false;
  scrollState.lastTs = null;
  scrollState.residual = 0;
  if (scrollState.raf) { cancelAnimationFrame(scrollState.raf); scrollState.raf = null; }
  const el = document.getElementById("song-body");
  if (el) el.scrollTop = 0;
  updateScrollUI();
}
function addCurrentToSetlist() {
  if (!state.currentSongId) return;
  const sl = state.setlists[state.currentSetlistId];
  if (sl.songIds.includes(state.currentSongId)) return;
  sl.songIds.push(state.currentSongId);
  saveToStorage(); renderSetlist();
}
function removeFromSetlist(idx) {
  const sl = state.setlists[state.currentSetlistId];
  sl.songIds.splice(idx, 1);
  saveToStorage(); renderSetlist();
}
function newSetlist() {
  const name = prompt("Nom de la nouvelle setlist :", "Concert " + new Date().toLocaleDateString("fr-FR"));
  if (!name) return;
  const id = "sl_" + Date.now();
  state.setlists[id] = { id, name, songIds: [], notes: "" };
  state.currentSetlistId = id;
  saveToStorage(); renderSetlist();
}
function loadSetlist(id) { state.currentSetlistId = id; saveToStorage(); renderSetlist(); }

// ===== PLAY SETLIST (lecture continue de tous les morceaux enchaînés) =====
function playSetlist() {
  const sl = state.setlists[state.currentSetlistId];
  if (!sl || !sl.songIds.length) {
    alert("Cette setlist est vide. Ajoutez des chansons avant de la lancer.");
    return;
  }
  const app = document.getElementById("app");
  app.classList.add("setlist-mode");
  enterFullscreen();
  renderSetlistFlow();
  fitSongBodies();
  // Le scroll part du haut
  setTimeout(() => { const c = document.getElementById("song-body"); if (c) c.scrollTop = 0; }, 50);
}

function exitSetlistPlay() {
  const app = document.getElementById("app");
  app.classList.remove("setlist-mode");
  exitFullscreen();
  resetScroll();
  renderSong();
}

function renderSetlistFlow() {
  const sl = state.setlists[state.currentSetlistId];
  const titleEl = document.getElementById("song-title");
  const bodyEl = document.getElementById("song-body");
  if (!sl || !sl.songIds.length) { exitSetlistPlay(); return; }

  titleEl.textContent = `🎤 ${sl.name} — ${sl.songIds.length} chanson${sl.songIds.length > 1 ? "s" : ""}`;
  document.documentElement.style.setProperty("--song-fs", state.fontSize + "px");

  const blocks = [];
  sl.songIds.forEach((sid, idx) => {
    const song = state.songs.find(s => s.id === sid);
    if (!song) {
      blocks.push(`<div class="song-block"><div class="song-block-header"><span class="song-num">${idx+1}.</span><span class="song-block-title" style="color:var(--err)">Chanson introuvable (id ${escapeHtml(String(sid))})</span></div></div>`);
      return;
    }
    const meta = [];
    const keyDisplay = songKey(song) ? (state.transposeSteps !== 0 ? transposeChord(songKey(song), state.transposeSteps) + " (transp.)" : songKey(song)) : "";
    if (keyDisplay) meta.push(`Tonalité <strong>${keyDisplay}</strong>`);
    if (song.capo) meta.push(`Capo <strong>${song.capo}</strong>`);
    if (song.bpm) meta.push(`<strong>${escapeHtml(String(song.bpm))} BPM</strong>`);
    if (song.style) meta.push(`<strong>${escapeHtml(song.style)}</strong>`);
    blocks.push(`
      <div class="song-block">
        <div class="song-block-header">
          <span class="song-num">${idx + 1}.</span>
          <div style="flex:1">
            <div class="song-block-title">${escapeHtml(song.title)}</div>
            ${meta.length ? `<div class="song-meta">${meta.join(" · ")}</div>` : ""}
          </div>
        </div>
        <pre class="song-body">${renderSongBody(song.body_full || "", state.transposeSteps, song.id)}</pre>
        ${idx < sl.songIds.length - 1 ? '<div class="song-sep"></div>' : ""}
      </div>`);
  });

  bodyEl.innerHTML = `<div class="setlist-flow">${blocks.join("")}</div>`;
  showScrollControls(true);
  updateScrollUI();
}
function renameSetlist(name) {
  const sl = state.setlists[state.currentSetlistId];
  if (sl && name) { sl.name = name; saveToStorage(); renderSetlist(); }
}
function deleteSetlist() {
  if (Object.keys(state.setlists).length <= 1) { alert("Vous devez garder au moins une setlist."); return; }
  if (!confirm("Supprimer la setlist « " + state.setlists[state.currentSetlistId].name + " » ?")) return;
  delete state.setlists[state.currentSetlistId];
  state.currentSetlistId = Object.keys(state.setlists)[0];
  saveToStorage(); renderSetlist();
}
function exportSetlist() {
  const sl = state.setlists[state.currentSetlistId];
  const lines = [`Setlist : ${sl.name}`, ""];
  sl.songIds.forEach((sid, i) => {
    const s = state.songs.find(x => x.id === sid);
    if (!s) return;
    lines.push(`${i+1}. ${s.title}${songKey(s) ? "  ["+songKey(s)+"]" : ""}${s.capo ? "  (Capo "+s.capo+")" : ""}`);
  });
  const text = lines.join("\n");
  if (navigator.clipboard) navigator.clipboard.writeText(text).then(() => alert("Setlist copiée ✓\n\n" + text)).catch(() => alert(text));
  else alert(text);
}

// ===== SETTINGS MODAL =====
function openSettings() {
  document.getElementById("drive-url").value = state.driveUrl || "";
  document.getElementById("gas-url").value = state.gasUrl || "";
  document.getElementById("settings-modal").classList.add("show");
}
function closeSettings() { document.getElementById("settings-modal").classList.remove("show"); }
document.getElementById("settings-modal").addEventListener("click", e => {
  if (e.target.id === "settings-modal") closeSettings();
});
document.getElementById("add-modal").addEventListener("click", e => {
  if (e.target.id === "add-modal") closeAddSong();
});
document.getElementById("singers-modal").addEventListener("click", e => {
  if (e.target.id === "singers-modal") closeSingersModal();
});
document.getElementById("paste-area").addEventListener("input", updateParsePreview);

// ===== FILTERS =====
function applyFilters() {
  const q = document.getElementById("search").value.toLowerCase().trim();
  const fk = document.getElementById("filter-key").value;
  const fc = document.getElementById("filter-capo").value;
  const fs = document.getElementById("filter-style").value;
  const ft = document.getElementById("filter-tag").value.toLowerCase();
  state.filtered = state.songs.filter(s => {
    if (fk && (songKey(s) || "") !== fk) return false;
    if (fc && String(s.capo || "") !== fc) return false;
    if (fs && (s.style || "") !== fs) return false;
    if (ft && !(s.tags || "").toLowerCase().includes(ft)) return false;
    if (!q) return true;
    const hay = [s.title, s.artiste, s.key_detected, s.key_played, (s.chords || []).join(" "), s.body_full, s.style, s.genre, s.tags, s.notes].join(" ").toLowerCase();
    return hay.includes(q);
  });
  renderSongList();
}

function populateFilterOptions() {
  const fillSelect = (id, values) => {
    const sel = document.getElementById(id);
    const current = sel.value;
    const placeholder = sel.options[0]?.outerHTML || "";
    sel.innerHTML = placeholder + values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    sel.value = current;
  };
  fillSelect("filter-key", [...new Set(state.songs.map(s => songKey(s)).filter(Boolean))].sort());
  fillSelect("filter-capo", [...new Set(state.songs.map(s => s.capo).filter(Boolean))].sort((a,b)=>a-b).map(String));
  fillSelect("filter-style", [...new Set(state.songs.map(s => s.style).filter(Boolean))].sort());
  // Tags: split comma-separated values
  const tagSet = new Set();
  state.songs.forEach(s => (s.tags || "").split(/[,;]/).map(t => t.trim()).filter(Boolean).forEach(t => tagSet.add(t)));
  fillSelect("filter-tag", [...tagSet].sort());
}


// ===== FICHE INFOS (genre, style, rythme, BPM, capo, tonalité jouée, durée, niveau, tags) =====
// Tonalité affichée = celle que tu as saisie (jouée) si elle existe, sinon la détection automatique.
function songKey(s) { return (s && (s.key_played || s.key_detected)) || ""; }

const REF_DEFAULTS = {
  style:  ["Variété française", "Pop", "Rock", "Soul / R&B", "Funk", "Jazz", "Reggae", "Disco", "Latin", "Chanson", "Ballade"],
  genre:  ["Français", "Anglais", "Hébreu", "Espagnol", "Instrumental"],
  rythme: ["4/4 medium", "4/4 lent", "Bossa", "Reggae", "Slow", "Shuffle", "Disco", "Cha-cha", "Rumba", "Valse"],
  niveau: ["Maîtrisé", "À revoir", "En apprentissage"],
  key:    ["C","Cm","C#","C#m","D","Dm","D#","D#m","E","Em","F","Fm","F#","F#m","G","Gm","G#","G#m","A","Am","A#","A#m","B","Bm"],
};
// Valeurs déjà utilisées dans le répertoire d'abord (les plus fréquentes), puis les suggestions par défaut
function refValues(field, stateField) {
  const counts = {};
  state.songs.forEach(s => { const v = String(s[stateField || field] || "").trim(); if (v) counts[v] = (counts[v] || 0) + 1; });
  const used = Object.keys(counts).sort((a, b) => counts[b] - counts[a] || a.localeCompare(b));
  const lower = new Set(used.map(v => v.toLowerCase()));
  return used.concat((REF_DEFAULTS[field] || []).filter(v => !lower.has(v.toLowerCase())));
}
function fillDatalist(id, values) {
  document.getElementById(id).innerHTML = values.map(v => `<option value="${escapeHtml(v)}"></option>`).join("");
}
function fillInfoSelect(id, values, current, emptyLabel) {
  const vals = current && !values.includes(current) ? [current].concat(values) : values;
  document.getElementById(id).innerHTML = `<option value="">${emptyLabel}</option>` + vals.map(v => `<option value="${escapeHtml(String(v))}">${escapeHtml(String(v))}</option>`).join("");
  document.getElementById(id).value = current || "";
}
function parseTagList(str) { return String(str || "").split(/[,;]/).map(t => t.trim()).filter(Boolean); }

function openInfoModal() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) { alert("Sélectionnez d'abord une chanson."); return; }
  document.getElementById("info-title").textContent = song.title;
  document.getElementById("info-artiste").value = song.artiste || "";
  document.getElementById("info-spotify").value = song.spotify || "";
  document.getElementById("info-lien").value = song.lien_ecoute || "";
  document.getElementById("info-genre").value = song.genre || "";
  document.getElementById("info-style").value = song.style || "";
  document.getElementById("info-rythme").value = song.rythme || "";
  document.getElementById("info-bpm").value = song.bpm || "";
  document.getElementById("info-key").value = song.key_played || "";
  document.getElementById("info-duree").value = song.duree || "";
  document.getElementById("info-tags").value = song.tags || "";
  fillDatalist("dl-genre", refValues("genre"));
  fillDatalist("dl-style", refValues("style"));
  fillDatalist("dl-rythme", refValues("rythme"));
  fillDatalist("dl-key", REF_DEFAULTS.key);
  fillInfoSelect("info-capo", ["1","2","3","4","5","6","7","8","9","10","11","12"], song.capo ? String(song.capo) : "", "Sans capo");
  fillInfoSelect("info-niveau", refValues("niveau"), song.niveau || "", "—");
  document.getElementById("info-key-hint").textContent = song.key_detected ? `Détectée automatiquement : ${song.key_detected}` : "";
  const tagMap = new Map(); // dédoublonnage insensible à la casse (rock / Rock)
  state.songs.forEach(s => parseTagList(s.tags).forEach(t => { if (!tagMap.has(t.toLowerCase())) tagMap.set(t.toLowerCase(), t); }));
  ["mariage","cocktail","bal","concert intime","bar","festival"].forEach(t => { if (!tagMap.has(t.toLowerCase())) tagMap.set(t.toLowerCase(), t); });
  const allTags = tagMap.values();
  document.getElementById("info-tag-chips").innerHTML = [...allTags].sort((a, b) => a.localeCompare(b)).map(t =>
    `<button type="button" class="tag-chip" data-tag="${escapeHtml(t)}" onclick="toggleTagChip(this.dataset.tag)">${escapeHtml(t)}</button>`).join("");
  syncTagChips();
  document.getElementById("info-modal").classList.add("show");
}
function closeInfoModal() { document.getElementById("info-modal").classList.remove("show"); }
function syncTagChips() {
  const on = new Set(parseTagList(document.getElementById("info-tags").value).map(t => t.toLowerCase()));
  document.querySelectorAll("#info-tag-chips .tag-chip").forEach(c => c.classList.toggle("on", on.has(c.dataset.tag.toLowerCase())));
}
function toggleTagChip(tag) {
  const inp = document.getElementById("info-tags");
  const list = parseTagList(inp.value);
  const i = list.findIndex(t => t.toLowerCase() === tag.toLowerCase());
  if (i >= 0) list.splice(i, 1); else list.push(tag);
  inp.value = list.join(", ");
  syncTagChips();
}

function showToast(kind, msg) {
  const el = document.getElementById("toast");
  el.className = "show " + kind; el.textContent = msg;
  clearTimeout(showToast._t); showToast._t = setTimeout(() => el.classList.remove("show"), 2800);
}

async function gasPushFields(song, fields) {
  // N'envoie au Sheet QUE les champs modifiés (le texte de la chanson n'est pas retouché)
  if (state.source !== "gas" || !state.gasUrl || !song.id) return;
  state.gasSyncing = true; renderDirtyBar();
  try {
    await gasPost({ action: "update_song", id: song.id, fields });
    state.dirty = false;
    showToast("ok", "✓ Enregistré dans le Sheet");
  } catch (e) {
    console.error("Push infos GAS échoué :", e);
    showToast("err", "Sauvegarde Sheet échouée : " + e.message);
    state.dirty = true;
  } finally {
    state.gasSyncing = false; renderDirtyBar();
  }
}

function saveInfo() {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) return;
  const val = id => document.getElementById(id).value.trim();
  const bpmRaw = val("info-bpm"), capoRaw = val("info-capo");
  const bpm = bpmRaw === "" ? "" : Math.round(Number(bpmRaw));
  if (bpmRaw !== "" && !(bpm >= 30 && bpm <= 300)) { alert("BPM : entrez un nombre entre 30 et 300."); return; }
  const duree = val("info-duree");
  if (duree && !/^\d{1,2}:[0-5]\d$/.test(duree)) { alert("Durée : utilisez le format minutes:secondes, par exemple 3:45."); return; }
  // Liens : on accepte http(s) et les adresses « spotify:track:… » (converties en lien web)
  const cleanLink = (v) => {
    v = (v || "").trim();
    const m = v.match(/^spotify:(track|album|playlist|artist):([A-Za-z0-9]+)$/i);
    return m ? `https://open.spotify.com/${m[1].toLowerCase()}/${m[2]}` : v;
  };
  const spotify = cleanLink(val("info-spotify")), lien = cleanLink(val("info-lien"));
  const okLink = v => v === "" || /^https?:\/\/\S+$/i.test(v);
  if (!okLink(spotify) || !okLink(lien)) { alert("Lien : collez une adresse complète qui commence par https://"); return; }
  const fields = {
    artiste: val("info-artiste"), spotify: spotify, lien_ecoute: lien,
    genre: val("info-genre"), style: val("info-style"), rythme: val("info-rythme"),
    bpm: bpm, capo: capoRaw === "" ? "" : Number(capoRaw),
    key_played: val("info-key"), duree: duree, niveau: val("info-niveau"),
    tags: parseTagList(val("info-tags")).join(", "),
  };
  Object.assign(song, fields, { capo: fields.capo === "" ? null : fields.capo });
  closeInfoModal();
  refreshAll();
  if (state.source === "gas") gasPushFields(song, fields);
  else { markDirty(); showToast("ok", "✓ Enregistré (pense à télécharger le .xlsx)"); }
}


// ===== SUPPRESSION D'UNE CHANSON (liste de gauche) =====
async function deleteSongById(id) {
  const song = state.songs.find(s => s.id === id);
  if (!song) return;
  const inSets = Object.values(state.setlists || {}).filter(sl => (sl.songIds || []).includes(id)).length;
  const msg = `Supprimer « ${song.title} » ?\n\n`
    + (state.source === "gas" ? "Elle sera aussi supprimée du Google Sheet. " : "")
    + (inSets ? `Elle sera retirée de ${inSets} setlist${inSets > 1 ? "s" : ""}. ` : "")
    + "Cette action est définitive.";
  if (!confirm(msg)) return;
  // 1) d'abord le Sheet : si ça échoue, on ne touche à rien en local
  if (state.source === "gas" && state.gasUrl) {
    state.gasSyncing = true; renderDirtyBar();
    try {
      await gasPost({ action: "delete_song", id: id });
    } catch (e) {
      state.gasSyncing = false; renderDirtyBar();
      if (!/introuvable/i.test(e.message)) {   // « ID introuvable » = déjà supprimée dans le Sheet : on continue
        showToast("err", "Suppression impossible : " + e.message);
        return;
      }
    }
    state.gasSyncing = false; renderDirtyBar();
  }
  // 2) puis en local : liste, setlists, chanteurs
  state.songs = state.songs.filter(s => s.id !== id);
  Object.values(state.setlists || {}).forEach(sl => { sl.songIds = (sl.songIds || []).filter(x => x !== id); });
  delete state.songSingers[id];
  if (state.currentSongId === id) state.currentSongId = null;
  saveToStorage();                       // sauvegarde aussi les setlists nettoyées (sync Sheet)
  if (state.source !== "gas") markDirty();
  refreshAll();
  showToast("ok", state.source === "gas" ? "✓ Chanson supprimée (Sheet à jour)" : "✓ Chanson supprimée (pense à télécharger le .xlsx)");
}


// ===== ÉCOUTER (Spotify / autre lien / recherche) =====
function updateListenButtons(song) {
  const show = (id, on) => { const b = document.getElementById(id); if (b) b.style.display = on ? "" : "none"; };
  show("listen-spotify", !!song.spotify);
  show("listen-other", !!song.lien_ecoute);
  show("listen-search", !song.spotify && !song.lien_ecoute);   // aucun lien saisi : recherche dans Spotify
}
function openListen(kind) {
  const song = state.songs.find(s => s.id === state.currentSongId);
  if (!song) return;
  let url = "";
  if (kind === "spotify") url = song.spotify;
  else if (kind === "other") url = song.lien_ecoute;
  else url = "https://open.spotify.com/search/" + encodeURIComponent(((song.artiste || "") + " " + song.title).trim());
  if (!/^https?:\/\//i.test(url || "")) { showToast("err", "Lien invalide : corrige-le dans « ⓘ Infos »."); return; }
  window.open(url, "_blank", "noopener");
}

// ===== CHANTEURS =====
function getSongSingers(songId) {
  if (!state.songSingers[songId]) state.songSingers[songId] = { count: 1, names: [], assignments: {} };
  return state.songSingers[songId];
}

function openSingersModal() {
  if (!state.currentSongId) { alert("Selectionnez d'abord une chanson."); return; }
  refreshSingersModal();
  document.getElementById("singers-modal").classList.add("show");
}
function closeSingersModal() { document.getElementById("singers-modal").classList.remove("show"); }

function refreshSingersModal() {
  const cfg = getSongSingers(state.currentSongId);
  document.querySelectorAll("#singer-count-row .btn").forEach(b => {
    b.classList.toggle("active", Number(b.dataset.n) === cfg.count);
  });
  const list = document.getElementById("singer-names-list");
  if (cfg.count <= 1) {
    list.innerHTML = `<p style="color:var(--text-dim);font-size:12px;margin:8px 0 0">Avec un seul chanteur, pas besoin de couleurs. Choisissez 2 chanteurs ou plus pour activer l'attribution.</p>`;
  } else {
    let html = "<label>Noms (optionnels)</label>";
    for (let i = 1; i <= cfg.count; i++) {
      const name = cfg.names[i - 1] || "";
      html += `<div class="singer-name-row s-${i}"><span class="swatch"></span><input type="text" value="${escapeHtml(name)}" placeholder="Chanteur ${i}" onchange="setSingerName(${i}, this.value)"></div>`;
    }
    list.innerHTML = html;
  }
  refreshSingersBtnLabel();
}

function setSingerCount(n) {
  const cfg = getSongSingers(state.currentSongId);
  cfg.count = Math.max(1, Math.min(4, n));
  Object.keys(cfg.assignments).forEach(k => {
    if (cfg.assignments[k] > cfg.count) cfg.assignments[k] = 0;
  });
  cfg.names = (cfg.names || []).slice(0, cfg.count);
  saveToStorage();
  refreshSingersModal();
  renderSong();
}

function setSingerName(idx, name) {
  const cfg = getSongSingers(state.currentSongId);
  cfg.names = cfg.names || [];
  while (cfg.names.length < idx) cfg.names.push("");
  cfg.names[idx - 1] = name.trim();
  saveToStorage();
  refreshSingersBtnLabel();
}

function enterSingerAssignMode() {
  const cfg = getSongSingers(state.currentSongId);
  if (cfg.count < 2) {
    alert("Choisissez au moins 2 chanteurs avant d'activer l'attribution.");
    return;
  }
  closeSingersModal();
  state.singerAssignMode = true;
  document.getElementById("app").classList.add("singer-assign-mode");
  renderSingerPalette();
  renderSong();
}

function exitSingerAssignMode() {
  state.singerAssignMode = false;
  document.getElementById("app").classList.remove("singer-assign-mode");
}

function renderSingerPalette() {
  const cfg = getSongSingers(state.currentSongId);
  const pal = document.getElementById("singer-palette");
  if (!pal) return;
  let html = `<span class="swatch sw-0">tous</span>`;
  for (let i = 1; i <= cfg.count; i++) {
    const name = (cfg.names && cfg.names[i - 1]) || ("Chanteur " + i);
    html += `<span class="swatch sw-${i}">${escapeHtml(name)}</span>`;
  }
  pal.innerHTML = html;
}

function cycleLineSinger(lineEl) {
  const songId = state.currentSongId;
  if (!songId) return;
  const cfg = getSongSingers(songId);
  const lineKey = lineEl.dataset.line;
  if (!lineKey) return;
  const cur = cfg.assignments[lineKey] || 0;
  const next = (cur + 1) % (cfg.count + 1);
  if (next === 0) delete cfg.assignments[lineKey];
  else cfg.assignments[lineKey] = next;
  saveToStorage();
  for (let i = 0; i <= 4; i++) lineEl.classList.remove("line-singer-" + i);
  lineEl.classList.add("line-singer-" + next);
}

function resetSongAssignments() {
  if (!confirm("Reinitialiser toutes les attributions de chanteurs pour cette chanson ?")) return;
  const cfg = getSongSingers(state.currentSongId);
  cfg.assignments = {};
  saveToStorage();
  renderSong();
}

function refreshSingersBtnLabel() {
  const btn = document.getElementById("singers-btn");
  const cnt = document.getElementById("singers-btn-count");
  if (!btn || !cnt) return;
  if (!state.currentSongId) { cnt.textContent = ""; return; }
  const cfg = getSongSingers(state.currentSongId);
  cnt.textContent = cfg.count > 1 ? cfg.count : "";
}

// ===== INIT =====
loadFromStorage();
refreshAll();
setConnectBtn(state.gasUrl ? "idle" : null);
document.getElementById("search").addEventListener("input", applyFilters);
document.getElementById("filter-key").addEventListener("change", applyFilters);
document.getElementById("filter-capo").addEventListener("change", applyFilters);
document.getElementById("filter-style").addEventListener("change", applyFilters);
document.getElementById("filter-tag").addEventListener("change", applyFilters);

// Auto-load au démarrage : Apps Script (préféré) sinon Drive
if (state.gasUrl) {
  (async () => {
    setConnectBtn("loading");
    const tryLoad = async () => {
      const [dataList, dataSetlists, dataSingers] = await Promise.all([
        gasGet("list"),
        gasGet("setlists").catch(() => null),
        gasGet("singers").catch(() => null),
      ]);
      return { dataList, dataSetlists, dataSingers };
    };
    let attempt = 0;
    while (attempt < 2) {
      try {
        const { dataList, dataSetlists, dataSingers } = await tryLoad();
        const songs = (dataList.songs || []).map(normalizeSongFromGAS);
        if (songs.length) {
          state.songs = songs;
          state.source = "gas";
          state.lastLoaded = new Date();
          if (dataSetlists && dataSetlists.setlists) {
            const remote = dataSetlists.setlists;
            if (Object.keys(remote).length) { state.setlists = remote; state.currentSetlistId = dataSetlists.currentId || Object.keys(remote)[0]; }
          }
          if (dataSingers && dataSingers.singers) state.songSingers = dataSingers.singers;
          setConnectBtn("ok");
          refreshAll();
        }
        break;
      } catch (e) {
        attempt++;
        if (attempt < 2) {
          await new Promise(r => setTimeout(r, 4000));
        } else {
          console.warn("Auto-load GAS échoué après 2 tentatives :", e.message);
          setConnectBtn("err");
        }
      }
    }
  })();
} else if (state.driveUrl) {
  (async () => {
    try {
      const ab = await fetchXlsxFromDrive(state.driveUrl);
      const songs = parseXlsxArrayBuffer(ab);
      if (songs.length) {
        state.songs = songs;
        state.source = "drive";
        state.lastLoaded = new Date();
        refreshAll();
      }
    } catch (e) { /* silent fail, user can retry from settings */ }
  })();
}

// ===== DRAG & DROP for .xlsx =====
(function setupDragDrop() {
  let counter = 0;
  const overlay = document.createElement("div");
  overlay.id = "drop-overlay";
  overlay.innerHTML = "📥 Déposez le .xlsx pour le charger";
  document.body.appendChild(overlay);
  document.addEventListener("dragenter", e => {
    if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files")) {
      counter++; overlay.classList.add("show");
    }
  });
  document.addEventListener("dragover", e => { if (counter > 0) e.preventDefault(); });
  document.addEventListener("dragleave", () => { counter = Math.max(0, counter - 1); if (counter === 0) overlay.classList.remove("show"); });
  document.addEventListener("drop", e => {
    counter = 0; overlay.classList.remove("show");
    if (!e.dataTransfer || !e.dataTransfer.files || !e.dataTransfer.files.length) return;
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!/\.(xlsx|xls)$/i.test(file.name)) { alert("Format non supporté. Glissez un fichier .xlsx."); return; }
    // Open settings panel so the status message is visible, then load
    document.getElementById("settings-modal").classList.add("show");
    loadFromFile(file);
  });
})();

document.addEventListener("keydown", e => {
  // Ne pas intercepter les touches quand on saisit du texte
  const tag = (e.target && e.target.tagName) || "";
  const isTyping = tag === "INPUT" || tag === "TEXTAREA" || (e.target && e.target.isContentEditable);

  if (e.key === "Escape") {
    const app = document.getElementById("app");
    if (document.getElementById("settings-modal").classList.contains("show")) closeSettings();
    else if (document.getElementById("add-modal").classList.contains("show")) closeAddSong();
    else if (document.getElementById("info-modal").classList.contains("show")) closeInfoModal();
    else if (document.getElementById("singers-modal").classList.contains("show")) closeSingersModal();
    else if (app.classList.contains("singer-assign-mode")) exitSingerAssignMode();
    else if (app.classList.contains("sidebar-open") || app.classList.contains("setlist-open")) app.classList.remove("sidebar-open", "setlist-open");
    else if (app.classList.contains("setlist-mode")) exitSetlistPlay();
    else { app.classList.remove("stage"); renderSong(); }
  }

  // Téléprompteur : Espace = play/pause, +/- = vitesse, R = reset (uniquement si une chanson est affichée et qu'on ne saisit pas)
  if (state.currentSongId && !isTyping) {
    if (e.code === "Space") { e.preventDefault(); toggleAutoScroll(); return; }
    if (e.key === "+" || e.key === "=") { e.preventDefault(); changeScrollSpeed(5); return; }
    if (e.key === "-" || e.key === "_") { e.preventDefault(); changeScrollSpeed(-5); return; }
    if (e.key === "r" || e.key === "R") { e.preventDefault(); resetScroll(); return; }
    if (e.key === "f" || e.key === "F") { e.preventDefault(); toggleFullscreen(); return; }
  }

  // Pédale / clavier : flèches et Page haut/bas = défilement d'un écran (comme une tourne-page).
  // Maj + flèche gauche/droite = chanson précédente/suivante de la setlist (mode scène).
  const modalOpen = document.querySelector(".modal-bg.show");
  if (state.currentSongId && !isTyping && !modalOpen) {
    const app = document.getElementById("app");
    const sl = state.setlists[state.currentSetlistId];
    if (e.shiftKey && (e.key === "ArrowRight" || e.key === "ArrowLeft") && app.classList.contains("stage") && sl && sl.songIds.length) {
      const idx = sl.songIds.indexOf(state.currentSongId);
      const ni = e.key === "ArrowRight" ? idx + 1 : idx - 1;
      if (ni >= 0 && ni < sl.songIds.length) { e.preventDefault(); selectSong(sl.songIds[ni]); }
      return;
    }
    const down = ["ArrowDown", "PageDown", "ArrowRight"].includes(e.key);
    const up = ["ArrowUp", "PageUp", "ArrowLeft"].includes(e.key);
    if ((down || up) && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const el = document.getElementById("song-body");
      if (el) {
        e.preventDefault();
        const step = Math.max(80, Math.round(el.clientHeight * 0.85));
        el.scrollBy({ top: down ? step : -step, behavior: "smooth" });
      }
    }
  }
});

// Mode attribution chanteur : tap sur une ligne cycle le chanteur
document.addEventListener("click", e => {
  const app = document.getElementById("app");
  if (app.classList.contains("singer-assign-mode")) {
    const lineEl = e.target.closest(".line");
    if (lineEl && lineEl.dataset.line) {
      e.preventDefault();
      e.stopPropagation();
      cycleLineSinger(lineEl);
      return;
    }
    return; // Bloque tout autre comportement pendant le mode attribution
  }
  // Cliquer sur le voile derriere la sidebar mobile → la fermer
  if (window.innerWidth > 1100) return;
  if (app.classList.contains("sidebar-open")) {
    const left = document.querySelector(".left");
    const toggle = document.getElementById("sidebar-toggle");
    if (left && !left.contains(e.target) && toggle && !toggle.contains(e.target)) {
      app.classList.remove("sidebar-open");
    }
  }
  if (app.classList.contains("setlist-open")) {
    const right = document.querySelector(".right");
    if (right && !right.contains(e.target) && !e.target.closest(".setlist-toggle")) {
      app.classList.remove("setlist-open");
    }
  }
});
</script>
</body>
</html>
"""

OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
final_html = (HTML
              .replace("__SONGS_JSON__", songs_json)
              .replace("__BUILD_VERSION__", BUILD_VERSION))
OUT_HTML.write_text(final_html, encoding="utf-8")
print(f"Wrote {OUT_HTML}  ({OUT_HTML.stat().st_size:,} bytes)  — build {BUILD_VERSION}")
