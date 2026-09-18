# SORA-Präsentationen v0.3.0

Komplette Toolchain für SORA-Wochenberichte und Referate (Klasse 11-12).

## Was ist neu in v0.3.0

Drei harte Lessons aus dem Politik-Wochenbericht KW06/2026 sind eingearbeitet:

**1. Animations-XML neu geschrieben (`inject_clicks.py`)** — Die alte v2 produzierte ungültiges Timing-XML, das PowerPoint beim Öffnen "reparierte" und dabei alle Klick-Funktionen entfernte. v3 folgt jetzt strikt dem PowerPoint-Standard:
- Jede Klick-Gruppe in eigenem `<p:par>`-Container mit `delay="indefinite"`
- `prevCondLst` + `nextCondLst` im `<p:seq>` (PFLICHT)
- Erstes Shape pro Gruppe = `clickEffect`, Rest = `withEffect`
- Global eindeutige, aufsteigende Effect-IDs

**2. Selektive Animation per `--skip`** — Nicht jede Folie braucht Animation. Standardmäßig animiert: FLOW, Cause-Effect, Mindmap, sequenzielle Icon-Karten. Standardmäßig statisch: Hero, Data-Hero, Quote, Split, Summary. CLI-Skip: `python3 inject_clicks.py in.pptx out.pptx --skip 4,5,9,10`.

**3. Echte Umlaute via `fix_umlauts.py`** — Statt ASCII-Transliteration im build.js (Sondervermoegen, Buerger, schuetzte) wird ein Wörterbuch-Replace nach dem Build gefahren. Behandelt Folien-Text UND Speaker-Notes.

## Was drin ist

**sora-pptx-design** — Design-System + 3 Helper-Scripts:
- `inject_clicks.py` (NEU) — PowerPoint-konformes Timing-XML
- `fix_umlauts.py` (NEU) — ae/oe/ue → ä/ö/ü Wörterbuch
- Zusammen ergeben die Pipeline: build.js → fix_umlauts → inject_clicks → fertig

**cowork-visuals** — Bildbeschaffung. Stock-Foto via Chrome-Extension (Unsplash/Pexels/Wikimedia), Native-Shape Napkin-Diagramme (Flow, Mindmap, Cause-Effect), keine KI-Bilder.

## Installation

Als Marketplace direkt aus diesem Repo:

```
/plugin marketplace add Kaytm93/sora-praesentationen
/plugin install sora-praesentationen@kay-sora
```

Alternativ: wenn Claude dich fragt, ob du dieses Plugin installieren willst, drück „Annehmen". Die Skills werden dann automatisch eingehängt und beim nächsten Chat verfügbar.

## Trigger

Skills laden automatisch bei: *SORA, Wochenbericht, Referat, pptx erstellen, Klimapolitik, Politik-Wochenbericht, Bild, Hero-Image, Flowchart, Mindmap*.

## Typischer Ablauf

1. „Erstelle eine SORA-Präsentation zu [Thema]" → `sora-pptx-design` lädt automatisch
2. Build-Script läuft → produziert pptx mit ASCII-Platzhaltern und animierbaren Shape-Namen (`g1_`, `g2_`, ...)
3. `fix_umlauts.py` ersetzt ASCII durch echte Umlaute
4. `inject_clicks.py --skip <statische_folien>` injiziert nur dort Klick-Reveal, wo es sinnvoll ist
5. PPTX öffnet in PowerPoint OHNE Reparatur-Dialog

## Version

0.3.0 · Autor: Kay
