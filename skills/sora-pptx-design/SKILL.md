---
name: sora-pptx-design
description: Premium Design-System für SORA-Präsentationen (Wochenberichte & Referate, Klasse 11-12). Aktiviere diesen Skill IMMER zusammen mit dem pptx-Skill, wenn du eine SORA-Präsentation erstellst. Erzwingt Text-Minimalismus (Keywords statt Sätze), Oberstufen-Sprache, echte Umlaute, PowerPoint-konforme Click-to-Reveal-Animationen auf ausgewählten Folien, thematische Hero-Images, variierende Layouts (Icon-Karten, Flow, Mindmap, Split, Data-Hero, Quote) und Canva-Kompatibilität. Triggers: SORA, Wochenbericht, Präsentation, Politik, Wirtschaft, Referat, pptx erstellen.
---

# SORA PPTX-Design v3

Premium Design-System für Schul-Oberstufen-Präsentationen. Output: pptx, die in PowerPoint ohne Reparatur-Dialog öffnet, selektiv animiert ist, und auf Canva importierbar bleibt.

## Goldene Regel

**Folien sind visuelle Anker, keine Vortragsmanuskripte.** Alles Gesprochene gehört in die Speaker-Notes. Auf der Folie stehen nur Keywords, Zahlen, Begriffe. Der Zuhörer soll DICH anschauen, nicht die Folie lesen.

## Pipeline

```
build.js (pptxgenjs)  →  fix_umlauts.py  →  inject_clicks.py  →  fertig.pptx
```

Drei Schritte, drei Gründe:
1. **build.js** erzeugt die Folien mit ASCII-Text (`Sondervermoegen`, `schuetzte`) und `objectName`-Tags (`g1_card`, `g2_card`, ...)
2. **fix_umlauts.py** ersetzt ASCII durch echte Unicode-Umlaute (ä, ö, ü, ß) in Folien UND Speaker-Notes
3. **inject_clicks.py** injiziert PowerPoint-konformes Timing-XML nur auf den Folien, die es brauchen

## Text-Regeln

- **Pro Karte max. 3 Keywords.** Nicht „Die deutsche Wirtschaft wächst langsam" sondern „BIP +0,3 %".
- **12.-Klasse-Sprache.** Fachbegriffe statt Umgangssprache: UNFCCC statt „Klimakonferenz", Fit-for-55 statt „EU-Klimapaket", NATO-Ziel 2 % statt „Rüstungsausgaben".
- **Speaker-Notes tragen den gesprochenen Satz.** `addNotes()` pro Folie, deutsch, Du-Form nicht nötig (Kay liest sich das vorher einmal durch).
- **Eine INTERAKTION-Zeile pro Speaker-Note.** „Daumen hoch/runter für...?", „Schätzt mal: ...?", „Diskutiert kurz mit Nachbarn: ...?"
- **Alle Umlaute und ß sind Pflicht.** Keine ae/oe/ue-Transliterationen auf der finalen Folie. Transliteration im build.js ist OK, solange `fix_umlauts.py` danach läuft.

## Layout-Katalog (wähle variierend)

| Layout | Wann | Animation? |
|---|---|---|
| **Hero-Image** (Titel-Folie) | Folie 1, thematisches Foto mit Darkening-Overlay | NEIN |
| **FLOW** (Zeitstrahl mit Pfeilen) | historische Entwicklung, Schritt-für-Schritt | JA, sequenziell |
| **Icon-Karten 2x2 oder 1x4** | Themen-Übersicht, Versprechen, Kategorien | JA, pro Karte |
| **Status-Ampel (grün/gelb/rot)** | Bilanz, Realitäts-Check, Bewertung | NEIN (Zusammenschau) |
| **Data-Hero** (Donut/Balken + zentrale Zahl) | Umfrage-Ergebnis, einzelne harte Zahl | NEIN |
| **Cause-Effect Domino** | Warum-Ketten, Ursache-Folge | JA, sequenziell |
| **SPLIT** (zwei Spalten pro/contra) | Entscheidung, Abwägung, Dialektik | NEIN (alles auf Blick) |
| **QUOTE** (dunkel, großes Zitat) | Abschluss-Wirkung, Experten-Meinung | NEIN (Zitat wirkt sofort) |
| **IMPACT** (dunkler Hintergrund, ein Satz) | Kernaussage, Abschluss | NEIN |
| **Mindmap / Cluster** | Begriffs-Netz, Abhängigkeiten | JA, Äste nacheinander |

**Regel: mindestens 6 verschiedene Layouts in 10 Folien.** Monotonie ist der Tod von Präsentationen.

## Animations-Granularität

**Der Vortragende darf nie von seinen eigenen Folien überholt werden.** Jede Klick-Gruppe wartet auf User-Klick (`delay="indefinite"`). Nie Auto-Advance.

**Wann mehrere Klicks pro Folie:**
- FLOW: ein Klick pro Station, Pfeil gehört zur nachfolgenden Station
- Cause-Effect: ein Klick pro Schritt
- Icon-Karten mit eigenständigen Themen: ein Klick pro Karte
- Mindmap: ein Klick pro Ast

**Wann ALLES zusammen einblenden (= kein `g<N>_`-Prefix):**
- Hero/Titel-Folie: statisch, kein Klick nötig
- Data-Hero: Donut + zentrale Zahl wirken gemeinsam
- Quote: das Zitat muss sofort da sein, um zu wirken
- Split: Pro und Contra müssen parallel sichtbar sein für Vergleich
- Summary / Realitäts-Check: Überblick, keine Dramaturgie

**Selektiver Skip beim Injizieren:**
```bash
python3 inject_clicks.py in.pptx out.pptx --skip 4,5,9,10
```

## Farbpaletten

**Palette A — Politik** (ernst, staatstragend):
- `#1A2E5A` (Navy) — Titel, Headlines
- `#C8102E` (Accent-Red) — Akzentlinien, Badges, Warnungen
- `#F5F4F0` (Cream) — Hintergrund
- `#E8E6DF` (Card-Soft) — Card-Background
- `#2E7D32` (Ok-Green) — Status OK
- `#F9A825` (Warn-Amber) — Status WIP

**Palette B — Wirtschaft** (sachlich, zahlengetrieben):
- `#0A2540` (Deep-Blue)
- `#00A693` (Teal)
- `#F4E7D3` (Warm-Cream)
- `#E85D04` (Accent-Orange) für negative Werte

**Palette C — Klima/Umwelt**:
- `#1B4332` (Forest)
- `#2D6A4F` (Mid-Green)
- `#E9F5EC` (Light-Green)
- `#D62828` (Alert-Red)

## Typografie

```js
const F = {
  TITLE: "Georgia",    // Serif für Titel (klassisch, lesbar, in PowerPoint + Canva vorhanden)
  BODY:  "Calibri"     // Sans-Serif für Fließtext (Standard-Font, rendert überall)
};
```

**Titel-Fontsize-Ladder** (abhängig von Länge):
- ≤ 20 Zeichen → 36pt
- 21-30 Zeichen → 30pt
- 31-40 Zeichen → 26pt
- > 40 Zeichen → 22pt (und Titel umformulieren!)

**Immer `charSpacing: 1`** auf Titeln (in pptxgenjs-Einheit, ≈ 0.1em). Ohne fühlt sich Georgia-Spacing zu eng an.

## Hero-Image-Pattern (Folie 1)

```js
// Hintergrundbild vollflächig
s.addImage({ path: "assets/hero.jpg", x: 0, y: 0, w: 10, h: 5.625, sizing: { type: "cover", w: 10, h: 5.625 } });
// Darkening-Overlay (Navy 70% transparent)
s.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 5.625,
  fill: { color: "1A2E5A", transparency: 30 },
  line: { color: "1A2E5A", transparency: 100 }
});
// Großer Titel, Serif, white
s.addText("THEMA.", {
  x: 0.5, y: 1.8, w: 9, h: 1.8,
  fontFace: "Georgia", fontSize: 72, color: "FFFFFF",
  bold: true, align: "left", charSpacing: 1
});
// Subline
s.addText("Was bedeutet das eigentlich?", {
  x: 0.5, y: 3.5, w: 9, h: 0.7,
  fontFace: "Georgia", fontSize: 28, color: "F5F4F0",
  italic: true, align: "left"
});
// Akzentlinie + Datumszeile + Eyebrow siehe titleBlock()
```

## Titel-Block mit Akzentlinie

```js
function titleBlock(s, title, opts) {
  opts = opts || {};
  var size = opts.size || (title.length <= 20 ? 36 : title.length <= 30 ? 30 : title.length <= 40 ? 26 : 22);
  // Eyebrow (oberhalb)
  if (opts.eyebrow) {
    s.addText(opts.eyebrow, {
      x: 0.5, y: 0.25, w: 9, h: 0.3,
      fontFace: "Calibri", fontSize: 11, color: "8A8A8A",
      bold: true, charSpacing: 3, align: "left"
    });
  }
  // Titel
  s.addText(title, {
    x: 0.5, y: 0.55, w: 9, h: 0.9,
    fontFace: "Georgia", fontSize: size, color: "1A2E5A",
    bold: true, align: "left", charSpacing: 1
  });
  // Akzentlinie (dynamisch unter Titel, nicht auf fester y)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.55 + (size / 72) + 0.05, w: 1.2, h: 0.05,
    fill: { color: "C8102E" }, line: { color: "C8102E" }
  });
}
```

## Blob-Motif (Hintergrunds-Deko)

Dezente Navy-Kreise in oberer rechter + unterer linker Ecke. NICHT auf Hero-Folien. Auf allen Content-Folien.

```js
function blobMotif(s, variant) {
  var col = variant === "dark" ? "0F1C3A" : "1A2E5A";
  s.addShape(pres.shapes.OVAL, {
    x: 8.5, y: -1, w: 3, h: 3,
    fill: { color: col, transparency: variant === "dark" ? 0 : 60 },
    line: { color: col, transparency: 100 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: -0.8, y: 4.3, w: 2.5, h: 2.5,
    fill: { color: col, transparency: 70 },
    line: { color: col, transparency: 100 }
  });
}
```

## Wort-Wrap-Schutz

**Problem:** Lange deutsche Wörter (Unterzeichnet, Ordnungskollaps, Sondervermögen) brechen in schmalen Karten um und zerreißen das Layout.

**Regeln:**
1. Wenn ein Wort > 12 Zeichen und Karten-Breite < 3.0 Zoll → entweder Wort shortnen (Signiert statt Unterzeichnet) ODER Font reduzieren
2. FLOW-Stationen sollten 2-Wörter-Keywords haben, nicht 1-Wort-Komposita
3. Bei Icon-Karten 1x3 (nur 3 Karten in einer Reihe): Card-Width mindestens 2.8 Zoll, Font max 22pt

**Quick-Check:** Nach jedem Build die PDF-Preview rendern lassen (soffice → pdftoppm) und bei jeder Folie scannen ob irgendwo ein Wort zweigeteilt ist.

## Footer / Quellenangabe

Unter jeder Content-Folie (nicht Hero, nicht Impact):
```js
function footerSource(s, txt) {
  s.addText(txt, {
    x: 0.5, y: 5.3, w: 9, h: 0.25,
    fontFace: "Calibri", fontSize: 9, color: "9B9B9B",
    italic: true, align: "left", margin: 0
  });
}
```

Inhalt: `"Quelle: [Institution] [Jahr]  ·  KW [NR] • Folie [N]"`

## Canva-Kompatibilität

- **Nur native Shapes** (RECTANGLE, ROUNDED_RECTANGLE, OVAL, RIGHT_ARROW, DOWN_ARROW, TRIANGLE). Keine Custom-Paths.
- **Standard-Fonts** (Georgia, Calibri, Arial). Keine exotischen Web-Fonts.
- **Keine Gruppen** (pptxgenjs hat `addGroup` nicht — gut so).
- **Keine Master-Layouts.** Alles auf der einzelnen Folie, damit Canva sie einzeln editierbar hält.
- **Bilder als JPG/PNG eingebettet**, nicht verlinkt.

## QA-Checkliste vor Abgabe

- [ ] Folie 1 Hero-Image nicht verzerrt, Text lesbar (Weiß auf Dark-Overlay)
- [ ] Mindestens 6 verschiedene Layouts in 10 Folien
- [ ] Keine Wortumbrüche in schmalen Karten
- [ ] Alle Umlaute gerendert (keine „ae"/„oe"/„ue" sichtbar)
- [ ] Titel läuft nicht in Blob hinein (Titel-Fontsize-Ladder beachten)
- [ ] Footer-Source-Text überlappt nicht mit Content
- [ ] Animationen NUR auf Folien, die es brauchen (FLOW, Cause-Effect, sequenzielle Icon-Karten)
- [ ] PowerPoint öffnet ohne „muss repariert werden"-Dialog
- [ ] Speaker-Notes pro Folie vorhanden, mit INTERAKTION-Vorschlag
- [ ] Quellen-Angabe pro Folie (Institution + Jahr)

## Pipeline-Commands

```bash
# 1. Build
cd project_dir
node build.js   # erzeugt z.B. sicherheit.pptx mit ASCII

# 2. Umlaute
python3 /pfad/zu/fix_umlauts.py sicherheit.pptx sicherheit_umlaut.pptx

# 3. Selektive Animation
python3 /pfad/zu/inject_clicks.py sicherheit_umlaut.pptx sicherheit_final.pptx --skip 4,5,9,10

# 4. Visual QA
soffice --headless --convert-to pdf sicherheit_final.pptx
pdftoppm -jpeg -r 100 sicherheit_final.pdf qa/slide
# → qa/slide-01.jpg ... qa/slide-10.jpg zum Durchschauen
```

## Scripts im Skill

- `inject_clicks.py` — Click-to-Reveal-Animations-Injector, PowerPoint-konform
- `fix_umlauts.py` — ASCII → Umlaute mit deutschem Wörterbuch (Folien + Speaker-Notes)

Beide sind idempotent (mehrfach ausführbar ohne Schaden) und funktionieren auf jeder pptxgenjs-Ausgabe, solange Shapes den `g<N>_name`-Objektnamen tragen.
