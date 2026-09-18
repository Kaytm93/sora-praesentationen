---
name: cowork-visuals
description: Beschafft und generiert Bilder & Grafiken für Präsentationen, Dokumente und Reports in Cowork. Nutzt Chrome-Extension für Stock-Fotos (Unsplash/Pexels/Wikimedia), baut Napkin-Style-Diagramme (Flow, Mindmap, Prozess) direkt als PPTX-Shapes und kann auch napkin.ai live ansteuern. Triggert auf — Bild, Foto, Grafik, Diagramm, Mindmap, Flowchart, Illustration, Hintergrundbild, Hero-Image, Napkin, Unsplash, Pexels, Wikimedia, visualisieren.
---

# cowork-visuals

Findet, lädt und generiert Bilder & Grafiken für Präsentationen / Dokumente.
Zusammen mit `pptx`, `sora-pptx-design` oder `docx` verwenden.

---

## Entscheidungs-Matrix: Welche Methode wann?

| Bedarf | Methode |
|--------|---------|
| Titel-Hero-Foto (Windräder, Protest, Stadt) | **1. Stock-Foto via Chrome** |
| Historisches Foto, Politiker, Ereignis | **1. Stock-Foto via Chrome (Wikimedia bevorzugt)** |
| Prozess, Flow, Ursache-Wirkung-Diagramm | **2. Napkin-Style nativ in PPTX** |
| Mindmap, Begriffsnetz | **2. Napkin-Style nativ in PPTX** |
| "Premium-polishes" Diagramm | **3. napkin.ai live via Chrome** |
| Symbol / Icon (klein, funktional) | **2. Nativ als Unicode / Shape** (kein Skill nötig) |

**Faustregel:** Erst Methode 2 versuchen (schnell, konsistent, Canva-kompatibel). Nur wenn das Ergebnis nicht reicht → Methode 3. Methode 1 immer für Hero-Fotos.

---

## Methode 1 — Stock-Foto via Chrome Extension

### Wann
Titelfolien-Hintergrund, Themen-Bebilderung (Klima → Windräder; Krieg → Trümmer; Wirtschaft → Börse).

### Quellen (in dieser Reihenfolge versuchen)
1. **Unsplash** (`unsplash.com/s/photos/<query>`) — Free-to-use-Lizenz, beste Ästhetik
2. **Pexels** (`pexels.com/search/<query>/`) — Ähnlich, oft mehr Auswahl
3. **Wikimedia Commons** (`commons.wikimedia.org/w/index.php?search=<query>`) — Für historische Fotos, Politiker, spezifische Ereignisse

### Ablauf

```
1. request_access für Chrome (Tier = read, Extension nötig für Klicks)
2. mcp__Claude_in_Chrome__navigate → Unsplash mit Such-Query
3. mcp__Claude_in_Chrome__read_page → Thumbnails & Alt-Texte lesen
4. Besten Treffer wählen (hochaufgelöst, thematisch passend, gute Komposition)
5. Auf Download-Button klicken (Unsplash gibt direkte JPG-URL)
6. Via Bash: curl -sL "<URL>" -o assets/hero.jpg
7. In PPTX einbauen mit sizing: { type: "cover", w, h }
8. Dunkler Overlay darüber (für Text-Kontrast) — siehe Code unten
```

### Such-Queries — Tipps für Schul-Themen

| Thema | Effektiver Query |
|-------|------------------|
| Klimapolitik | "wind turbine dusk" / "climate protest" / "melting ice" |
| Haushalt / Finanzen | "euro coins" / "german parliament" / "budget chart" |
| Verteidigung | "soldiers silhouette" / "nato flag" / "military parade" |
| Demokratie / Wahlen | "ballot box" / "bundestag hall" / "hand voting" |
| Außenpolitik | "diplomatic handshake" / "un general assembly" |
| Migration | "border fence dusk" / "refugee tent" |
| Digitalisierung | "server room blue" / "fiber optic" |

Einfache englische 2-3-Wort-Queries + Stimmungs-Keyword (`dusk`, `blue`, `dramatic`) liefern bessere Treffer als deutsche Langformen.

### Lizenz-Check
- Unsplash + Pexels: Nutzungsrecht frei (Attribution nicht zwingend, aber höflich im Notes-Feld)
- Wikimedia: Lizenz prüfen (meist CC-BY-SA — Attribution auf Quellenfolie)

### Hero-Image in PPTX einbauen (Standard-Pattern)

```javascript
// Volle Slide-Größe als Hintergrund
slide.addImage({
  path: "assets/hero.jpg",
  x: 0, y: 0, w: 10, h: 5.625,
  sizing: { type: "cover", w: 10, h: 5.625 }
});

// Dunkler Overlay für Text-Kontrast (80% Deckkraft unten, 30% oben für Gradient-Gefühl)
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 5.625,
  fill: { color: "1A2E5A", transparency: 40 },   // tint
  line: { color: "1A2E5A", transparency: 100 }
});
// Optional: zweiter Overlay unten für stärkeres Abdunkeln (falls Text weiß unten steht)
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 2.5, w: 10, h: 3.125,
  fill: { color: "000000", transparency: 55 },
  line: { color: "000000", transparency: 100 }
});
```

**Regel:** Transparency-Wert bei pptxgenjs ist invertiert (0 = voll sichtbar, 100 = unsichtbar). Für "blass/verdunkelt" → Overlay-Transparency zwischen 30-55.

---

## Methode 2 — Napkin-Style Diagramme nativ in PPTX

Claude baut Flow-Diagramme, Mindmaps und Prozess-Darstellungen direkt aus PPTX-Shapes. Ergebnis ist 100% Canva-kompatibel (keine embedded Bilder — alles als editierbare Shapes).

### Pattern A — Flow mit 3-4 Stationen + Pfeilen

```javascript
/**
 * addFlowDiagram — Horizontaler Prozess-Flow
 * @param {Array<{label, sub}>} stations   — z.B. [{label:"UN-Treffen", sub:"1992"}, ...]
 * @param {Object} opts  — x, y, w, h, colorPrimary, colorAccent
 */
function addFlowDiagram(slide, stations, opts) {
  const { x, y, w, h, colorPrimary="1A2E5A", colorAccent="C8102E" } = opts;
  const n = stations.length;
  const stationW = (w - (n-1)*0.35) / n;  // 0.35" Lücke für Pfeil

  stations.forEach((st, i) => {
    const cx = x + i * (stationW + 0.35);

    // Station-Box (rounded rect)
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: y, w: stationW, h: h,
      fill: { color: i === n-1 ? colorAccent : "FFFFFF" },
      line: { color: colorPrimary, width: 1.5 },
      rectRadius: 0.08,
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, opacity: 0.08 }
    });
    // Label
    slide.addText(st.label, {
      x: cx, y: y + 0.15, w: stationW, h: 0.5,
      fontFace: "Calibri", fontSize: 14, bold: true,
      color: i === n-1 ? "FFFFFF" : colorPrimary,
      align: "center", margin: 0
    });
    // Sublabel
    if (st.sub) slide.addText(st.sub, {
      x: cx, y: y + 0.65, w: stationW, h: 0.4,
      fontFace: "Calibri", fontSize: 11,
      color: i === n-1 ? "FFFFFF" : "666666",
      align: "center", italic: true, margin: 0
    });

    // Pfeil zur nächsten Station
    if (i < n-1) {
      const ax = cx + stationW;
      slide.addShape(pres.shapes.RIGHT_TRIANGLE, {  // Fallback: ein Pfeil-Driiek
        x: ax + 0.08, y: y + h/2 - 0.08, w: 0.2, h: 0.16,
        fill: { color: colorPrimary }, line: { color: colorPrimary },
        rotate: 90
      });
      // Linie zum Pfeil
      slide.addShape(pres.shapes.LINE, {
        x: ax, y: y + h/2, w: 0.12, h: 0,
        line: { color: colorPrimary, width: 2 }
      });
    }
  });
}
```

**Verwendung:**
```javascript
addFlowDiagram(slide, [
  { label: "1992", sub: "UN-Klimakonvention" },
  { label: "1997", sub: "Kyoto-Protokoll" },
  { label: "2015", sub: "Pariser Abkommen" },
  { label: "HEUTE", sub: "Umsetzung offen" }
], { x: 0.5, y: 2.2, w: 9, h: 1.4, colorPrimary: "1A2E5A", colorAccent: "C8102E" });
```

### Pattern B — Mindmap (1 Zentrum + 4-6 Äste)

```javascript
/**
 * addMindmap — Zentrales Keyword + strahlende Äste
 * @param {string} center     — Zentrales Wort
 * @param {Array<{label, angle}>} nodes  — angle in Grad, 0 = rechts
 */
function addMindmap(slide, center, nodes, opts) {
  const { cx=5, cy=3, radius=2.0, colorPrimary="1A2E5A", colorAccent="C8102E" } = opts;

  // Zentraler Kreis
  slide.addShape(pres.shapes.OVAL, {
    x: cx-0.9, y: cy-0.6, w: 1.8, h: 1.2,
    fill: { color: colorAccent },
    line: { color: colorAccent, width: 0 },
    shadow: { type: "outer", color: "000000", blur: 12, offset: 3, opacity: 0.15 }
  });
  slide.addText(center, {
    x: cx-0.9, y: cy-0.6, w: 1.8, h: 1.2,
    fontFace: "Georgia", fontSize: 16, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0
  });

  // Äste
  nodes.forEach(n => {
    const rad = (n.angle * Math.PI) / 180;
    const nx = cx + Math.cos(rad) * radius - 0.75;
    const ny = cy + Math.sin(rad) * radius - 0.3;

    // Verbindungslinie
    slide.addShape(pres.shapes.LINE, {
      x: cx + Math.cos(rad) * 0.95,
      y: cy + Math.sin(rad) * 0.65,
      w: Math.cos(rad) * (radius - 0.95),
      h: Math.sin(rad) * (radius - 0.65),
      line: { color: colorPrimary, width: 1.5, dashType: "dash" }
    });

    // Knoten-Box
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: nx, y: ny, w: 1.5, h: 0.6,
      fill: { color: "FFFFFF" },
      line: { color: colorPrimary, width: 1.25 },
      rectRadius: 0.08
    });
    slide.addText(n.label, {
      x: nx, y: ny, w: 1.5, h: 0.6,
      fontFace: "Calibri", fontSize: 12, bold: true,
      color: colorPrimary, align: "center", valign: "middle", margin: 0
    });
  });
}
```

**Verwendung:**
```javascript
addMindmap(slide, "Klima-\npolitik", [
  { label: "UN", angle: 0 },
  { label: "EU",  angle: 60 },
  { label: "NGOs", angle: 120 },
  { label: "Staaten", angle: 180 },
  { label: "Industrie", angle: 240 },
  { label: "Gesellschaft", angle: 300 }
], { cx: 5, cy: 3, radius: 2.0 });
```

### Pattern C — Ursache → Wirkung (2-Spalten mit Verbindungen)

```javascript
// Links: Ursachen
["CO2", "Methan", "Abholzung"].forEach((cause, i) => {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 1.5 + i*1.1, w: 2.5, h: 0.7,
    fill: { color: "FFFFFF" }, line: { color: "C8102E", width: 1.5 }, rectRadius: 0.08
  });
  slide.addText(cause, { x: 0.8, y: 1.5 + i*1.1, w: 2.5, h: 0.7,
    fontFace: "Calibri", fontSize: 14, bold: true,
    color: "C8102E", align: "center", valign: "middle", margin: 0 });

  // Pfeil zur rechten Seite
  slide.addShape(pres.shapes.LINE, {
    x: 3.3, y: 1.85 + i*1.1, w: 3.4, h: 0,
    line: { color: "888888", width: 1.25, endArrowType: "triangle" }
  });
});
// Rechts: Wirkungen
["Hitzewellen", "Meeresspiegel", "Ernteausfälle"].forEach((eff, i) => {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.7, y: 1.5 + i*1.1, w: 2.5, h: 0.7,
    fill: { color: "1A2E5A" }, line: { color: "1A2E5A" }, rectRadius: 0.08
  });
  slide.addText(eff, { x: 6.7, y: 1.5 + i*1.1, w: 2.5, h: 0.7,
    fontFace: "Calibri", fontSize: 14, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
});
```

---

## Methode 3 — napkin.ai live via Chrome Extension

### Wann
Für einmalig besonders polishes Diagramm (z.B. auf der wichtigsten Ergebnis-Folie). Nicht für jede Folie — kostet Zeit und hängt vom Napkin-Login ab.

### Voraussetzung
User muss in Chrome bei `napkin.ai` eingeloggt sein.

### Ablauf

```
1. request_access für Chrome
2. mcp__Claude_in_Chrome__navigate → https://app.napkin.ai/
3. mcp__Claude_in_Chrome__get_page_text → prüfen ob eingeloggt
   Falls nicht: dem User sagen, sich einzuloggen, dann warten
4. Neues Dokument → Text-Eingabe
5. Diagramm-Beschreibung eintippen (präzise, kurz, Fachsprache):
   "3-Schritt-Flow: Nationale Klimaziele → EU-Koordinierung → UN-Vertrag.
    Mit Pfeilen. Farbe: Navy und Rot."
6. Napkin generiert Grafiken → gewünschte auswählen
7. Export als PNG (transparent background)
8. Download via curl → assets/napkin_<name>.png
9. In PPTX einbauen mit slide.addImage
```

### Qualitäts-Checks für Napkin-Output
- Schriftgröße > 11pt (sonst bei Projektion unlesbar)
- Keine abgeschnittenen Beschriftungen
- Farben matchen die Palette der Präsentation (sonst: Napkin neu generieren mit Farb-Hinweis)

---

## Canva-Kompatibilität (Pflicht-Check)

Die Präsentation wird oft nach Canva hochgeladen und dort weiterbearbeitet. Dafür gilt:

### ✅ Canva-freundlich
- Native PPTX-Shapes (Rectangles, Ovals, Lines, Triangles) → bleiben editierbar
- Text in echten Textboxen (nicht als Pfad/SVG)
- Bilder als eingebettete JPG/PNG (nicht verlinkt)
- Einheitliche Fonts (Calibri, Georgia, Arial — nicht exotische)

### ⚠️ Probleme in Canva
- Komplexe SVG-Grafiken werden oft zu einem flachen Bild → nicht mehr editierbar
- Gruppierte Shape-Gruppen: Canva löst sie meist auf, Layout kann brechen
- Animationen gehen in Canva verloren (Canva hat eigene)
- Schatten + Transparency werden manchmal gerastert

### Empfehlung für Canva-Workflow
1. PPTX mit `cowork-visuals` + `sora-pptx-design` bauen
2. In PowerPoint / Keynote lokal öffnen → prüfen ob alles passt
3. In Canva uploaden
4. In Canva: Fonts neu zuweisen (Canva hat andere Default-Fonts)
5. Animationen in Canva neu hinzufügen (Canva hat einfachere aber hübschere Ein-/Aus-Animationen)

### Faustregel für Shapes
> Je weniger verschachtelte Gruppen, desto besser für Canva.
> Lieber 20 Einzel-Shapes als 3 gruppierte Komplex-Formen.

---

## Such-Query Templates für häufige Schulthemen

```javascript
const STOCK_QUERIES = {
  "Klima":             ["wind turbine dusk", "melting glacier", "climate protest"],
  "Haushalt":          ["euro coins stack", "german parliament", "budget spreadsheet"],
  "Verteidigung":      ["soldiers silhouette", "nato flag wind", "military aerial"],
  "Außenpolitik":      ["un general assembly", "diplomat handshake", "world flags"],
  "Migration":         ["border fence", "harbor boat", "refugee family silhouette"],
  "Demokratie":        ["ballot box", "voting hand", "parliament hall empty"],
  "Wirtschaft":        ["stock chart screen", "frankfurt skyline", "shipping containers"],
  "Digitalisierung":   ["server room blue", "hands keyboard", "fiber optic cables"],
  "Bildung":           ["classroom empty", "chalkboard math", "library books"],
  "Gesundheit":        ["hospital corridor", "medicine pills", "doctor stethoscope"]
};
```

---

## Quick-Start Checkliste für neue Präsentation

- [ ] Titelfolie → **Methode 1** (Hero-Foto) mit Overlay
- [ ] Mind. 1 Flow- oder Mindmap-Folie → **Methode 2** (native Shapes)
- [ ] Max. 1 Napkin-Diagramm auf wichtigster Folie → **Methode 3** (falls Budget)
- [ ] Alle Bilder in `assets/` Ordner ablegen und mitliefern
- [ ] Canva-Check (Fonts, Gruppen, Shapes) vor finaler Abgabe
