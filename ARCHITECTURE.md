# MedAI PPTX Renderer — Java Architecture v1.0

## Warum Java statt Python?

| Problem mit python-pptx          | Lösung mit Apache POI (Java)         |
|----------------------------------|--------------------------------------|
| Inline xmlns auf `<p:bg>` → weiße Hintergründe in PowerPoint | Korrekte Namespace-Hierarchie, kein JSZip-Workaround nötig |
| Keine nativen Charts             | OOXML Chart-API + JFreeChart für KM-Kurven |
| Begrenzte Template-Unterstützung | Volle SlideMaster/SlideLayout-Kontrolle |
| Text-Overflow nicht kontrollierbar | `shrinkToFit`, `autoFit` direkt im XML |
| Kein Multi-Column-Layout         | Volle OOXML-Shape-Manipulation |

## System-Architektur

```
┌──────────────────────────────────────────────────────────────┐
│  NETLIFY (Frontend)                                          │
│  medai-dashboard.netlify.app                                 │
│                                                              │
│  ┌─────────────┐  ┌──────────────┐  ┌────────────────────┐  │
│  │ MAP Builder  │  │ Slide Builder│  │ Strategic Narrative │  │
│  │ (HTML/JS)    │  │ (HTML/JS)    │  │ (HTML/JS)          │  │
│  └──────┬───────┘  └──────┬───────┘  └────────┬───────────┘  │
│         │                 │                    │              │
│         │    Claude API (Netlify Function claude.js)          │
│         │    → DeepResearch + Web Search                     │
│         │    → Structured JSON mit References                │
│         │                 │                    │              │
│         ▼                 ▼                    ▼              │
│  ┌────────────────────────────────────────────────────┐      │
│  │  Unified JSON → POST to Java Renderer              │      │
│  │  { module, slides[], theme, metadata, references } │      │
│  └────────────────────────┬───────────────────────────┘      │
└───────────────────────────┼──────────────────────────────────┘
                            │ HTTPS
                            ▼
┌──────────────────────────────────────────────────────────────┐
│  RAILWAY (Java Backend)                                      │
│  medai-pptx-renderer-production.up.railway.app               │
│                                                              │
│  Spring Boot 3.x + Apache POI 5.4 + JFreeChart 1.5          │
│                                                              │
│  POST /api/v1/render                                         │
│  ├── SlideRouter → wählt Module-spezifischen Builder         │
│  ├── TemplateEngine → lädt .pptx Master-Template             │
│  ├── SlideBuilderService → baut Slides aus JSON              │
│  ├── ChartService → JFreeChart → PNG → Embed                │
│  ├── ConfidenceScoreService → berechnet + rendert Score      │
│  └── Response: application/octet-stream (PPTX Binary)        │
│                                                              │
│  POST /api/v1/health                                         │
│  GET  /api/v1/templates (Liste verfügbarer Templates)        │
└──────────────────────────────────────────────────────────────┘
```

## Drei Säulen: Liability, Zeitersparnis, Kostenersparnis

### 1. Confidence/Liability Score (≥95% Ziel)

Jeder Slide enthält im Footer:
- **Confidence Score Badge**: z.B. "Confidence: 97% | 12/12 Sources Verified"
- **Source References**: `[1] DREAMM-8, Dimopoulos et al., NEJM 2024` etc.

Auf dem letzten Slide: **Confidence Summary**
- Source Verification Rate (SV): Wie viele Quellen verifiziert?
- Traceability Rate (TR): Wie viele Claims haben eine Quelle?
- Source Quality Score (SQ): PubMed(T1)=100%, Conference(T2)=85%, Guideline(T3)=70%, Blog(T4)=40%
- Cross-Reference Score (CR): Mehrfach-Quellen pro Claim?
- **Gesamt = SV×0.35 + TR×0.30 + SQ×0.20 + CR×0.15**

### 2. Zeitersparnis

Anzeige auf Title-Slide und Confidence-Slide:
```
⏱ Generated in 4 min 23 sec
📊 Industry Benchmark: 2-4 weeks (80-160 hours)
💡 Time Savings: ~99.9%
```

### 3. Kostenersparnis

```
💰 Estimated Cost Savings: €3,200-€6,400 per MAP
   Based on: 80-160h × €40/h avg. Medical Writer rate
   MedAI Suite: €79/month (Premium)
```

## Java-Projekt-Struktur

```
medai-pptx-renderer/
├── build.gradle                      # Gradle Build mit allen Dependencies
├── Dockerfile                        # Multi-stage Build für Railway
├── settings.gradle
├── src/main/
│   ├── java/com/medai/renderer/
│   │   ├── MedaiRendererApplication.java    # Spring Boot Entry
│   │   ├── config/
│   │   │   ├── CorsConfig.java              # CORS für Netlify
│   │   │   └── WebConfig.java
│   │   ├── controller/
│   │   │   └── RenderController.java        # REST Endpoints
│   │   ├── model/
│   │   │   ├── RenderRequest.java           # Input JSON Model
│   │   │   ├── SlideData.java               # Einzelner Slide
│   │   │   ├── ChartData.java               # Chart-Daten
│   │   │   ├── ReferenceData.java           # Quellen-Referenz
│   │   │   └── ConfidenceScore.java         # Score-Modell
│   │   ├── service/
│   │   │   ├── PptxRenderService.java       # Haupt-Render-Logik
│   │   │   ├── SlideFactory.java            # Slide-Type-spezifische Builder
│   │   │   ├── ChartService.java            # JFreeChart → PNG → Embed
│   │   │   ├── ConfidenceService.java       # Score-Berechnung
│   │   │   └── TemplateService.java         # Template-Verwaltung
│   │   ├── template/
│   │   │   ├── ThemeConfig.java             # Farben, Fonts, Spacing
│   │   │   ├── SlideLayouts.java            # Layout-Definitionen
│   │   │   └── BrandAssets.java             # Logo, Icons
│   │   └── util/
│   │       ├── PptxUtils.java               # OOXML Helper-Methoden
│   │       └── ColorUtils.java              # Hex → POI Color
│   └── resources/
│       ├── application.yml                  # Spring Config
│       ├── templates/
│       │   ├── medai-dark.pptx              # Master-Template Dark
│       │   └── medai-light.pptx             # Master-Template Light
│       └── assets/
│           └── medai-logo.png               # Brand-Logo
└── src/test/java/com/medai/renderer/
    └── RenderControllerTest.java
```

## Slide-Design-System (besser als Prezent)

### Was Prezent NICHT kann, wir aber schon:
1. **Echte pharmazeutische Daten** mit verifizierten Referenzen
2. **Kaplan-Meier Kurven** als hochauflösende Charts
3. **Confidence Score** pro Slide und gesamt
4. **Web Search Integration** für aktuelle Daten (Prezent nutzt nur statische Templates)
5. **SWOT-Matrix** als echtes visuelles Element (nicht nur Text)
6. **Timeline** als echte grafische Timeline mit Milestones

### Design-Prinzipien

**Farb-System (MedAI Brand):**
```
Primary Dark:    #0B1A3B (Navy — Titel, Divider)
Primary Mid:     #0D2B4E (Dunkelblau — Content-BG)
Surface:         #163060 (Cards, Boxen)
Accent Purple:   #7C6FFF (Akzent, Highlights)
Accent Teal:     #22D3A5 (Positive, Charts)
Accent Gold:     #F5C842 (Warnings, KPIs)
Accent Rose:     #FF5F7E (Negative, Alerts)
Text White:      #EAF0FF (Haupttext auf Dark)
Text Muted:      #7B9FD4 (Sekundärtext)
Light BG:        #F0F4FF (Content-Slides Light-Variante)
```

**Font-System:**
```
Titel:    Calibri Bold, 32-44pt
Subtitle: Calibri, 18-24pt
Body:     Calibri, 12-16pt
Caption:  Calibri, 8-10pt, Muted
Mono:     Consolas, 9pt (für Referenzen)
```

**Slide-Layouts (14 Typen):**

| Layout              | Verwendung                          |
|---------------------|-------------------------------------|
| `TITLE`             | Titelfolie mit KPIs                 |
| `TOC`               | Inhaltsverzeichnis mit Hyperlinks   |
| `DIVIDER`           | Abschnitts-Trenner                  |
| `CONTENT_FULL`      | Volltext-Slide                      |
| `CONTENT_TWO_COL`   | Zwei-Spalten-Layout                 |
| `CONTENT_CARDS`     | 2x2 oder 3x2 Karten-Grid           |
| `TABLE`             | Daten-Tabelle (Studien, Guidelines) |
| `CHART_KM`          | Kaplan-Meier Kurve                  |
| `CHART_BAR`         | Balkendiagramm                      |
| `SWOT`              | 2x2 SWOT-Matrix                     |
| `TIMELINE`          | Grafische Timeline                  |
| `KPI_DASHBOARD`     | KPI-Boxen mit großen Zahlen         |
| `REFERENCES`        | Quellen-Slide                       |
| `CONFIDENCE`        | Confidence Score Summary            |

## API-Spezifikation

### POST /api/v1/render

**Request Body (JSON):**
```json
{
  "module": "map",
  "theme": "dark",
  "widescreen": true,
  "metadata": {
    "title": "Medical Affairs Plan",
    "drug": "Belantamab Mafodotin",
    "indication": "Multiple Myeloma",
    "company": "GSK",
    "mapType": "country",
    "country": "Germany",
    "year": 2027,
    "generatedAt": "2026-03-21T10:00:00Z",
    "generationTimeSeconds": 263
  },
  "confidenceScore": {
    "overall": 97,
    "sourceVerification": 98,
    "traceability": 95,
    "sourceQuality": 96,
    "crossReference": 88,
    "totalClaims": 48,
    "verifiedSources": 52,
    "totalSources": 54
  },
  "slides": [
    {
      "id": "title",
      "layout": "TITLE",
      "section": "Title",
      "content": {
        "title": "Belantamab Mafodotin",
        "subtitle": "Country Medical Affairs Plan — Germany 2027",
        "badges": ["Country MAP", "Oncology", "3L+ RRMM"],
        "kpis": [
          {"label": "Confidence", "value": "97%", "color": "teal"},
          {"label": "Sources", "value": "54", "color": "accent"},
          {"label": "Generated", "value": "4m 23s", "color": "gold"}
        ]
      }
    },
    {
      "id": "pivotal",
      "layout": "TABLE",
      "section": "Pivotal Studies",
      "content": {
        "title": "Pivotal Clinical Evidence",
        "subtitle": "Phase 2/3 Studies with Belantamab Mafodotin",
        "table": {
          "headers": ["Study", "Phase", "Regimen", "N", "mPFS (mo)", "ORR", "Key AEs"],
          "rows": [
            ["DREAMM-7", "Phase 3", "BVd vs DVd", "494", "36.6 vs 13.4", "83% vs 72%", "Keratopathy 43%"],
            ["DREAMM-8", "Phase 3", "BPd vs Pd", "302", "NR vs 12.7", "77% vs 55%", "Keratopathy 38%"]
          ]
        },
        "references": [
          {"id": "ref-1", "text": "Dimopoulos MA et al. NEJM 2024;391:1–12", "tier": 1, "type": "pubmed"},
          {"id": "ref-2", "text": "Trudel S et al. Lancet 2024;403:1230–40", "tier": 1, "type": "pubmed"}
        ]
      }
    },
    {
      "id": "km_curve",
      "layout": "CHART_KM",
      "section": "Pivotal Studies",
      "content": {
        "title": "DREAMM-7: Progression-Free Survival",
        "chartData": {
          "arms": [
            {
              "name": "BVd (Belantamab + Bortezomib + Dex)",
              "color": "#22D3A5",
              "timepoints": [0, 6, 12, 18, 24, 30, 36],
              "survival": [1.0, 0.82, 0.72, 0.65, 0.60, 0.57, 0.55],
              "atRisk": [247, 210, 185, 162, 140, 120, 98],
              "median": 36.6
            },
            {
              "name": "DVd (Daratumumab + Bortezomib + Dex)",
              "color": "#FF5F7E",
              "timepoints": [0, 6, 12, 18, 24, 30, 36],
              "survival": [1.0, 0.70, 0.55, 0.42, 0.35, 0.28, 0.22],
              "atRisk": [247, 180, 140, 108, 85, 62, 45],
              "median": 13.4
            }
          ],
          "xlabel": "Time (months)",
          "ylabel": "Progression-Free Survival",
          "hazardRatio": "HR 0.41 (95% CI 0.31–0.53), p<0.001"
        },
        "references": [
          {"id": "ref-1", "text": "Dimopoulos MA et al. NEJM 2024;391:1–12", "tier": 1}
        ]
      }
    }
  ]
}
```

**Response:**
- Content-Type: `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- Body: PPTX Binary

## Migration: Python → Java

### Phase 1 — Java Basis (1 Woche)
1. Spring Boot Projekt auf GitHub erstellen
2. `/api/v1/render` Endpoint mit Basis-Slides (Title, Divider, Content)
3. Deployment auf Railway (gleiche URL, Python-Service ersetzen)
4. `map_generator.html` JSON-Format anpassen

### Phase 2 — Design Excellence (1 Woche)
1. Master-Templates (.pptx) mit korrekten SlideMasters/Layouts
2. Alle 14 Slide-Layouts implementieren
3. Korrekte Backgrounds (kein xmlns-Bug mehr)
4. Professionelle Typografie und Spacing

### Phase 3 — Charts + Confidence (1 Woche)
1. JFreeChart Integration für Kaplan-Meier Kurven
2. SWOT-Matrix als visuelles Element
3. Timeline als grafische Shapes
4. Confidence Score Slide + Footer-Integration

### Phase 4 — Universeller Renderer (1 Woche)
1. Slide Builder Migration
2. Strategic Narrative Migration
3. Clinical Trial Report PPTX Migration
4. Template-Auswahl pro Modul

## Railway Deployment

**Dockerfile (Multi-Stage):**
```dockerfile
FROM gradle:8.5-jdk17 AS build
WORKDIR /app
COPY . .
RUN gradle bootJar --no-daemon

FROM eclipse-temurin:17-jre-alpine
WORKDIR /app
COPY --from=build /app/build/libs/*.jar app.jar
EXPOSE 8080
ENTRYPOINT ["java", "-jar", "app.jar"]
```

**Railway Setup:**
1. Neues GitHub Repo: `MedAI-Academy/medai-pptx-renderer`
2. Railway → neues Service aus GitHub Repo
3. Auto-Detect: Dockerfile → Build → Deploy
4. Environment Variables: `PORT=8080`, `SPRING_PROFILES_ACTIVE=production`
5. Custom Domain: `medai-pptx-renderer-production.up.railway.app` (gleiche URL!)

## Konkurrenz-Vergleich

| Feature                    | Prezent Premium | Astrid AI | MedAI Suite (Ziel) |
|----------------------------|:--------------:|:---------:|:-------------------:|
| AI-generierte Slides       | ✅             | ✅        | ✅                  |
| Brand Templates            | ✅ (35K+)      | ✅        | ✅ (Custom)         |
| Web Search / Live Data     | ❌             | ❌        | ✅ **Unique**       |
| Pharma-spezifische Daten   | ⚠️ (generisch) | ⚠️        | ✅ **Deep**         |
| Confidence/Liability Score | ❌             | ❌        | ✅ **Unique**       |
| Verifizierte Referenzen    | ❌             | ❌        | ✅ **Unique**       |
| Kaplan-Meier Kurven        | ❌             | ❌        | ✅ **Unique**       |
| SWOT-Matrix visuell        | ⚠️ (Template)  | ❌        | ✅                  |
| Zeitersparnis-Anzeige      | ✅ "90%"       | ✅ "70-80%"| ✅ + exakte Messung |
| Kostenersparnis            | ⚠️ (Marketing) | ❌        | ✅ **Kalkuliert**   |
| Preis                      | Enterprise $$$ | Enterprise $$$ | Ab €79/Monat   |
