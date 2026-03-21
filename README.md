# MedAI PPTX Renderer

**Java/Spring Boot service for generating professional PPTX presentations.**  
Replaces the python-pptx renderer with Apache POI — fixing the namespace/background bugs permanently.

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Runtime | Java 17 (Eclipse Temurin) |
| Framework | Spring Boot 3.3 |
| PPTX Engine | Apache POI 5.4 (XSLF) |
| Charts | JFreeChart 1.5 |
| Build | Gradle 8.5 |
| Deploy | Railway (Docker) |

## Quick Start (Local)

```bash
# 1. Clone
git clone https://github.com/MedAI-Academy/medai-pptx-renderer.git
cd medai-pptx-renderer

# 2. Build
./gradlew bootJar

# 3. Run
java -jar build/libs/medai-renderer.jar

# 4. Test health
curl http://localhost:8080/api/v1/health
```

## Deploy to Railway

### Option A: From GitHub (Recommended)

1. Push this repo to `MedAI-Academy/medai-pptx-renderer`
2. In Railway Dashboard → New Service → Deploy from GitHub
3. Select the repo → Railway auto-detects Dockerfile
4. Add env var: `PORT=8080`
5. Generate domain → use the same URL as the old Python renderer

### Option B: Railway CLI

```bash
railway login
railway init
railway up
```

## API

### POST /api/v1/render
Accepts JSON, returns PPTX binary.

```bash
curl -X POST http://localhost:8080/api/v1/render \
  -H "Content-Type: application/json" \
  -d @test-request.json \
  -o output.pptx
```

### POST /render
Legacy endpoint — backwards compatible with existing map_generator.html.

### GET /api/v1/health
Returns service status and capabilities.

### GET /api/v1/templates
Returns available themes and layouts.

## Slide Layouts

| Layout | Description |
|--------|-------------|
| TITLE | Title slide with KPIs and badges |
| TOC | Table of contents with numbered items |
| DIVIDER | Section divider with large number |
| CONTENT_FULL | Full-width body text |
| CONTENT_TWO_COL | Two-column layout |
| CONTENT_CARDS | 2×2 or 3×2 card grid |
| TABLE | Data table (studies, guidelines) |
| CHART_KM | Kaplan-Meier survival curve |
| SWOT | 2×2 SWOT matrix |
| TIMELINE | Horizontal timeline with milestones |
| KPI_DASHBOARD | Big-number KPI boxes |
| REFERENCES | Source list with tier indicators |
| CONFIDENCE | Score breakdown + time/cost savings |

## Migration from Python Renderer

The existing `map_generator.html` sends POST to `/render`.  
This Java service exposes the same `/render` endpoint.  
**No frontend changes needed** — just swap the Railway service.

### What's Fixed

- **Background colors**: No more xmlns namespace bug → correct in PowerPoint
- **Text overflow**: Auto-shrink enabled on all text boxes
- **Charts**: JFreeChart renders Kaplan-Meier as high-res PNG
- **Reliability**: No more empty slides from API timeouts (handled by frontend)
