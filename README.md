# medaccur Playwright PPTX Renderer

FastAPI + Playwright + python-pptx service deployed on Railway.
Converts medaccur HTML slide templates into professional PPTX files.

---

## Architecture

```
Browser (MAP Generator)
    │
    │  POST /.netlify/functions/render-pptx
    │
Netlify Function (proxy)
    │
    │  POST https://medai-pptx-renderer.railway.app/render-pptx
    │
Railway Service (this repo)
    │
    ├── Playwright opens each slide URL on Netlify
    ├── Injects data via window.setXxxData()
    ├── Injects theme via window.setTheme()
    ├── Screenshots 1920×1080 → PNG
    └── python-pptx assembles PNGs → .pptx
```

---

## Deployment on Railway

### 1. Create new Railway service

```bash
# Option A: deploy from GitHub
# Push this folder to a GitHub repo
# Railway: New Project → Deploy from GitHub → select repo

# Option B: Railway CLI
railway login
railway init
railway up
```

### 2. Set environment variables in Railway

No secrets needed — the renderer is public.
Railway auto-sets `PORT`.

Optional:
- `ALLOWED_ORIGINS` — comma-separated allowed CORS origins (default: *)

### 3. Get the Railway URL

After deploy, copy the Railway URL e.g.:
```
https://medai-pptx-renderer-production.up.railway.app
```

### 4. Set Netlify env var

In Netlify dashboard → Site settings → Environment variables:
```
RAILWAY_RENDERER_URL = https://medai-pptx-renderer-production.up.railway.app
```

---

## Netlify integration

### Copy function file
```
netlify-function/render-pptx.js
→ netlify/functions/render-pptx.js
```

### Copy export patch
```
netlify-function/map_generator_export_patch.js
→ public/js/map_generator_export_patch.js
```

### Update map_generator.html

1. Add script tag before `</body>`:
```html
<script src="/js/map_generator_export_patch.js"></script>
```

2. Find the PPTX export button and change onclick:
```html
<!-- OLD -->
<button onclick="exportPptx()">Download PPTX</button>

<!-- NEW -->
<button onclick="exportViaRenderer()">Download PPTX</button>
```

3. Add a theme selector to Step 5 (optional):
```html
<select id="themeSelect">
  <option value="light">Light</option>
  <option value="dark">Dark</option>
  <option value="teal">Teal</option>
  <option value="slate">Slate</option>
</select>
```

---

## API Reference

### POST /render-pptx

```json
{
  "drug": "Belantamab Mafodotin",
  "indication": "Relapsed/Refractory Multiple Myeloma",
  "country": "Germany",
  "year": "2027",
  "theme": "light",
  "base_url": "https://medai-dashboard.netlify.app",
  "slides": [
    {
      "template": "slides/slide_swot.html",
      "section_id": "swot",
      "theme": "light",
      "data": {
        "strengths": ["Only ADC targeting BCMA", "..."],
        "weaknesses": ["..."],
        "opportunities": ["..."],
        "threats": ["..."]
      }
    }
  ]
}
```

Returns: `application/vnd.openxmlformats-officedocument.presentationml.presentation` binary

### GET /health

```json
{ "status": "ok", "service": "medaccur-renderer", "version": "1.0.0" }
```

### GET /slides

Returns list of all 23 slide templates with inject function names.

---

## Local development

```bash
# Install deps
pip install -r requirements.txt
playwright install chromium

# Run
uvicorn main:app --reload --port 8080

# Test
curl -X POST http://localhost:8080/render-pptx \
  -H "Content-Type: application/json" \
  -d '{
    "drug": "Belantamab Mafodotin",
    "indication": "RRMM",
    "country": "Germany",
    "year": "2027",
    "theme": "light",
    "base_url": "https://medai-dashboard.netlify.app",
    "slides": [
      {
        "template": "slides/slide_swot.html",
        "section_id": "swot",
        "data": {}
      }
    ]
  }' --output test.pptx
```

---

## Performance

| Slides | Render time |
|--------|-------------|
| 5      | ~15s        |
| 10     | ~30s        |
| 23     | ~60–90s     |

Bottleneck: Playwright page load + network (Netlify → Railway → Netlify).
Each slide loads fonts from Google Fonts and D3 from CDN.

**Optimisation options (future):**
- Cache fonts locally in Railway container
- Render slides in parallel (2–4 concurrent pages)
- Serve slides from Railway directly (avoid Netlify round-trip)
