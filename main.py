"""
medaccur Playwright PPTX Renderer — FastAPI App
"""
import os
import logging
from typing import Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="medaccur PPTX Renderer",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

# ── Pydantic models ────────────────────────────────────────────────────────

class SlideConfig(BaseModel):
    template: str
    section_id: str
    data: dict = {}
    theme: str = "light"
    number: Optional[int] = None

class RenderRequest(BaseModel):
    drug: str
    indication: str
    country: str
    year: str
    theme: str = "light"
    base_url: str = "https://medai-dashboard.netlify.app"
    slides: list[SlideConfig]

# ── Endpoints ──────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok", "service": "medaccur-renderer", "version": "1.0.0"}


@app.get("/slides")
async def list_slides():
    return {
        "slides": [
            {"id": "divider",               "template": "slides/divider.html"},
            {"id": "disease_intro",         "template": "slides/disease_introduction.html"},
            {"id": "prevalence",            "template": "slides/prevalence_kpis.html"},
            {"id": "treatment_landscape",   "template": "slides/treatment_landscape.html"},
            {"id": "guidelines",            "template": "slides/guidelines.html"},
            {"id": "unmet_needs",           "template": "slides/content.html"},
            {"id": "moa",                   "template": "slides/mode_of_action.html"},
            {"id": "pivotal_studies",       "template": "slides/slide_pivotal_studies.html"},
            {"id": "swimmer_plot",          "template": "slides/slide_swimmer.html"},
            {"id": "subgroup_analysis",     "template": "slides/subgroup_analysis.html"},
            {"id": "competitive_landscape", "template": "slides/competitive_landscape.html"},
            {"id": "market_access",         "template": "slides/market_access.html"},
            {"id": "swot",                  "template": "slides/slide_swot.html"},
            {"id": "differentiators",       "template": "slides/differentiators.html"},
            {"id": "scientific_narrative",  "template": "slides/scientific_narrative.html"},
            {"id": "strategic_imperatives", "template": "slides/strategic_imperatives.html"},
            {"id": "tactical_plan",         "template": "slides/tactical_plan.html"},
            {"id": "kol_mapping",           "template": "slides/kol_mapping.html"},
            {"id": "insights",              "template": "slides/insights.html"},
            {"id": "timeline",              "template": "slides/timeline.html"},
            {"id": "executive_summary",     "template": "slides/executive_summary.html"},
        ]
    }


@app.post("/render-pptx")
async def render_pptx(request: RenderRequest):
    if not request.slides:
        raise HTTPException(status_code=400, detail="No slides provided")

    logger.info(f"Rendering {len(request.slides)} slides | drug={request.drug} | theme={request.theme}")

    try:
        # Import here (not at module level) to avoid startup failures
        # if playwright has any init issues
        from renderer import render_presentation
        pptx_bytes = await render_presentation(request)
    except Exception as e:
        logger.exception("Render failed")
        raise HTTPException(status_code=500, detail=str(e))

    drug = request.drug.replace(" ", "_")
    filename = f"{drug}_{request.country}_{request.year}.pptx"
    return Response(
        content=pptx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
# ── Entrypoint ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    logger.info(f"Starting on port {port}")
    uvicorn.run("main:app", host="0.0.0.0", port=port, workers=1, log_level="info")
