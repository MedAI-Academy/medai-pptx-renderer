"""
medaccur Playwright PPTX Renderer
Railway Service — FastAPI + Playwright + python-pptx

POST /render-pptx   → accepts slide config JSON, returns .pptx binary
GET  /health        → health check
GET  /slides        → list available slide templates
"""

import os
import io
import asyncio
import logging
from typing import Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, JSONResponse
from pydantic import BaseModel
from renderer import render_presentation

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="medaccur PPTX Renderer",
    description="Playwright-based HTML→PPTX renderer for medaccur MAP Generator",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Restrict to medai-dashboard.netlify.app in production
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

# ── Request schema ──────────────────────────────────────────────────────────

class SlideConfig(BaseModel):
    """One slide to render."""
    template: str              # e.g. "slides/slide_swot.html"
    section_id: str            # e.g. "swot"
    data: dict                 # JSON data to inject via window.setXxxData()
    theme: str = "light"       # light | dark | teal | slate
    number: Optional[int] = None  # slide number for divider

class RenderRequest(BaseModel):
    """Full presentation render request."""
    drug: str
    indication: str
    country: str
    year: str
    theme: str = "light"
    base_url: str = "https://medai-dashboard.netlify.app"
    slides: list[SlideConfig]
    include_dividers: bool = True

# ── Endpoints ───────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok", "service": "medaccur-renderer", "version": "1.0.0"}


@app.get("/slides")
async def list_slides():
    """List all known slide templates."""
    return {
        "slides": [
            {"id": "divider",               "template": "slides/divider.html",                 "inject": "setDividerData"},
            {"id": "disease_intro",         "template": "slides/disease_introduction.html",    "inject": "setDiseaseData"},
            {"id": "prevalence",            "template": "slides/prevalence_kpis.html",          "inject": "setPrevalenceData"},
            {"id": "treatment_landscape",   "template": "slides/treatment_landscape.html",     "inject": "setTreatmentData"},
            {"id": "guidelines",            "template": "slides/guidelines.html",               "inject": "setGuidelineData"},
            {"id": "unmet_needs",           "template": "slides/content.html",                  "inject": "setContentData"},
            {"id": "moa",                   "template": "slides/mode_of_action.html",           "inject": "setMoaData"},
            {"id": "pivotal_studies",       "template": "slides/slide_pivotal_studies.html",    "inject": "setPivotalData"},
            {"id": "swimmer_plot",          "template": "slides/slide_swimmer.html",             "inject": "setSwimmerData"},
            {"id": "subgroup_analysis",     "template": "slides/subgroup_analysis.html",        "inject": "setSubgroupData"},
            {"id": "competitive_landscape", "template": "slides/competitive_landscape.html",   "inject": "setCompetitiveData"},
            {"id": "market_access",         "template": "slides/market_access.html",            "inject": "setMarketAccessData"},
            {"id": "market_access_timeline","template": "slides/market_access_timeline.html",  "inject": "setMarketAccessData"},
            {"id": "treatment_algorithm",   "template": "slides/treatment_algorithm.html",     "inject": "setAlgorithmData"},
            {"id": "differentiators",       "template": "slides/differentiators.html",          "inject": "setDifferentiatorsData"},
            {"id": "swot",                  "template": "slides/slide_swot.html",               "inject": "setSwotData"},
            {"id": "scientific_narrative",  "template": "slides/scientific_narrative.html",    "inject": "setNarrativeData"},
            {"id": "strategic_imperatives", "template": "slides/strategic_imperatives.html",   "inject": "setImperativesData"},
            {"id": "tactical_plan",         "template": "slides/tactical_plan.html",            "inject": "setTacticsData"},
            {"id": "kol_mapping",           "template": "slides/kol_mapping.html",              "inject": "setKOLData"},
            {"id": "insights",              "template": "slides/insights.html",                 "inject": "setInsightsData"},
            {"id": "timeline",              "template": "slides/timeline.html",                 "inject": "setTimelineData"},
            {"id": "executive_summary",     "template": "slides/executive_summary.html",       "inject": "setExecSummaryData"},
        ]
    }


@app.post("/render-pptx")
async def render_pptx(request: RenderRequest):
    """
    Render a full presentation to PPTX.

    - Opens each slide HTML via Playwright
    - Injects data + theme via window.setXxxData() / window.setTheme()
    - Screenshots at 1920×1080
    - Assembles PPTX with python-pptx
    - Returns binary PPTX
    """
    if not request.slides:
        raise HTTPException(status_code=400, detail="No slides provided")

    logger.info(f"Rendering {len(request.slides)} slides | drug={request.drug} | theme={request.theme}")

    try:
        pptx_bytes = await render_presentation(request)
        filename = f"{request.drug.replace(' ', '_')}_{request.country}_{request.year}.pptx"
        return Response(
            content=pptx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        logger.exception("Render failed")
        raise HTTPException(status_code=500, detail=str(e))
