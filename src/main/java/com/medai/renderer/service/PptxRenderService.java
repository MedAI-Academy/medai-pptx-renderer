package com.medai.renderer.service;

import com.medai.renderer.model.ConfidenceScore;
import com.medai.renderer.model.RenderRequest;
import com.medai.renderer.model.SlideData;
import com.medai.renderer.template.ThemeConfig;
import com.medai.renderer.util.PptxUtils;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

/**
 * Core PPTX rendering engine.
 * Takes a RenderRequest (JSON model) and produces a PPTX byte array.
 *
 * This replaces the python-pptx Railway renderer with:
 * - Correct XML namespace handling (no more white backgrounds)
 * - Professional slide layouts with theme system
 * - Embedded charts (Kaplan-Meier, SWOT) via ChartService
 * - Confidence Score integration on every slide
 */
@Service
public class PptxRenderService {

    private final ChartService chartService;

    public PptxRenderService(ChartService chartService) {
        this.chartService = chartService;
    }

    /**
     * Render a complete PPTX presentation from a RenderRequest.
     * @return PPTX file as byte array
     */
    public byte[] render(RenderRequest request) throws Exception {
        XMLSlideShow pptx = new XMLSlideShow();

        // Set widescreen dimensions (13.33" × 7.5" = 960pt × 540pt)
        if (request.isWidescreen()) {
            pptx.setPageSize(new Dimension(960, 540));
        }

        // Build each slide
        for (SlideData slideData : request.getSlides()) {
            XSLFSlide slide = pptx.createSlide();
            String layout = slideData.getLayout();
            String bgHex = ThemeConfig.bgColorFor(layout, request.getTheme());

            // Set background color (correct way — no xmlns bug)
            PptxUtils.setBackground(slide, bgHex);

            // Route to layout-specific builder
            switch (layout != null ? layout : "CONTENT_FULL") {
                case "TITLE"         -> buildTitleSlide(slide, pptx, slideData, request);
                case "TOC"           -> buildTocSlide(slide, slideData, request);
                case "DIVIDER"       -> buildDividerSlide(slide, slideData);
                case "CONTENT_FULL"  -> buildContentSlide(slide, slideData);
                case "CONTENT_TWO_COL" -> buildTwoColSlide(slide, slideData);
                case "CONTENT_CARDS" -> buildCardsSlide(slide, slideData);
                case "TABLE"         -> buildTableSlide(slide, slideData);
                case "CHART_KM"      -> buildKmChartSlide(slide, pptx, slideData);
                case "SWOT"          -> buildSwotSlide(slide, pptx, slideData);
                case "TIMELINE"      -> buildTimelineSlide(slide, slideData);
                case "KPI_DASHBOARD" -> buildKpiSlide(slide, slideData);
                case "REFERENCES"    -> buildReferencesSlide(slide, slideData);
                case "CONFIDENCE"    -> buildConfidenceSlide(slide, slideData, request);
                default              -> buildContentSlide(slide, slideData);
            }
        }

        // Write to byte array
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        pptx.write(baos);
        pptx.close();
        return baos.toByteArray();
    }

    // ═══════════════════════════════════════════════════════════
    // TITLE SLIDE
    // ═══════════════════════════════════════════════════════════

    private void buildTitleSlide(XSLFSlide slide, XMLSlideShow pptx,
                                  SlideData data, RenderRequest req) {
        // Accent bars (top + bottom)
        PptxUtils.addRect(slide, 0, 0, ThemeConfig.SLIDE_W, 0.06, ThemeConfig.HEX_ACCENT);
        PptxUtils.addRect(slide, 0, ThemeConfig.SLIDE_H - 0.06, ThemeConfig.SLIDE_W, 0.06, ThemeConfig.HEX_TEAL);

        // Drug name (large)
        PptxUtils.addText(slide, data.contentStr("title"),
            0.8, 1.2, 11.0, 1.5,
            ThemeConfig.SIZE_TITLE + 4, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
            true, TextParagraph.TextAlign.LEFT);

        // Subtitle
        PptxUtils.addText(slide, data.contentStr("subtitle"),
            0.8, 2.7, 11.0, 0.8,
            ThemeConfig.SIZE_SUBTITLE, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED,
            false, TextParagraph.TextAlign.LEFT);

        // Badges
        List<Map<String, Object>> badges = data.contentList("badges");
        double bx = 0.8;
        for (Map<String, Object> badge : badges) {
            String text = badge != null ? badge.toString() : "";
            if (badge instanceof String s) text = s;
            PptxUtils.addRect(slide, bx, 3.6, 2.2, 0.38, ThemeConfig.HEX_SURFACE);
            PptxUtils.addText(slide, text, bx + 0.1, 3.6, 2.0, 0.38,
                ThemeConfig.SIZE_SMALL, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEAL,
                true, TextParagraph.TextAlign.CENTER);
            bx += 2.4;
        }

        // KPI boxes
        List<Map<String, Object>> kpis = data.contentList("kpis");
        double kx = 0.8;
        for (Map<String, Object> kpi : kpis) {
            String label = kpi.getOrDefault("label", "").toString();
            String value = kpi.getOrDefault("value", "").toString();
            String colorKey = kpi.getOrDefault("color", "accent").toString();
            String colorHex = resolveColor(colorKey);

            PptxUtils.addRect(slide, kx, 4.4, 3.0, 1.2, ThemeConfig.HEX_SURFACE);
            PptxUtils.addText(slide, value, kx, 4.4, 3.0, 0.8,
                ThemeConfig.SIZE_TITLE - 4, ThemeConfig.FONT_TITLE, colorHex,
                true, TextParagraph.TextAlign.CENTER);
            PptxUtils.addText(slide, label, kx, 5.1, 3.0, 0.4,
                ThemeConfig.SIZE_SMALL, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED,
                false, TextParagraph.TextAlign.CENTER);
            kx += 3.3;
        }

        // Generation info
        PptxUtils.addText(slide,
            "Generated: " + req.meta("generatedAt") + " | Source: MedAI Suite DeepResearch",
            0.8, 6.4, 10.0, 0.3,
            ThemeConfig.SIZE_CAPTION, ThemeConfig.FONT_BODY, ThemeConfig.HEX_DIM,
            false, TextParagraph.TextAlign.LEFT);
    }

    // ═══════════════════════════════════════════════════════════
    // TABLE OF CONTENTS
    // ═══════════════════════════════════════════════════════════

    private void buildTocSlide(XSLFSlide slide, SlideData data, RenderRequest req) {
        PptxUtils.addHeader(slide, "Contents", req.meta("drug"));

        List<Map<String, Object>> items = data.contentList("items");
        double y = ThemeConfig.CONTENT_Y + 0.2;
        int colBreak = (items.size() + 1) / 2;

        for (int i = 0; i < items.size(); i++) {
            Map<String, Object> item = items.get(i);
            String num = String.valueOf(i + 1);
            String title = item.getOrDefault("title", "").toString();
            double x = (i < colBreak) ? 0.8 : 7.0;
            double cy = (i < colBreak) ? y + i * 0.55 : y + (i - colBreak) * 0.55;

            // Number badge
            PptxUtils.addRect(slide, x, cy, 0.45, 0.40, ThemeConfig.HEX_ACCENT);
            PptxUtils.addText(slide, num, x, cy, 0.45, 0.40,
                ThemeConfig.SIZE_BODY, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
                true, TextParagraph.TextAlign.CENTER);

            // Title
            PptxUtils.addText(slide, title, x + 0.55, cy, 5.0, 0.40,
                ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
                false, TextParagraph.TextAlign.LEFT);
        }

        PptxUtils.addFooter(slide, null);
    }

    // ═══════════════════════════════════════════════════════════
    // DIVIDER SLIDE
    // ═══════════════════════════════════════════════════════════

    private void buildDividerSlide(XSLFSlide slide, SlideData data) {
        // Accent bar top
        PptxUtils.addRect(slide, 0, 0, ThemeConfig.SLIDE_W, 0.04, ThemeConfig.HEX_ACCENT);

        // Large section number
        PptxUtils.addText(slide, String.format("%02d", data.getSectionIndex()),
            0.8, 1.5, 4.0, 2.5,
            80.0, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_ACCENT,
            true, TextParagraph.TextAlign.LEFT);

        // Section title
        PptxUtils.addText(slide, data.contentStr("title"),
            0.8, 3.8, 10.0, 1.2,
            ThemeConfig.SIZE_TITLE, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
            true, TextParagraph.TextAlign.LEFT);

        // Subtitle
        String subtitle = data.contentStr("subtitle");
        if (!subtitle.isEmpty()) {
            PptxUtils.addText(slide, subtitle,
                0.8, 5.0, 10.0, 0.6,
                ThemeConfig.SIZE_SUBTITLE, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED,
                false, TextParagraph.TextAlign.LEFT);
        }

        // Bottom accent
        PptxUtils.addRect(slide, 0, ThemeConfig.SLIDE_H - 0.04, ThemeConfig.SLIDE_W, 0.04, ThemeConfig.HEX_TEAL);
    }

    // ═══════════════════════════════════════════════════════════
    // CONTENT SLIDES (full-width text)
    // ═══════════════════════════════════════════════════════════

    private void buildContentSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        // Main body text
        String body = data.contentStr("body");
        if (!body.isEmpty()) {
            PptxUtils.addText(slide, body,
                ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y,
                ThemeConfig.CONTENT_W, ThemeConfig.CONTENT_H,
                ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
                false, TextParagraph.TextAlign.LEFT);
        }

        // Bullet items
        List<Map<String, Object>> items = data.contentList("items");
        if (!items.isEmpty()) {
            String[] texts = items.stream()
                .map(m -> m.getOrDefault("text", m.toString()).toString())
                .toArray(String[]::new);
            PptxUtils.addBulletText(slide, texts,
                ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y,
                ThemeConfig.CONTENT_W, ThemeConfig.CONTENT_H,
                ThemeConfig.SIZE_BODY, ThemeConfig.HEX_TEXT);
        }

        // References in footer
        addReferenceFooter(slide, data);
        PptxUtils.addFooter(slide, buildConfidenceBadge(data));
    }

    // ═══════════════════════════════════════════════════════════
    // TWO-COLUMN LAYOUT
    // ═══════════════════════════════════════════════════════════

    private void buildTwoColSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        double colW = (ThemeConfig.CONTENT_W - 0.4) / 2;

        // Left column
        PptxUtils.addText(slide, data.contentStr("leftTitle"),
            ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y, colW, 0.5,
            ThemeConfig.SIZE_HEADING - 2, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_TEAL,
            true, TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(slide, data.contentStr("leftBody"),
            ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y + 0.55, colW, ThemeConfig.CONTENT_H - 0.55,
            ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
            false, TextParagraph.TextAlign.LEFT);

        // Right column
        double rx = ThemeConfig.CONTENT_X + colW + 0.4;
        PptxUtils.addText(slide, data.contentStr("rightTitle"),
            rx, ThemeConfig.CONTENT_Y, colW, 0.5,
            ThemeConfig.SIZE_HEADING - 2, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_ACCENT,
            true, TextParagraph.TextAlign.LEFT);
        PptxUtils.addText(slide, data.contentStr("rightBody"),
            rx, ThemeConfig.CONTENT_Y + 0.55, colW, ThemeConfig.CONTENT_H - 0.55,
            ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
            false, TextParagraph.TextAlign.LEFT);

        // Vertical divider
        PptxUtils.addRect(slide,
            ThemeConfig.CONTENT_X + colW + 0.17, ThemeConfig.CONTENT_Y,
            0.03, ThemeConfig.CONTENT_H, ThemeConfig.HEX_ACCENT);

        addReferenceFooter(slide, data);
        PptxUtils.addFooter(slide, buildConfidenceBadge(data));
    }

    // ═══════════════════════════════════════════════════════════
    // CARDS LAYOUT (2x2 or 3x2 grid)
    // ═══════════════════════════════════════════════════════════

    private void buildCardsSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        List<Map<String, Object>> cards = data.contentList("cards");
        int cols = cards.size() <= 4 ? 2 : 3;
        double cardW = (ThemeConfig.CONTENT_W - (cols - 1) * 0.25) / cols;
        double cardH = 2.4;

        for (int i = 0; i < Math.min(cards.size(), 6); i++) {
            Map<String, Object> card = cards.get(i);
            int col = i % cols, row = i / cols;
            double cx = ThemeConfig.CONTENT_X + col * (cardW + 0.25);
            double cy = ThemeConfig.CONTENT_Y + row * (cardH + 0.25);
            String accentHex = ThemeConfig.ACCENT_CYCLE[i % ThemeConfig.ACCENT_CYCLE.length];

            // Card background
            PptxUtils.addRect(slide, cx, cy, cardW, cardH, ThemeConfig.HEX_SURFACE);
            // Accent bar top of card
            PptxUtils.addRect(slide, cx, cy, cardW, 0.04, accentHex);

            // Card title
            PptxUtils.addText(slide, card.getOrDefault("title", "").toString(),
                cx + 0.15, cy + 0.12, cardW - 0.3, 0.4,
                ThemeConfig.SIZE_BODY + 1, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
                true, TextParagraph.TextAlign.LEFT);

            // Card body
            PptxUtils.addText(slide, card.getOrDefault("body", "").toString(),
                cx + 0.15, cy + 0.55, cardW - 0.3, cardH - 0.7,
                ThemeConfig.SIZE_BODY - 1, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
                false, TextParagraph.TextAlign.LEFT);
        }

        PptxUtils.addFooter(slide, buildConfidenceBadge(data));
    }

    // ═══════════════════════════════════════════════════════════
    // TABLE LAYOUT
    // ═══════════════════════════════════════════════════════════

    private void buildTableSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        Map<String, Object> tableData = data.contentMap("table");
        if (tableData.isEmpty()) {
            buildContentSlide(slide, data); // Fallback
            return;
        }

        @SuppressWarnings("unchecked")
        List<String> headers = ((List<Object>) tableData.getOrDefault("headers", List.of()))
            .stream().map(Object::toString).toList();
        @SuppressWarnings("unchecked")
        List<List<String>> rows = ((List<List<Object>>) tableData.getOrDefault("rows", List.of()))
            .stream().map(row -> row.stream().map(Object::toString).toList()).toList();

        String[] headerArr = headers.toArray(new String[0]);
        String[][] rowArr = rows.stream().map(r -> r.toArray(new String[0])).toArray(String[][]::new);

        PptxUtils.addTable(slide, headerArr, rowArr,
            ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y,
            ThemeConfig.CONTENT_W, 0.40);

        addReferenceFooter(slide, data);
        PptxUtils.addFooter(slide, buildConfidenceBadge(data));
    }

    // ═══════════════════════════════════════════════════════════
    // KAPLAN-MEIER CHART
    // ═══════════════════════════════════════════════════════════

    private void buildKmChartSlide(XSLFSlide slide, XMLSlideShow pptx, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        Map<String, Object> chartData = data.contentMap("chartData");
        if (!chartData.isEmpty()) {
            try {
                byte[] chartPng = chartService.generateKaplanMeier(chartData);
                PptxUtils.addImage(slide, pptx, chartPng,
                    ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y,
                    ThemeConfig.CONTENT_W, ThemeConfig.CONTENT_H - 0.5);
            } catch (Exception e) {
                PptxUtils.addText(slide, "Chart generation failed: " + e.getMessage(),
                    ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y + 2,
                    ThemeConfig.CONTENT_W, 1.0,
                    ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_ROSE,
                    false, TextParagraph.TextAlign.CENTER);
            }
        }

        addReferenceFooter(slide, data);
        PptxUtils.addFooter(slide, buildConfidenceBadge(data));
    }

    // ═══════════════════════════════════════════════════════════
    // SWOT MATRIX
    // ═══════════════════════════════════════════════════════════

    private void buildSwotSlide(XSLFSlide slide, XMLSlideShow pptx, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        Map<String, Object> swotData = data.contentMap("swotData");
        if (!swotData.isEmpty()) {
            try {
                byte[] swotPng = chartService.generateSwotMatrix(swotData);
                PptxUtils.addImage(slide, pptx, swotPng,
                    ThemeConfig.CONTENT_X + 0.5, ThemeConfig.CONTENT_Y,
                    ThemeConfig.CONTENT_W - 1.0, ThemeConfig.CONTENT_H - 0.3);
            } catch (Exception e) {
                PptxUtils.addText(slide, "SWOT generation failed: " + e.getMessage(),
                    ThemeConfig.CONTENT_X, ThemeConfig.CONTENT_Y + 2,
                    ThemeConfig.CONTENT_W, 1.0,
                    ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_ROSE,
                    false, TextParagraph.TextAlign.CENTER);
            }
        }

        PptxUtils.addFooter(slide, buildConfidenceBadge(data));
    }

    // ═══════════════════════════════════════════════════════════
    // TIMELINE
    // ═══════════════════════════════════════════════════════════

    private void buildTimelineSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        List<Map<String, Object>> events = data.contentList("events");
        if (events.isEmpty()) return;

        // Horizontal timeline bar
        double barY = ThemeConfig.CONTENT_Y + 2.0;
        PptxUtils.addRect(slide, ThemeConfig.CONTENT_X, barY, ThemeConfig.CONTENT_W, 0.06, ThemeConfig.HEX_ACCENT);

        double spacing = ThemeConfig.CONTENT_W / Math.max(events.size(), 1);

        for (int i = 0; i < events.size(); i++) {
            Map<String, Object> event = events.get(i);
            double ex = ThemeConfig.CONTENT_X + i * spacing + spacing / 2 - 0.75;

            // Dot on timeline
            PptxUtils.addRect(slide, ex + 0.65, barY - 0.08, 0.20, 0.20,
                ThemeConfig.ACCENT_CYCLE[i % ThemeConfig.ACCENT_CYCLE.length]);

            // Event date (above)
            PptxUtils.addText(slide, event.getOrDefault("date", "").toString(),
                ex, barY - 0.7, 1.5, 0.5,
                ThemeConfig.SIZE_SMALL, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_TEAL,
                true, TextParagraph.TextAlign.CENTER);

            // Event title (below)
            PptxUtils.addText(slide, event.getOrDefault("title", "").toString(),
                ex, barY + 0.25, 1.5, 0.4,
                ThemeConfig.SIZE_SMALL, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
                true, TextParagraph.TextAlign.CENTER);

            // Event detail (below title)
            PptxUtils.addText(slide, event.getOrDefault("detail", "").toString(),
                ex, barY + 0.65, 1.5, 1.2,
                ThemeConfig.SIZE_CAPTION + 1, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED,
                false, TextParagraph.TextAlign.CENTER);
        }

        PptxUtils.addFooter(slide, null);
    }

    // ═══════════════════════════════════════════════════════════
    // KPI DASHBOARD
    // ═══════════════════════════════════════════════════════════

    private void buildKpiSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, data.contentStr("title"), data.getSection());

        List<Map<String, Object>> kpis = data.contentList("kpis");
        int cols = Math.min(kpis.size(), 4);
        double kpiW = (ThemeConfig.CONTENT_W - (cols - 1) * 0.3) / cols;
        double kpiH = 2.0;

        for (int i = 0; i < Math.min(kpis.size(), 8); i++) {
            Map<String, Object> kpi = kpis.get(i);
            int col = i % cols, row = i / cols;
            double kx = ThemeConfig.CONTENT_X + col * (kpiW + 0.3);
            double ky = ThemeConfig.CONTENT_Y + row * (kpiH + 0.3);
            String accentHex = ThemeConfig.ACCENT_CYCLE[i % ThemeConfig.ACCENT_CYCLE.length];

            PptxUtils.addRect(slide, kx, ky, kpiW, kpiH, ThemeConfig.HEX_SURFACE);
            PptxUtils.addRect(slide, kx, ky, kpiW, 0.04, accentHex);

            // Big number
            PptxUtils.addText(slide, kpi.getOrDefault("value", "").toString(),
                kx, ky + 0.2, kpiW, 1.0,
                ThemeConfig.SIZE_TITLE + 8, ThemeConfig.FONT_TITLE, accentHex,
                true, TextParagraph.TextAlign.CENTER);

            // Label
            PptxUtils.addText(slide, kpi.getOrDefault("label", "").toString(),
                kx, ky + 1.3, kpiW, 0.5,
                ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_MUTED,
                false, TextParagraph.TextAlign.CENTER);
        }

        PptxUtils.addFooter(slide, null);
    }

    // ═══════════════════════════════════════════════════════════
    // REFERENCES SLIDE
    // ═══════════════════════════════════════════════════════════

    private void buildReferencesSlide(XSLFSlide slide, SlideData data) {
        PptxUtils.addHeader(slide, "References", data.getSection());

        List<Map<String, Object>> refs = data.contentList("references");
        double y = ThemeConfig.CONTENT_Y;

        for (int i = 0; i < refs.size(); i++) {
            Map<String, Object> ref = refs.get(i);
            String text = "[" + (i + 1) + "] " + ref.getOrDefault("text", "").toString();
            int tier = ref.containsKey("tier") ? ((Number) ref.get("tier")).intValue() : 4;
            String tierColor = switch (tier) {
                case 1 -> ThemeConfig.HEX_TEAL;    // Gold standard (PubMed, NCT)
                case 2 -> ThemeConfig.HEX_ACCENT;   // Conference
                case 3 -> ThemeConfig.HEX_GOLD;     // Guidelines
                default -> ThemeConfig.HEX_DIM;     // Other
            };

            // Tier dot
            PptxUtils.addRect(slide, ThemeConfig.CONTENT_X, y + 0.05, 0.12, 0.12, tierColor);

            // Reference text
            PptxUtils.addText(slide, text,
                ThemeConfig.CONTENT_X + 0.2, y, ThemeConfig.CONTENT_W - 0.3, 0.3,
                ThemeConfig.SIZE_CAPTION + 1, ThemeConfig.FONT_MONO, ThemeConfig.HEX_TEXT,
                false, TextParagraph.TextAlign.LEFT);
            y += 0.35;
        }

        PptxUtils.addFooter(slide, null);
    }

    // ═══════════════════════════════════════════════════════════
    // CONFIDENCE SCORE SLIDE
    // ═══════════════════════════════════════════════════════════

    private void buildConfidenceSlide(XSLFSlide slide, SlideData data, RenderRequest req) {
        PptxUtils.addRect(slide, 0, 0, ThemeConfig.SLIDE_W, 0.04, ThemeConfig.HEX_TEAL);

        ConfidenceScore cs = req.getConfidenceScore();
        if (cs == null) cs = new ConfidenceScore();

        // Title
        PptxUtils.addText(slide, "Confidence Score: " + cs.getOverall() + "% (" + cs.getGrade() + ")",
            0.8, 0.5, 10.0, 1.0,
            ThemeConfig.SIZE_TITLE, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_TEAL,
            true, TextParagraph.TextAlign.LEFT);

        // Score breakdown boxes
        String[][] scores = {
            {"Source Verification", cs.getSourceVerification() + "%", "35%", ThemeConfig.HEX_TEAL},
            {"Traceability", cs.getTraceability() + "%", "30%", ThemeConfig.HEX_ACCENT},
            {"Source Quality", cs.getSourceQuality() + "%", "20%", ThemeConfig.HEX_GOLD},
            {"Cross-Reference", cs.getCrossReference() + "%", "15%", ThemeConfig.HEX_ROSE}
        };

        double bx = 0.8;
        for (String[] score : scores) {
            PptxUtils.addRect(slide, bx, 2.0, 2.8, 1.8, ThemeConfig.HEX_SURFACE);
            PptxUtils.addRect(slide, bx, 2.0, 2.8, 0.04, score[3]);

            PptxUtils.addText(slide, score[1], bx, 2.2, 2.8, 0.9,
                ThemeConfig.SIZE_TITLE, ThemeConfig.FONT_TITLE, score[3],
                true, TextParagraph.TextAlign.CENTER);
            PptxUtils.addText(slide, score[0], bx, 3.0, 2.8, 0.4,
                ThemeConfig.SIZE_SMALL, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
                false, TextParagraph.TextAlign.CENTER);
            PptxUtils.addText(slide, "Weight: " + score[2], bx, 3.4, 2.8, 0.3,
                ThemeConfig.SIZE_CAPTION, ThemeConfig.FONT_BODY, ThemeConfig.HEX_DIM,
                false, TextParagraph.TextAlign.CENTER);
            bx += 3.1;
        }

        // Time + Cost Savings
        String genTime = req.meta("generationTimeSeconds");
        if (!genTime.isEmpty()) {
            int secs = Integer.parseInt(genTime);
            String timeStr = (secs / 60) + "m " + (secs % 60) + "s";
            PptxUtils.addText(slide, "⏱ Generated in " + timeStr + "   |   " +
                "📊 Industry Benchmark: 2-4 weeks (80-160h)   |   " +
                "💡 Time Savings: ~99.9%",
                0.8, 4.3, 11.0, 0.5,
                ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_TEXT,
                false, TextParagraph.TextAlign.LEFT);
        }

        PptxUtils.addText(slide,
            "💰 Estimated Cost Savings: €3,200–€6,400 per MAP\n" +
            "Based on 80–160h × €40/h avg. Medical Writer rate vs. MedAI Suite €79/month",
            0.8, 5.0, 11.0, 0.8,
            ThemeConfig.SIZE_BODY, ThemeConfig.FONT_BODY, ThemeConfig.HEX_GOLD,
            false, TextParagraph.TextAlign.LEFT);

        // Methodology note
        PptxUtils.addText(slide,
            "Methodology: SV×0.35 + TR×0.30 + SQ×0.20 + CR×0.15 | " +
            "Source Tiers: T1=PubMed/NCT (100%), T2=Conference (85%), T3=Guideline (70%), T4=Other (40%) | " +
            "EU AI Act Classification: LIMITED-RISK (Art 6(3) exemption)",
            0.8, 6.2, 11.5, 0.6,
            ThemeConfig.SIZE_CAPTION, ThemeConfig.FONT_MONO, ThemeConfig.HEX_DIM,
            false, TextParagraph.TextAlign.LEFT);

        PptxUtils.addRect(slide, 0, ThemeConfig.SLIDE_H - 0.04, ThemeConfig.SLIDE_W, 0.04, ThemeConfig.HEX_TEAL);
    }

    // ═══════════════════════════════════════════════════════════
    // HELPERS
    // ═══════════════════════════════════════════════════════════

    private String resolveColor(String key) {
        return switch (key.toLowerCase()) {
            case "accent" -> ThemeConfig.HEX_ACCENT;
            case "teal"   -> ThemeConfig.HEX_TEAL;
            case "gold"   -> ThemeConfig.HEX_GOLD;
            case "rose"   -> ThemeConfig.HEX_ROSE;
            case "orange" -> ThemeConfig.HEX_ORANGE;
            default       -> key.length() == 6 ? key : ThemeConfig.HEX_ACCENT;
        };
    }

    private void addReferenceFooter(XSLFSlide slide, SlideData data) {
        List<Map<String, Object>> refs = data.contentList("references");
        if (refs.isEmpty()) return;

        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < Math.min(refs.size(), 6); i++) {
            if (i > 0) sb.append("  ·  ");
            sb.append("[").append(i + 1).append("] ");
            sb.append(refs.get(i).getOrDefault("text", "").toString());
        }

        PptxUtils.addText(slide, sb.toString(),
            ThemeConfig.CONTENT_X, ThemeConfig.FOOTER_Y - 0.30,
            ThemeConfig.CONTENT_W, 0.28,
            ThemeConfig.SIZE_FOOTER, ThemeConfig.FONT_MONO, ThemeConfig.HEX_DIM,
            false, TextParagraph.TextAlign.LEFT);
    }

    private String buildConfidenceBadge(SlideData data) {
        List<Map<String, Object>> refs = data.contentList("references");
        if (refs.isEmpty()) return null;
        long tier1 = refs.stream()
            .filter(r -> r.containsKey("tier") && ((Number) r.get("tier")).intValue() == 1)
            .count();
        return "Sources: " + refs.size() + " (" + tier1 + " Gold) | MedAI Suite";
    }
}
