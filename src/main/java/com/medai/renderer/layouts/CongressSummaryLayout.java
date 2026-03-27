package com.medai.renderer.layouts;

// ============================================================
// CongressSummaryLayout.java
// MedAI Suite – PPTX Renderer (Railway/Apache POI)
// Layout: CONGRESS_SUMMARY
//
// Place in: src/main/java/com/medai/pptx/layouts/
// Register in: LayoutRouter.java (see bottom of this file)
// ============================================================

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.util.Units;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.List;
import java.util.regex.*;

public class CongressSummaryLayout {

    // ── MedAI Dark Theme ──────────────────────────────────────
    private static final Color BG_DARK      = new Color(15,  22,  40);   // #0F1628
    private static final Color BG_CARD      = new Color(30,  37,  64);   // #1E2540
    private static final Color ACCENT       = new Color(124, 111, 255);  // #7C6FFF
    private static final Color TEAL         = new Color(34,  211, 165);  // #22D3A5
    private static final Color GOLD         = new Color(245, 200, 66);   // #F5C842
    private static final Color ROSE         = new Color(255, 95,  126);  // #FF5F7E
    private static final Color AMBER        = new Color(245, 158, 11);   // #F59E0B
    private static final Color TEXT_PRIMARY = new Color(232, 234, 240);  // #E8EAF0
    private static final Color TEXT_MUTED   = new Color(136, 146, 164);  // #8892A4
    private static final Color WHITE        = Color.WHITE;

    // ── Slide dimensions (16:9 widescreen) ────────────────────
    private static final int W  = 9144000;  // EMU – 10 inches
    private static final int H  = 5143500;  // EMU – 5.625 inches
    private static final int M  = 457200;   // EMU – 0.5 inch margin

    // ============================================================
    // ENTRY POINT – called by LayoutRouter
    // Input JSON structure:
    // {
    //   "meta": {
    //     "congressName": "ESMO 2026",
    //     "indication": "Oncology",
    //     "exportedAt": "2026-03-26T...",
    //     "confidenceScore": 96.5,
    //     "totalSources": 47,
    //     "peerReviewed": 44
    //   },
    //   "summary": "## ESMO 2026 – Post-Congress Summary\n### ...",
    //   "sources": [{ "id":"pm_1","title":"...","url":"...","year":"2026","source":"PubMed","type":"peer-reviewed" }]
    // }
    // ============================================================
    public static byte[] render(Map<String, Object> payload) throws Exception {
        XMLSlideShow pptx = new XMLSlideShow();
        pptx.setPageSize(new java.awt.Dimension(
            (int)(W / 914400.0 * 72),
            (int)(H / 914400.0 * 72)
        ));

        // Parse payload
        @SuppressWarnings("unchecked")
        Map<String, Object> meta = (Map<String, Object>) payload.getOrDefault("meta", new HashMap<>());
        String summary           = (String) payload.getOrDefault("summary", "");
        @SuppressWarnings("unchecked")
        List<Map<String,Object>> sources = (List<Map<String,Object>>) payload.getOrDefault("sources", new ArrayList<>());

        String congressName  = str(meta, "congressName",  "Congress Summary");
        String indication    = str(meta, "indication",    "");
        String exportedAt    = str(meta, "exportedAt",    "");
        double confScore     = dbl(meta, "confidenceScore", 0.0);
        int    totalSources  = (int) dbl(meta, "totalSources", 0);
        int    peerReviewed  = (int) dbl(meta, "peerReviewed", 0);

        // Parse summary markdown into sections
        List<Section> sections = parseSummary(summary);

        // ── Build slides ─────────────────────────────────────
        renderTitleSlide(pptx, congressName, indication, exportedAt, confScore, totalSources, peerReviewed);
        renderExecutiveSummary(pptx, sections, congressName, indication);
        for (Section sec : sections) {
            if (!sec.isExecutive && !sec.isImplications && !sec.isConfidence && !sec.isReferences) {
                renderSectionSlide(pptx, sec, congressName, indication, confScore);
            }
        }
        renderImplicationsSlide(pptx, sections, congressName, indication);
        renderSourcesSlide(pptx, sources, congressName, confScore);
        renderDisclaimerSlide(pptx, congressName, confScore);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        pptx.write(out);
        pptx.close();
        return out.toByteArray();
    }

    // ============================================================
    // SLIDE 1 – Title
    // ============================================================
    private static void renderTitleSlide(XMLSlideShow pptx, String congress, String indication,
                                          String date, double score, int total, int peer) {
        XSLFSlide slide = pptx.createSlide();
        fillBackground(slide, BG_DARK);

        // Left accent bar
        addRect(slide, 0, 0, emu(0.18), H, ACCENT);

        // MedAI logo text (top right)
        addText(slide, "MedAI Suite", W - emu(2.5), emu(0.3), emu(2.2), emu(0.4),
                TEXT_MUTED, 11, false, TextAlign.RIGHT);

        // Congress name (large)
        addText(slide, congress, emu(0.5), emu(1.2), emu(8.5), emu(1.2),
                WHITE, 40, true, TextAlign.LEFT);

        // Indication pill background + text
        int pillW = emu(2.2); int pillH = emu(0.42);
        int pillX = emu(0.5); int pillY = emu(2.6);
        addRoundRect(slide, pillX, pillY, pillW, pillH, ACCENT, 10);
        addText(slide, indication, pillX + emu(0.15), pillY + emu(0.04),
                pillW - emu(0.3), pillH - emu(0.08), WHITE, 14, true, TextAlign.CENTER);

        // Subtitle – "Post-Congress Summary"
        addText(slide, "Post-Congress Summary",
                emu(0.5), emu(3.2), emu(5), emu(0.4),
                TEXT_MUTED, 16, false, TextAlign.LEFT);

        // Date
        if (!date.isEmpty()) {
            String dateShort = date.length() > 10 ? date.substring(0, 10) : date;
            addText(slide, "Published: " + dateShort,
                    emu(0.5), emu(3.7), emu(4), emu(0.3),
                    TEXT_MUTED, 11, false, TextAlign.LEFT);
        }

        // Confidence score box (right side)
        int boxX = W - emu(3.2); int boxY = emu(1.5);
        int boxW = emu(2.8);    int boxH = emu(2.5);
        addRoundRect(slide, boxX, boxY, boxW, boxH, BG_CARD, 12);
        addRoundRect(slide, boxX, boxY, boxW, emu(0.04), scoreColor(score), 0); // top border

        addText(slide, "Liability Score", boxX + emu(0.2), boxY + emu(0.2),
                boxW - emu(0.4), emu(0.35), TEXT_MUTED, 10, true, TextAlign.CENTER);
        addText(slide, String.format("%.1f%%", score),
                boxX, boxY + emu(0.6), boxW, emu(0.9),
                scoreColor(score), 48, true, TextAlign.CENTER);
        addText(slide, score >= 95 ? "✓ Enterprise Standard" : score >= 90 ? "⚠ Review recommended" : "✕ Below threshold",
                boxX, boxY + emu(1.55), boxW, emu(0.3),
                scoreColor(score), 10, false, TextAlign.CENTER);
        addText(slide, total + " sources · " + peer + " peer-reviewed",
                boxX, boxY + emu(1.95), boxW, emu(0.3),
                TEXT_MUTED, 9, false, TextAlign.CENTER);

        // Bottom bar
        addRect(slide, 0, H - emu(0.06), W, emu(0.06), ACCENT);
    }

    // ============================================================
    // SLIDE 2 – Executive Summary
    // ============================================================
    private static void renderExecutiveSummary(XMLSlideShow pptx, List<Section> sections,
                                                String congress, String indication) {
        Section exec = sections.stream().filter(s -> s.isExecutive).findFirst()
                .orElseGet(() -> { Section s = new Section(); s.title = "Key Highlights"; s.bullets = new ArrayList<>(); return s; });

        XSLFSlide slide = pptx.createSlide();
        fillBackground(slide, BG_DARK);
        addRect(slide, 0, 0, emu(0.06), H, TEAL); // left accent – teal for executive

        // Header
        renderSlideHeader(slide, "Executive Summary", congress, indication, TEAL);

        // Bullet cards – max 6
        List<String> bullets = exec.bullets.isEmpty()
            ? List.of("No executive summary extracted – see section slides for detail.")
            : exec.bullets;

        int cols = bullets.size() <= 3 ? 1 : 2;
        int rows = (int) Math.ceil(bullets.size() / (double) cols);
        int cardW = cols == 1 ? W - emu(1.0) : (W - emu(1.2)) / 2;
        int cardH = Math.min(emu(1.0), (H - emu(1.8)) / rows);
        int startY = emu(1.4);

        for (int i = 0; i < Math.min(bullets.size(), 6); i++) {
            int col = i % cols;
            int row = i / cols;
            int x   = emu(0.5) + col * (cardW + emu(0.2));
            int y   = startY + row * (cardH + emu(0.15));

            addRoundRect(slide, x, y, cardW, cardH - emu(0.1), BG_CARD, 8);
            addRect(slide, x, y, emu(0.04), cardH - emu(0.1), TEAL); // left pill

            // Bullet number
            addText(slide, String.format("%02d", i+1),
                    x + emu(0.12), y + emu(0.05), emu(0.45), cardH - emu(0.2),
                    TEAL, 18, true, TextAlign.LEFT);

            // Bullet text
            addText(slide, bullets.get(i),
                    x + emu(0.62), y + emu(0.08), cardW - emu(0.78), cardH - emu(0.22),
                    TEXT_PRIMARY, 13, false, TextAlign.LEFT);
        }

        addRect(slide, 0, H - emu(0.06), W, emu(0.06), TEAL);
    }

    // ============================================================
    // SLIDE N – Section slide (Key Findings per section)
    // ============================================================
    private static void renderSectionSlide(XMLSlideShow pptx, Section sec,
                                            String congress, String indication, double score) {
        XSLFSlide slide = pptx.createSlide();
        fillBackground(slide, BG_DARK);
        addRect(slide, 0, 0, emu(0.06), H, ACCENT);

        renderSlideHeader(slide, sec.title, congress, indication, ACCENT);

        // Bullets as content blocks
        int y = emu(1.45);
        int usedBullets = 0;
        for (String bullet : sec.bullets) {
            if (y > H - emu(1.0)) break;
            int blockH = estimateBlockHeight(bullet);
            addRoundRect(slide, emu(0.5), y, W - emu(1.0), blockH, BG_CARD, 7);

            // Citation highlight if present
            String cleanBullet = bullet;
            addText(slide, cleanBullet,
                    emu(0.75), y + emu(0.08), W - emu(1.5), blockH - emu(0.18),
                    TEXT_PRIMARY, 13, false, TextAlign.LEFT);
            y += blockH + emu(0.12);
            usedBullets++;
        }

        // Overflow note
        if (usedBullets < sec.bullets.size()) {
            addText(slide, "+" + (sec.bullets.size()-usedBullets) + " more findings – see source report",
                    emu(0.5), y, W - emu(1.0), emu(0.3),
                    TEXT_MUTED, 10, false, TextAlign.LEFT);
        }

        // Confidence badge (bottom right)
        addText(slide, String.format("Liability Score: %.1f%%", score),
                W - emu(2.5), H - emu(0.55), emu(2.2), emu(0.3),
                scoreColor(score), 10, true, TextAlign.RIGHT);

        addRect(slide, 0, H - emu(0.06), W, emu(0.06), ACCENT);
    }

    // ============================================================
    // SLIDE – Clinical Implications
    // ============================================================
    private static void renderImplicationsSlide(XMLSlideShow pptx, List<Section> sections,
                                                 String congress, String indication) {
        Section impl = sections.stream().filter(s -> s.isImplications).findFirst()
                .orElseGet(() -> { Section s = new Section(); s.title = "Key Clinical Implications"; s.bullets = new ArrayList<>(); return s; });

        XSLFSlide slide = pptx.createSlide();
        fillBackground(slide, BG_DARK);
        addRect(slide, 0, 0, emu(0.06), H, GOLD);

        renderSlideHeader(slide, "Key Clinical Implications", congress, indication, GOLD);

        // 2-column grid of implication cards
        List<String> items = impl.bullets.isEmpty()
            ? List.of("Review source abstracts for clinical implications specific to your indication.")
            : impl.bullets;

        int cols = 2;
        int cardW = (W - emu(1.4)) / 2;
        int cardH = emu(1.35);

        String[] icons = {"💊", "🔬", "📊", "🏥", "⚡", "🎯"};

        for (int i = 0; i < Math.min(items.size(), 6); i++) {
            int col = i % cols;
            int row = i / cols;
            int x   = emu(0.5) + col * (cardW + emu(0.4));
            int y   = emu(1.45) + row * (cardH + emu(0.15));

            addRoundRect(slide, x, y, cardW, cardH, BG_CARD, 10);
            addRect(slide, x, y, cardW, emu(0.04), GOLD); // top accent

            // Icon
            addText(slide, icons[i % icons.length],
                    x + emu(0.15), y + emu(0.15), emu(0.45), emu(0.45),
                    WHITE, 20, false, TextAlign.LEFT);

            // Text
            addText(slide, items.get(i),
                    x + emu(0.15), y + emu(0.65), cardW - emu(0.3), cardH - emu(0.8),
                    TEXT_PRIMARY, 12, false, TextAlign.LEFT);
        }

        addRect(slide, 0, H - emu(0.06), W, emu(0.06), GOLD);
    }

    // ============================================================
    // SLIDE – Sources / References
    // ============================================================
    private static void renderSourcesSlide(XMLSlideShow pptx, List<Map<String,Object>> sources,
                                            String congress, double score) {
        XSLFSlide slide = pptx.createSlide();
        fillBackground(slide, BG_DARK);
        addRect(slide, 0, 0, emu(0.06), H, TEXT_MUTED);

        // Title
        addText(slide, "References & Sources", emu(0.35), emu(0.25), W - emu(0.7), emu(0.55),
                WHITE, 24, true, TextAlign.LEFT);
        addText(slide, congress + "  |  " + sources.size() + " sources",
                emu(0.35), emu(0.85), W - emu(0.7), emu(0.3),
                TEXT_MUTED, 11, false, TextAlign.LEFT);

        // Source list (2 columns, max 20)
        int cols = 2;
        int colW = (W - emu(1.2)) / 2;
        int rowH = emu(0.38);
        int startY = emu(1.3);
        int maxRows = (int)((H - emu(1.8)) / rowH);

        for (int i = 0; i < Math.min(sources.size(), maxRows * cols); i++) {
            Map<String,Object> src = sources.get(i);
            int col = i % cols;
            int row = i / cols;
            if (row >= maxRows) break;

            int x = emu(0.5) + col * (colW + emu(0.2));
            int y = startY + row * rowH;

            String title  = str(src, "title", "(No title)");
            String source = str(src, "source", "");
            String year   = str(src, "year", "");
            String type   = str(src, "type", "");

            // Number badge
            addText(slide, "[" + (i+1) + "]",
                    x, y, emu(0.4), rowH - emu(0.04),
                    ACCENT, 9, true, TextAlign.LEFT);

            // Title (truncated)
            String label = title.length() > 70 ? title.substring(0,67) + "..." : title;
            addText(slide, label,
                    x + emu(0.42), y, colW - emu(0.45), emu(0.25),
                    TEXT_PRIMARY, 9, false, TextAlign.LEFT);

            // Source meta
            Color typeColor = type.equals("preprint") ? AMBER : TEAL;
            addText(slide, source + (year.isEmpty() ? "" : " · " + year) + (type.equals("preprint") ? " ⚠ Preprint" : ""),
                    x + emu(0.42), y + emu(0.24), colW - emu(0.45), emu(0.14),
                    typeColor, 8, false, TextAlign.LEFT);
        }

        if (sources.size() > maxRows * cols) {
            addText(slide, "+" + (sources.size() - maxRows * cols) + " additional sources in full report",
                    emu(0.5), H - emu(0.6), W - emu(1.0), emu(0.25),
                    TEXT_MUTED, 9, false, TextAlign.LEFT);
        }

        addRect(slide, 0, H - emu(0.06), W, emu(0.06), TEXT_MUTED);
    }

    // ============================================================
    // SLIDE – Disclaimer / Methodology
    // ============================================================
    private static void renderDisclaimerSlide(XMLSlideShow pptx, String congress, double score) {
        XSLFSlide slide = pptx.createSlide();
        fillBackground(slide, BG_DARK);
        addRect(slide, 0, 0, emu(0.06), H, ROSE);

        addText(slide, "Disclaimer & Methodology", emu(0.35), emu(0.25), W - emu(0.7), emu(0.55),
                WHITE, 24, true, TextAlign.LEFT);

        String disclaimer =
            "Data Sources & Methodology\n\n" +
            "This congress summary was generated using MedAI Suite's Post-Congress Harvest tool. " +
            "Sources include PubMed, ClinicalTrials.gov, Europe PMC, and official congress publications. " +
            "All data points are cited with source references and have been subject to human review " +
            "prior to publication.\n\n" +
            "Liability Score: " + String.format("%.1f", score) + "% — " +
            (score >= 95 ? "Meets MedAI Enterprise Standard (≥95%)" :
             score >= 90 ? "Reviewed – minor verification recommended" :
             "Requires additional verification") + "\n\n" +
            "Important Notice\n\n" +
            "This summary is intended for Medical Affairs professionals and is for informational " +
            "purposes only. It does not constitute medical advice, clinical guidance, or regulatory " +
            "submission material. Data from preprint sources is marked ⚠ and has not been peer-reviewed. " +
            "Users are responsible for verifying all data against primary sources before use in " +
            "medical communications, publications, or regulatory documents.\n\n" +
            "© MedAI Suite · " + congress + " · Generated " + new java.util.Date().toString().substring(0,10);

        addText(slide, disclaimer, emu(0.5), emu(1.1), W - emu(1.0), H - emu(1.6),
                TEXT_MUTED, 11, false, TextAlign.LEFT);

        addRect(slide, 0, H - emu(0.06), W, emu(0.06), ROSE);
    }

    // ============================================================
    // MARKDOWN PARSER – extracts sections from harvest summary
    // ============================================================
    private static List<Section> parseSummary(String markdown) {
        List<Section> sections = new ArrayList<>();
        Section current = null;

        for (String line : markdown.split("\n")) {
            String trimmed = line.trim();

            if (trimmed.startsWith("## ") || trimmed.startsWith("# ")) {
                // Top-level heading – new section
                if (current != null) sections.add(current);
                current = new Section();
                current.title = trimmed.replaceAll("^#{1,3}\\s*", "").trim();
                current.isExecutive = current.title.toLowerCase().contains("highlight") ||
                                      current.title.toLowerCase().contains("executive") ||
                                      current.title.toLowerCase().contains("key finding");
                current.isImplications = current.title.toLowerCase().contains("implication");
                current.isConfidence   = current.title.toLowerCase().contains("confidence");
                current.isReferences   = current.title.toLowerCase().contains("reference");
                current.bullets = new ArrayList<>();

            } else if (trimmed.startsWith("### ")) {
                // Sub-heading – new section
                if (current != null) sections.add(current);
                current = new Section();
                current.title = trimmed.replaceAll("^###\\s*", "").trim();
                current.isExecutive = current.title.toLowerCase().contains("highlight");
                current.isImplications = current.title.toLowerCase().contains("implication");
                current.isConfidence   = current.title.toLowerCase().contains("confidence");
                current.isReferences   = current.title.toLowerCase().contains("reference");
                current.bullets = new ArrayList<>();

            } else if (current != null && (trimmed.startsWith("- ") || trimmed.startsWith("* "))) {
                current.bullets.add(trimmed.substring(2).trim());

            } else if (current != null && !trimmed.isEmpty() && !trimmed.startsWith("#")) {
                // Paragraph text – add as a bullet if section has no bullets yet
                if (!trimmed.startsWith("[") && trimmed.length() > 20) {
                    current.bullets.add(trimmed);
                }
            }
        }
        if (current != null) sections.add(current);

        // Ensure at least an executive section
        if (sections.stream().noneMatch(s -> s.isExecutive) && !sections.isEmpty()) {
            sections.get(0).isExecutive = true;
        }
        return sections;
    }

    // ============================================================
    // DRAWING HELPERS
    // ============================================================
    private static void fillBackground(XSLFSlide slide, Color color) {
        XSLFBackground bg = slide.getBackground();
        bg.setFillColor(color);
    }

    private static void addRect(XSLFSlide slide, int x, int y, int w, int h, Color color) {
        XSLFAutoShape rect = slide.createAutoShape();
        rect.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT);
        rect.setAnchor(new java.awt.Rectangle(x, y, w, h));
        rect.setFillColor(color);
        rect.setLineColor(color);
    }

    private static void addRoundRect(XSLFSlide slide, int x, int y, int w, int h, Color fill, int radius) {
        XSLFAutoShape rect = slide.createAutoShape();
        rect.setShapeType(org.apache.poi.sl.usermodel.ShapeType.ROUND_RECT);
        rect.setAnchor(new java.awt.Rectangle(x, y, w, h));
        rect.setFillColor(fill);
        rect.setLineColor(fill);
        // Note: corner radius set via XML manipulation if needed
    }

    private static XSLFTextBox addText(XSLFSlide slide, String text,
                                        int x, int y, int w, int h,
                                        Color color, int fontSize,
                                        boolean bold, TextAlign align) {
        XSLFTextBox tb = slide.createTextBox();
        tb.setAnchor(new java.awt.Rectangle(x, y, w, h));
        tb.clearText();
        XSLFTextParagraph para = tb.addNewTextParagraph();
        para.setTextAlign(align);
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        run.setFontColor(color);
        run.setFontSize((double) fontSize);
        run.setBold(bold);
        run.setFontFamily("Calibri");
        tb.setWordWrap(true);
        return tb;
    }

    private static void renderSlideHeader(XSLFSlide slide, String title,
                                           String congress, String indication, Color accentColor) {
        // Section title
        addText(slide, title, emu(0.35), emu(0.2), W - emu(0.7), emu(0.55),
                WHITE, 26, true, TextAlign.LEFT);

        // Congress + indication badge (top right)
        addText(slide, congress + "  ·  " + indication,
                W - emu(3.8), emu(0.25), emu(3.5), emu(0.35),
                TEXT_MUTED, 10, false, TextAlign.RIGHT);

        // Separator line
        addRect(slide, emu(0.35), emu(0.85), W - emu(0.7), emu(0.015), accentColor);
    }

    // ── EMU helper ───────────────────────────────────────────────
    private static int emu(double inches) { return (int)(inches * 914400); }

    private static Color scoreColor(double score) {
        if (score >= 95) return TEAL;
        if (score >= 90) return GOLD;
        return ROSE;
    }

    private static int estimateBlockHeight(String text) {
        int chars    = text.length();
        int lines    = Math.max(1, (int) Math.ceil(chars / 100.0));
        return emu(0.25) + lines * emu(0.22);
    }

    private static String str(Map<String,Object> m, String k, String def) {
        Object v = m.get(k);
        return v != null ? v.toString() : def;
    }

    private static double dbl(Map<String,Object> m, String k, double def) {
        Object v = m.get(k);
        if (v == null) return def;
        if (v instanceof Number) return ((Number)v).doubleValue();
        try { return Double.parseDouble(v.toString()); } catch(Exception e) { return def; }
    }

    // ── Inner data class ─────────────────────────────────────────
    static class Section {
        String       title         = "";
        List<String> bullets       = new ArrayList<>();
        boolean      isExecutive   = false;
        boolean      isImplications = false;
        boolean      isConfidence  = false;
        boolean      isReferences  = false;
    }
}

/*
 ================================================================
 REGISTRATION IN LayoutRouter.java
 ================================================================
 Add this case to your existing switch in LayoutRouter.java:

     case "CONGRESS_SUMMARY":
         return CongressSummaryLayout.render(payload);

 Full switch should look like:
     switch (layoutType) {
         case "TITLE":               return TitleLayout.render(payload);
         case "CONTENT_CARDS_4":    return ContentCards4Layout.render(payload);
         ...
         case "CONGRESS_SUMMARY":   return CongressSummaryLayout.render(payload);   // ← NEW
         default:
             throw new IllegalArgumentException("Unknown layout: " + layoutType);
     }

 ================================================================
 ENDPOINT – RenderController.java (no change needed)
 ================================================================
 Existing POST /render endpoint handles this automatically.
 The frontend sends:
 {
   "layoutType": "CONGRESS_SUMMARY",
   "meta":    { ... },
   "summary": "...",
   "sources": [ ... ]
 }
*/
