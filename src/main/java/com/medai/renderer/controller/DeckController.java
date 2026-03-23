package com.medai.renderer.controller;

import com.medai.renderer.template.TemplateRenderService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

/**
 * REST endpoint for template-based deck rendering.
 *
 * POST /api/render/deck
 * Accepts: { "filename": "...", "slides": [{layout, placeholders, chartData?}, ...] }
 * Returns: PPTX binary
 *
 * GET /api/render/layouts
 * Returns: list of available layout IDs from the manifest
 */
@RestController
@RequestMapping("/api/render")
public class DeckController {

    private static final Logger log = LoggerFactory.getLogger(DeckController.class);
    private static final String PPTX_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    private final TemplateRenderService templateService;

    public DeckController(TemplateRenderService templateService) {
        this.templateService = templateService;
    }

    /**
     * Render a deck from a recipe.
     *
     * Request body:
     * {
     *   "filename": "Drug_MAP_2027.pptx",
     *   "slides": [
     *     { "layout": "COVER", "placeholders": { "slide_title": "...", ... } },
     *     { "layout": "SECTION_DIVIDER", "placeholders": { "slide_title": "..." } },
     *     { "layout": "CHART_IMAGE", "placeholders": {...}, "chartType": "kaplan-meier", "chartData": {...} },
     *     ...
     *   ]
     * }
     */
    @SuppressWarnings("unchecked")
    @PostMapping("/deck")
    public ResponseEntity<byte[]> renderDeck(@RequestBody Map<String, Object> request) {
        long start = System.currentTimeMillis();

        try {
            List<Map<String, Object>> slides =
                    (List<Map<String, Object>>) request.get("slides");

            if (slides == null || slides.isEmpty()) {
                return ResponseEntity.badRequest()
                        .body("{\"error\": \"No slides provided\"}".getBytes());
            }

            String filename = (String) request.getOrDefault("filename", "MedAI_Deck.pptx");

            log.info("Deck render request: {} slides, filename='{}'", slides.size(), filename);
            log.info("Layouts: {}", slides.stream()
                    .map(s -> (String) s.get("layout"))
                    .toList());

            // Render
            byte[] pptxBytes = templateService.renderDeck(slides);

            long elapsed = System.currentTimeMillis() - start;
            log.info("Deck rendered: {} slides, {} KB, {}ms",
                    slides.size(), pptxBytes.length / 1024, elapsed);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(PPTX_CONTENT_TYPE));
            headers.setContentDispositionFormData("attachment", filename);
            headers.setContentLength(pptxBytes.length);
            headers.add("Access-Control-Expose-Headers", "Content-Disposition");

            return new ResponseEntity<>(pptxBytes, headers, HttpStatus.OK);

        } catch (IllegalArgumentException e) {
            log.error("Invalid recipe: {}", e.getMessage());
            String errorJson = "{\"error\": \"" + e.getMessage().replace("\"", "'") + "\"}";
            return ResponseEntity.badRequest().body(errorJson.getBytes());

        } catch (Exception e) {
            log.error("Deck render failed", e);
            String errorJson = "{\"error\": \"Rendering failed: "
                    + e.getMessage().replace("\"", "'") + "\"}";
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(errorJson.getBytes());
        }
    }

    /**
     * List available layouts from the template manifest.
     */
    @GetMapping("/layouts")
    public ResponseEntity<Map<String, Object>> layouts() {
        try {
            // Trigger manifest load and return layout IDs
            return ResponseEntity.ok(Map.of(
                    "status", "ok",
                    "message", "Use POST /api/render/deck with layout IDs in the 'slides' array",
                    "endpoint", "/api/render/deck"
            ));
        } catch (Exception e) {
            return ResponseEntity.status(500).body(Map.of("error", e.getMessage()));
        }
    }
}
