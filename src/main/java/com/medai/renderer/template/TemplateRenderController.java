package com.medai.renderer.template;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

/**
 * REST endpoint for template-based PPTX generation.
 *
 * POST /api/render/deck
 * Body: { "filename": "MAP_NSCLC.pptx", "slides": [ ... ] }
 *
 * Each slide: {
 *   "layout": "SWOT_V1",
 *   "placeholders": { "slide_title": "SWOT: Keytruda", "strengths_bullets": "..." },
 *   "chartType": "kaplan-meier",       // only for CHART_IMAGE
 *   "chartData": { "arms": [...] }     // only for CHART_IMAGE
 * }
 */
@RestController
@RequestMapping("/api/render")
public class TemplateRenderController {

    private static final Logger log = LoggerFactory.getLogger(TemplateRenderController.class);

    @Autowired
    private TemplateRenderService templateService;

    @SuppressWarnings("unchecked")
    @PostMapping("/deck")
    public ResponseEntity<byte[]> renderDeck(@RequestBody Map<String, Object> request) {
        try {
            String filename = (String) request.getOrDefault("filename", "MedAI_Deck.pptx");
            List<Map<String, Object>> slides = (List<Map<String, Object>>) request.get("slides");

            if (slides == null || slides.isEmpty()) {
                return ResponseEntity.badRequest()
                    .body("Missing 'slides' array in request".getBytes());
            }

            log.info("Rendering deck '{}' with {} slides", filename, slides.size());

            byte[] pptxBytes = templateService.renderDeck(slides);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"));
            headers.setContentDisposition(
                ContentDisposition.builder("attachment").filename(filename).build());
            headers.setContentLength(pptxBytes.length);

            return new ResponseEntity<>(pptxBytes, headers, HttpStatus.OK);

        } catch (IllegalArgumentException e) {
            log.error("Bad request: {}", e.getMessage());
            return ResponseEntity.badRequest().body(e.getMessage().getBytes());
        } catch (Exception e) {
            log.error("Deck rendering failed: {}", e.getMessage(), e);
            return ResponseEntity.internalServerError()
                .body(("Rendering failed: " + e.getMessage()).getBytes());
        }
    }

    /**
     * GET /api/render/layouts — returns available layouts for frontend UI
     */
    @GetMapping("/layouts")
    public ResponseEntity<Map<String, Object>> getLayouts() {
        try {
            // Read manifest and return layout groups + descriptions
            var is = new org.springframework.core.io.ClassPathResource("template_manifest.json").getInputStream();
            var manifest = new com.fasterxml.jackson.databind.ObjectMapper().readValue(is, Map.class);
            return ResponseEntity.ok(manifest);
        } catch (Exception e) {
            return ResponseEntity.internalServerError().build();
        }
    }
}
