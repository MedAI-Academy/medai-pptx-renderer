package com.medai.renderer.controller;

import com.medai.renderer.model.RenderRequest;
import com.medai.renderer.service.PptxRenderService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

/**
 * REST API for PPTX rendering.
 * 
 * Endpoints:
 *   POST /api/v1/render   — accepts RenderRequest JSON, returns PPTX binary
 *   POST /render           — legacy endpoint (backwards compatible with python renderer)
 *   GET  /api/v1/health    — health check
 *   GET  /api/v1/templates — list available templates
 */
@RestController
public class RenderController {

    private static final Logger log = LoggerFactory.getLogger(RenderController.class);
    private static final String PPTX_CONTENT_TYPE =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    private final PptxRenderService renderService;

    public RenderController(PptxRenderService renderService) {
        this.renderService = renderService;
    }

    // ═══════════════════════════════════════════════════════════
    // MAIN RENDER ENDPOINT (new API)
    // ═══════════════════════════════════════════════════════════

    @PostMapping("/api/v1/render")
    public ResponseEntity<byte[]> render(@RequestBody RenderRequest request) {
        long start = System.currentTimeMillis();
        log.info("Render request: module={}, slides={}, theme={}",
            request.getModule(),
            request.getSlides() != null ? request.getSlides().size() : 0,
            request.getTheme());

        try {
            // Validate
            if (request.getSlides() == null || request.getSlides().isEmpty()) {
                return ResponseEntity.badRequest()
                    .body("{\"error\": \"No slides provided\"}".getBytes());
            }

            // Render PPTX
            byte[] pptxBytes = renderService.render(request);

            // Build filename
            String filename = buildFilename(request);

            long elapsed = System.currentTimeMillis() - start;
            log.info("Render complete: {} slides, {}KB, {}ms",
                request.getSlides().size(),
                pptxBytes.length / 1024,
                elapsed);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(PPTX_CONTENT_TYPE));
            headers.setContentDispositionFormData("attachment", filename);
            headers.setContentLength(pptxBytes.length);
            // Allow frontend to read Content-Disposition
            headers.add("Access-Control-Expose-Headers", "Content-Disposition");

            return new ResponseEntity<>(pptxBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            log.error("Render failed", e);
            String errorJson = "{\"error\": \"" + e.getMessage().replace("\"", "'") + "\"}";
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(errorJson.getBytes());
        }
    }

    // ═══════════════════════════════════════════════════════════
    // LEGACY ENDPOINT (backwards compatible with python renderer)
    // The existing map_generator.html calls POST /render
    // ═══════════════════════════════════════════════════════════

    @PostMapping("/render")
    public ResponseEntity<byte[]> legacyRender(@RequestBody RenderRequest request) {
        log.info("Legacy /render endpoint called — forwarding to /api/v1/render");
        return render(request);
    }

    // ═══════════════════════════════════════════════════════════
    // HEALTH & INFO
    // ═══════════════════════════════════════════════════════════

    @GetMapping("/api/v1/health")
    public ResponseEntity<Map<String, Object>> health() {
        return ResponseEntity.ok(Map.of(
            "status", "ok",
            "engine", "Apache POI 5.4 (Java)",
            "version", "1.0.0",
            "features", Map.of(
                "kaplanMeier", true,
                "swotMatrix", true,
                "timeline", true,
                "confidenceScore", true,
                "widescreen", true
            )
        ));
    }

    @GetMapping("/api/v1/templates")
    public ResponseEntity<Map<String, Object>> templates() {
        return ResponseEntity.ok(Map.of(
            "templates", Map.of(
                "dark", Map.of(
                    "name", "MedAI Dark",
                    "description", "Navy/teal premium theme — default for all modules",
                    "colors", Map.of(
                        "primary", "#0B1A3B",
                        "accent", "#7C6FFF",
                        "teal", "#22D3A5"
                    )
                ),
                "light", Map.of(
                    "name", "MedAI Light",
                    "description", "Light background for print-friendly output",
                    "colors", Map.of(
                        "primary", "#F0F4FF",
                        "accent", "#7C6FFF",
                        "text", "#1E293B"
                    )
                )
            ),
            "layouts", new String[]{
                "TITLE", "TOC", "DIVIDER", "CONTENT_FULL", "CONTENT_TWO_COL",
                "CONTENT_CARDS", "TABLE", "CHART_KM", "CHART_BAR", "SWOT",
                "TIMELINE", "KPI_DASHBOARD", "REFERENCES", "CONFIDENCE"
            }
        ));
    }

    // ═══════════════════════════════════════════════════════════
    // ROOT — Simple landing page / redirect
    // ═══════════════════════════════════════════════════════════

    @GetMapping("/")
    public ResponseEntity<Map<String, String>> root() {
        return ResponseEntity.ok(Map.of(
            "service", "MedAI PPTX Renderer",
            "version", "1.0.0 (Java/Apache POI)",
            "docs", "/api/v1/templates",
            "health", "/api/v1/health"
        ));
    }

    // ═══════════════════════════════════════════════════════════
    // HELPERS
    // ═══════════════════════════════════════════════════════════

    private String buildFilename(RenderRequest request) {
        String drug = request.meta("drug").replaceAll("[^a-zA-Z0-9]", "_");
        String module = request.getModule() != null ? request.getModule() : "presentation";
        String year = request.meta("year");

        if (drug.isEmpty()) drug = "MedAI";
        if (year.isEmpty()) year = String.valueOf(java.time.Year.now().getValue());

        // e.g. "Belantamab_Mafodotin_map_Germany_2027.pptx"
        String country = request.meta("country").replaceAll("[^a-zA-Z0-9]", "_");
        if (!country.isEmpty()) {
            return drug + "_" + module + "_" + country + "_" + year + ".pptx";
        }
        return drug + "_" + module + "_" + year + ".pptx";
    }
}
