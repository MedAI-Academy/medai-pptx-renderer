package com.medai.renderer.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * HTTP client for the MedAI Python Chart Service.
 * Sends JSON payloads, receives publication-quality PNG bytes.
 *
 * Usage in PptxRenderService:
 *   byte[] png = pythonChartClient.fetchKaplanMeier(chartData);
 *   // embed png into PPTX slide via Apache POI
 */
@Component
public class PythonChartClient {

    private static final Logger log = LoggerFactory.getLogger(PythonChartClient.class);
    private static final ObjectMapper mapper = new ObjectMapper();
    private static final int TIMEOUT_MS = 30_000; // 30s for complex charts

    @Value("${CHART_SERVICE_URL:}")
    private String chartServiceUrl;

    /**
     * Check if the Python chart service is configured and reachable.
     */
    public boolean isAvailable() {
        if (chartServiceUrl == null || chartServiceUrl.isBlank()) return false;
        try {
            URL url = new URL(chartServiceUrl + "/health");
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5000);
            conn.setReadTimeout(5000);
            int code = conn.getResponseCode();
            conn.disconnect();
            return code == 200;
        } catch (Exception e) {
            log.warn("Python chart service not reachable at {}: {}", chartServiceUrl, e.getMessage());
            return false;
        }
    }

    // ═══════════════════════════════════════════════════
    // KAPLAN-MEIER
    // ═══════════════════════════════════════════════════

    /**
     * Convert the existing Java ChartService KM data format to the Python API format,
     * then fetch the PNG from the Python service.
     *
     * Java format (from frontend/Claude):
     * {
     *   "arms": [{ "label": "...", "timepoints": [...], "survival": [...], "censored": [...], "median": 12.5 }],
     *   "xlabel": "Time (months)",
     *   "ylabel": "Overall Survival",
     *   "hazardRatio": "HR 0.56 ...",
     *   "showAtRisk": true,
     *   "showMedian": true,
     *   "showCensoring": true
     * }
     *
     * Python format:
     * {
     *   "arms": [{ "label": "...", "times": [...], "events": [...] }],
     *   "title": null,
     *   "xlabel": "Time (months)",
     *   "ylabel": "Overall Survival (%)",
     *   "show_ci": true,
     *   "show_at_risk": true,
     *   "show_median": true,
     *   "show_censoring": true,
     *   "hr_text": "HR 0.56 ...",
     *   "dpi": 300
     * }
     */
    @SuppressWarnings("unchecked")
    public byte[] fetchKaplanMeier(Map<String, Object> data) throws Exception {
        List<Map<String, Object>> javaArms = (List<Map<String, Object>>) data.get("arms");
        List<Map<String, Object>> pythonArms = new ArrayList<>();

        for (int i = 0; i < javaArms.size(); i++) {
            Map<String, Object> arm = javaArms.get(i);
            Map<String, Object> pyArm = new LinkedHashMap<>();

            String label = arm.getOrDefault("label", "Arm " + (i + 1)).toString();
            pyArm.put("label", label);

            // Convert timepoints + survival step data → times + events
            // Java format uses pre-computed KM curve (timepoints, survival, censored)
            // Python format needs raw event data (times, events where 1=event, 0=censored)
            List<Number> timepoints = getNumberList(arm, "timepoints");
            List<Number> survival = getNumberList(arm, "survival");
            List<Number> censored = getNumberList(arm, "censored");

            // Check if raw times/events are provided directly (new format)
            if (arm.containsKey("times") && arm.containsKey("events")) {
                pyArm.put("times", arm.get("times"));
                pyArm.put("events", arm.get("events"));
            } else {
                // Convert step-function data to raw event data
                List<Double> times = new ArrayList<>();
                List<Integer> events = new ArrayList<>();

                // Each timepoint with a survival drop = event
                for (int j = 0; j < timepoints.size(); j++) {
                    double t = timepoints.get(j).doubleValue();
                    double sPrev = j > 0 ? survival.get(j - 1).doubleValue() : 1.0;
                    double sCurr = survival.get(j).doubleValue();
                    if (sCurr < sPrev) {
                        // Approximate number of events from survival drop
                        // This is a rough reconstruction — ideally frontend sends raw data
                        times.add(t);
                        events.add(1);
                    }
                }

                // Add censored observations
                for (Number ct : censored) {
                    times.add(ct.doubleValue());
                    events.add(0);
                }

                pyArm.put("times", times);
                pyArm.put("events", events);
            }

            // Pass color if present
            if (arm.containsKey("color")) {
                pyArm.put("color", arm.get("color").toString());
            }

            pythonArms.add(pyArm);
        }

        // Build Python API payload
        Map<String, Object> payload = new LinkedHashMap<>();
        payload.put("arms", pythonArms);
        payload.put("title", data.getOrDefault("title", null));
        payload.put("xlabel", data.getOrDefault("xlabel", "Time (months)"));
        payload.put("ylabel", data.getOrDefault("ylabel", "Overall Survival (%)"));
        payload.put("show_ci", true);
        payload.put("show_censoring", data.getOrDefault("showCensoring", true));
        payload.put("show_at_risk", data.getOrDefault("showAtRisk", true));
        payload.put("show_median", data.getOrDefault("showMedian", true));
        payload.put("hr_text", data.getOrDefault("hazardRatio", null));
        payload.put("width", 10);
        payload.put("height", 7);
        payload.put("dpi", 300);

        return postForPng("/charts/kaplan-meier", payload);
    }

    // ═══════════════════════════════════════════════════
    // FOREST PLOT (ready for when Python endpoint exists)
    // ═══════════════════════════════════════════════════

    public byte[] fetchForestPlot(Map<String, Object> data) throws Exception {
        // TODO: implement when Python /charts/forest-plot endpoint is ready
        throw new UnsupportedOperationException("Forest plot not yet available in Python service");
    }

    // ═══════════════════════════════════════════════════
    // SWIMMER PLOT (ready for when Python endpoint exists)
    // ═══════════════════════════════════════════════════

    public byte[] fetchSwimmerPlot(Map<String, Object> data) throws Exception {
        // TODO: implement when Python /charts/swimmer endpoint is ready
        throw new UnsupportedOperationException("Swimmer plot not yet available in Python service");
    }

    // ═══════════════════════════════════════════════════
    // HTTP
    // ═══════════════════════════════════════════════════

    private byte[] postForPng(String endpoint, Map<String, Object> payload) throws Exception {
        String jsonBody = mapper.writeValueAsString(payload);
        log.info("Calling Python chart service: {} ({} bytes)", endpoint, jsonBody.length());

        URL url = new URL(chartServiceUrl + endpoint);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("POST");
        conn.setRequestProperty("Content-Type", "application/json");
        conn.setRequestProperty("Accept", "image/png");
        conn.setConnectTimeout(TIMEOUT_MS);
        conn.setReadTimeout(TIMEOUT_MS);
        conn.setDoOutput(true);

        try (OutputStream os = conn.getOutputStream()) {
            os.write(jsonBody.getBytes(StandardCharsets.UTF_8));
        }

        int responseCode = conn.getResponseCode();
        if (responseCode != 200) {
            String error = "";
            try (InputStream es = conn.getErrorStream()) {
                if (es != null) error = new String(es.readAllBytes(), StandardCharsets.UTF_8);
            }
            throw new RuntimeException("Chart service returned " + responseCode + ": " + error);
        }

        try (InputStream is = conn.getInputStream()) {
            byte[] png = is.readAllBytes();
            log.info("Received chart PNG: {} KB", png.length / 1024);
            return png;
        } finally {
            conn.disconnect();
        }
    }

    @SuppressWarnings("unchecked")
    private List<Number> getNumberList(Map<String, Object> m, String key) {
        Object v = m.get(key);
        return v instanceof List ? (List<Number>) v : List.of();
    }
}
