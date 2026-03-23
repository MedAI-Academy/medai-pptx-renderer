package com.medai.renderer.template;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.medai.renderer.service.PythonChartClient;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.PictureData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.awt.geom.Rectangle2D;
import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * Template-based PPTX renderer.
 *
 * Flow:
 * 1. Receives a "deck recipe" — ordered list of (layoutId, placeholders)
 * 2. Copies the master template to a temp file
 * 3. Keeps only the slides referenced in the recipe (removes the rest)
 * 4. Replaces all {{placeholder}} strings with actual content
 * 5. For CHART_IMAGE slides, fetches PNG from Python Chart Service and embeds it
 * 6. Returns the final PPTX bytes
 */
@Service
public class TemplateRenderService {

    private static final Logger log = LoggerFactory.getLogger(TemplateRenderService.class);
    private static final ObjectMapper mapper = new ObjectMapper();

    @Autowired
    private PythonChartClient pythonCharts;

    @Value("${TEMPLATE_PATH:classpath:MedAI_Template_v2_final.pptx}")
    private String templatePath;

    // Manifest: layout_id → slide_number (1-indexed)
    private Map<String, Integer> layoutSlideMap;

    /**
     * Load manifest on first use.
     */
    private synchronized Map<String, Integer> getLayoutMap() throws Exception {
        if (layoutSlideMap != null) return layoutSlideMap;

        layoutSlideMap = new LinkedHashMap<>();
        InputStream is = new ClassPathResource("template_manifest.json").getInputStream();
        Map<String, Object> manifest = mapper.readValue(is, Map.class);
        Map<String, Map<String, Object>> layouts = (Map<String, Map<String, Object>>) manifest.get("layouts");

        for (Map.Entry<String, Map<String, Object>> entry : layouts.entrySet()) {
            String layoutId = entry.getKey();
            int slideNum = ((Number) entry.getValue().get("slide_number")).intValue();
            layoutSlideMap.put(layoutId, slideNum);
        }
        log.info("Template manifest loaded: {} layouts", layoutSlideMap.size());
        return layoutSlideMap;
    }

    /**
     * Render a complete deck from a recipe.
     *
     * @param recipe Ordered list of slide specs: [{layout, placeholders, chartData?}, ...]
     * @return PPTX file as byte array
     */
    @SuppressWarnings("unchecked")
    public byte[] renderDeck(List<Map<String, Object>> recipe) throws Exception {
        Map<String, Integer> layoutMap = getLayoutMap();

        // 1. Copy template to temp file (we modify the copy)
        Path tempFile = Files.createTempFile("medai_deck_", ".pptx");
        try {
            copyTemplate(tempFile);

            try (XMLSlideShow pptx = new XMLSlideShow(new FileInputStream(tempFile.toFile()))) {

                // 2. Determine which source slides we need and in what order
                List<Integer> slideIndices = new ArrayList<>(); // 0-indexed
                for (Map<String, Object> slideSpec : recipe) {
                    String layoutId = (String) slideSpec.get("layout");
                    Integer slideNum = layoutMap.get(layoutId);
                    if (slideNum == null) {
                        throw new IllegalArgumentException("Unknown layout: " + layoutId);
                    }
                    slideIndices.add(slideNum - 1); // Convert to 0-indexed
                }

                log.info("Deck recipe: {} slides from {} layouts", slideIndices.size(), recipe.size());

                // 3. Remove all slides that are NOT in our recipe
                //    Strategy: mark slides to keep, remove the rest in reverse order
                Set<Integer> keepIndices = new HashSet<>(slideIndices);
                List<XSLFSlide> allSlides = pptx.getSlides();
                int totalSlides = allSlides.size();

                // Remove from end to start (so indices don't shift)
                for (int i = totalSlides - 1; i >= 0; i--) {
                    if (!keepIndices.contains(i)) {
                        pptx.removeSlide(i);
                    }
                }

                // 4. Now reorder slides to match recipe order
                //    After removal, we need to map original indices to new indices
                //    Build a mapping: original_index → new_index
                List<Integer> keptOriginalIndices = new ArrayList<>(keepIndices);
                Collections.sort(keptOriginalIndices);
                Map<Integer, Integer> origToNew = new HashMap<>();
                for (int newIdx = 0; newIdx < keptOriginalIndices.size(); newIdx++) {
                    origToNew.put(keptOriginalIndices.get(newIdx), newIdx);
                }

                // Reorder: build the desired order as new indices
                List<Integer> desiredNewOrder = new ArrayList<>();
                for (int origIdx : slideIndices) {
                    Integer newIdx = origToNew.get(origIdx);
                    if (newIdx != null) {
                        desiredNewOrder.add(newIdx);
                    }
                }

                // Apply reorder via setSlideOrder
                for (int targetPos = 0; targetPos < desiredNewOrder.size(); targetPos++) {
                    int currentPos = desiredNewOrder.get(targetPos);
                    if (currentPos != targetPos) {
                        pptx.setSlideOrder(pptx.getSlides().get(currentPos), targetPos);
                        // Update remaining indices after the move
                        for (int j = targetPos + 1; j < desiredNewOrder.size(); j++) {
                            int idx = desiredNewOrder.get(j);
                            if (idx >= targetPos && idx < currentPos) {
                                desiredNewOrder.set(j, idx + 1);
                            } else if (idx == currentPos) {
                                desiredNewOrder.set(j, targetPos);
                            }
                        }
                    }
                }

                // 5. Replace placeholders and handle chart images
                List<XSLFSlide> finalSlides = pptx.getSlides();
                for (int i = 0; i < Math.min(finalSlides.size(), recipe.size()); i++) {
                    XSLFSlide slide = finalSlides.get(i);
                    Map<String, Object> slideSpec = recipe.get(i);
                    String layoutId = (String) slideSpec.get("layout");
                    Map<String, String> placeholders = (Map<String, String>) slideSpec.get("placeholders");

                    if (placeholders != null) {
                        replacePlaceholders(slide, placeholders);
                    }

                    // Handle CHART_IMAGE slides — fetch PNG and embed
                    if ("CHART_IMAGE".equals(layoutId)) {
                        Map<String, Object> chartData = (Map<String, Object>) slideSpec.get("chartData");
                        String chartType = (String) slideSpec.getOrDefault("chartType", "kaplan-meier");
                        if (chartData != null) {
                            embedChart(pptx, slide, chartType, chartData);
                        }
                    }
                }

                // 6. Write output
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                pptx.write(baos);
                log.info("Deck rendered: {} slides, {} KB", finalSlides.size(), baos.size() / 1024);
                return baos.toByteArray();
            }
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    /**
     * Replace all {{placeholder}} strings in a slide's text shapes.
     */
    private void replacePlaceholders(XSLFSlide slide, Map<String, String> placeholders) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape textShape) {
                for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
                    for (XSLFTextRun run : para.getTextRuns()) {
                        String text = run.getRawText();
                        if (text != null && text.contains("{{")) {
                            for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                                String key = "{{" + entry.getKey() + "}}";
                                if (text.contains(key)) {
                                    text = text.replace(key, entry.getValue());
                                }
                            }
                            run.setText(text);
                        }
                    }
                }
            }
            // Handle grouped shapes
            else if (shape instanceof XSLFGroupShape groupShape) {
                for (XSLFShape child : groupShape.getShapes()) {
                    if (child instanceof XSLFTextShape textChild) {
                        for (XSLFTextParagraph para : textChild.getTextParagraphs()) {
                            for (XSLFTextRun run : para.getTextRuns()) {
                                String text = run.getRawText();
                                if (text != null && text.contains("{{")) {
                                    for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                                        text = text.replace("{{" + entry.getKey() + "}}", entry.getValue());
                                    }
                                    run.setText(text);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Fetch chart PNG from Python service and embed into the CHART_IMAGE slide.
     * Replaces the dashed rectangle placeholder with the actual image.
     */
    private void embedChart(XMLSlideShow pptx, XSLFSlide slide, String chartType, Map<String, Object> chartData) throws Exception {
        // Fetch PNG from Python Chart Service
        byte[] pngBytes;
        try {
            if ("kaplan-meier".equals(chartType)) {
                pngBytes = pythonCharts.fetchKaplanMeier(chartData);
            } else {
                log.warn("Chart type '{}' not yet supported in Python service, skipping", chartType);
                return;
            }
        } catch (Exception e) {
            log.error("Failed to fetch chart from Python service: {}", e.getMessage());
            return;
        }

        log.info("Received {} KB chart PNG for type '{}'", pngBytes.length / 1024, chartType);

        // Find the placeholder shape ("Chart Placeholder" rectangle)
        XSLFShape placeholderShape = null;
        for (XSLFShape shape : slide.getShapes()) {
            if ("Chart Placeholder".equals(shape.getShapeName())) {
                placeholderShape = shape;
                break;
            }
            // Also match by text content
            if (shape instanceof XSLFTextShape textShape) {
                String text = textShape.getText();
                if (text != null && text.contains("{{chart_image}}")) {
                    placeholderShape = shape;
                    break;
                }
            }
        }

        if (placeholderShape == null) {
            log.warn("No chart placeholder found on slide, embedding at default position");
            // Default position matching our template
            XSLFPictureData picData = pptx.addPicture(pngBytes, PictureData.PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(picData);
            pic.setAnchor(new Rectangle2D.Double(
                554736.0 / 914400.0 * 72,   // Convert EMU to points
                1400000.0 / 914400.0 * 72,
                11082528.0 / 914400.0 * 72,
                4500000.0 / 914400.0 * 72
            ));
            return;
        }

        // Get the position and size of the placeholder
        Rectangle2D anchor = placeholderShape.getAnchor();

        // Remove the placeholder shape
        slide.removeShape(placeholderShape);

        // Add the chart image at the same position
        XSLFPictureData picData = pptx.addPicture(pngBytes, PictureData.PictureType.PNG);
        XSLFPictureShape pic = slide.createPicture(picData);
        pic.setAnchor(anchor);

        log.info("Chart image embedded at ({}, {}) size ({} x {})",
            (int) anchor.getX(), (int) anchor.getY(),
            (int) anchor.getWidth(), (int) anchor.getHeight());
    }

    /**
     * Copy the master template to a temp file.
     */
    private void copyTemplate(Path target) throws Exception {
        InputStream is;
        if (templatePath.startsWith("classpath:")) {
            String resource = templatePath.substring("classpath:".length());
            is = new ClassPathResource(resource).getInputStream();
        } else {
            is = new FileInputStream(templatePath);
        }
        try (is) {
            Files.copy(is, target, StandardCopyOption.REPLACE_EXISTING);
        }
    }
}
