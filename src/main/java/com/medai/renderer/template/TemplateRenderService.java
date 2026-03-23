package com.medai.renderer.template;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.medai.renderer.service.PythonChartClient;
import org.apache.poi.openxml4j.util.ZipSecureFile;
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
 * Template-based PPTX renderer — CLONE strategy.
 *
 * Flow:
 * 1. Receives a "deck recipe" — ordered list of (layoutId, placeholders)
 * 2. Opens the master template READ-ONLY as source
 * 3. Creates a NEW empty PPTX as output
 * 4. For each recipe entry: imports (clones) the source slide into the output
 *    → Duplicates are naturally handled (each clone is independent)
 * 5. Replaces all {{placeholder}} strings with actual content
 * 6. For CHART_IMAGE slides, fetches PNG from Python Chart Service and embeds it
 * 7. Returns the final PPTX bytes
 *
 * KEY FIX: The old "keep & reorder" strategy used a Set<Integer> for slide indices,
 * which collapsed duplicates. If a recipe used SECTION_DIVIDER 3 times (all pointing
 * to template slide 5), the Set only kept one copy, then the reorder logic tried to
 * access indices beyond the deck size → IndexOutOfBoundsException.
 *
 * The new CLONE strategy creates a fresh deck and imports each slide individually,
 * so using the same layout 10 times creates 10 independent slide copies.
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
     * Load manifest on first use (thread-safe lazy init).
     */
    private synchronized Map<String, Integer> getLayoutMap() throws Exception {
        if (layoutSlideMap != null) return layoutSlideMap;

        layoutSlideMap = new LinkedHashMap<>();
        InputStream is = new ClassPathResource("template_manifest.json").getInputStream();
        Map<String, Object> manifest = mapper.readValue(is, Map.class);
        Map<String, Map<String, Object>> layouts =
                (Map<String, Map<String, Object>>) manifest.get("layouts");

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

        // Increase zip entry limit — template has 216 slides = 1500+ internal XML files
        ZipSecureFile.setMaxFileCount(5000);

        // 1. Resolve all layout IDs to 0-indexed template slide numbers FIRST
        //    (fail fast if any layout is unknown)
        List<Integer> sourceIndices = new ArrayList<>();
        for (Map<String, Object> slideSpec : recipe) {
            String layoutId = (String) slideSpec.get("layout");
            Integer slideNum = layoutMap.get(layoutId);
            if (slideNum == null) {
                throw new IllegalArgumentException("Unknown layout: " + layoutId
                        + ". Available: " + layoutMap.keySet());
            }
            sourceIndices.add(slideNum - 1); // Convert to 0-indexed
        }

        log.info("Deck recipe: {} slides, layouts: {}",
                recipe.size(),
                recipe.stream().map(s -> (String) s.get("layout")).toList());

        // 2. Open template as READ-ONLY source
        Path tempSource = Files.createTempFile("medai_src_", ".pptx");
        try {
            copyTemplate(tempSource);

            try (XMLSlideShow source = new XMLSlideShow(new FileInputStream(tempSource.toFile()))) {

                List<XSLFSlide> sourceSlides = source.getSlides();
                int sourceTotal = sourceSlides.size();
                log.info("Template loaded: {} source slides", sourceTotal);

                // Validate all indices are in range
                for (int i = 0; i < sourceIndices.size(); i++) {
                    int idx = sourceIndices.get(i);
                    if (idx < 0 || idx >= sourceTotal) {
                        String layoutId = (String) recipe.get(i).get("layout");
                        throw new IllegalArgumentException(
                                "Layout '" + layoutId + "' maps to slide " + (idx + 1)
                                        + " but template only has " + sourceTotal + " slides");
                    }
                }

                // 3. Create NEW empty output presentation
                //    Copy slide dimensions from source
                XMLSlideShow output = new XMLSlideShow();
                output.setPageSize(source.getPageSize());

                // 4. Clone each recipe slide from source → output
                //    EACH clone is independent, so duplicates work naturally
                for (int i = 0; i < recipe.size(); i++) {
                    int srcIdx = sourceIndices.get(i);
                    XSLFSlide srcSlide = sourceSlides.get(srcIdx);

                    // Import slide from source into output
                    // This deep-copies all shapes, backgrounds, layouts, masters
                    XSLFSlide cloned = output.createSlide();
                    cloneSlideContent(srcSlide, cloned, source, output);

                    log.debug("Cloned template slide {} → output slide {} (layout: {})",
                            srcIdx + 1, i + 1, recipe.get(i).get("layout"));
                }

                // 5. Replace placeholders and handle chart images
                //    Now we ONLY work with output indices (0 to recipe.size()-1)
                List<XSLFSlide> outputSlides = output.getSlides();
                for (int i = 0; i < outputSlides.size(); i++) {
                    XSLFSlide slide = outputSlides.get(i);
                    Map<String, Object> slideSpec = recipe.get(i);
                    String layoutId = (String) slideSpec.get("layout");

                    // Replace {{placeholders}}
                    Map<String, String> placeholders =
                            (Map<String, String>) slideSpec.get("placeholders");
                    if (placeholders != null) {
                        replacePlaceholders(slide, placeholders);
                    }

                    // Handle CHART_IMAGE slides — fetch PNG and embed
                    if ("CHART_IMAGE".equals(layoutId)) {
                        Map<String, Object> chartData =
                                (Map<String, Object>) slideSpec.get("chartData");
                        String chartType = (String) slideSpec.getOrDefault("chartType", "kaplan-meier");
                        if (chartData != null) {
                            embedChart(output, slide, chartType, chartData);
                        }
                    }
                }

                // 6. Write output
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                output.write(baos);
                output.close();

                log.info("Deck rendered: {} slides, {} KB",
                        outputSlides.size(), baos.size() / 1024);
                return baos.toByteArray();
            }
        } finally {
            Files.deleteIfExists(tempSource);
        }
    }

    /**
     * Deep-clone all content from a source slide to a target slide.
     *
     * Uses XSLFSlide.importContent() which copies all shapes, text,
     * images, background, and formatting. This is the Apache POI
     * equivalent of "duplicate slide" in PowerPoint.
     */
    private void cloneSlideContent(XSLFSlide src, XSLFSlide dst,
                                    XMLSlideShow srcPptx, XMLSlideShow dstPptx) {
        try {
            // importContent copies shapes, text runs, images, fills, etc.
            dst.importContent(src);
        } catch (Exception e) {
            log.warn("importContent failed for slide, falling back to shape-by-shape copy: {}",
                    e.getMessage());
            // Fallback: copy shapes individually
            copyShapesManually(src, dst, srcPptx, dstPptx);
        }

        // Copy background if set on the slide directly
        try {
            if (src.getBackground() != null) {
                // Background is typically inherited from layout/master
                // importContent should handle it, but we verify
                copyBackground(src, dst);
            }
        } catch (Exception e) {
            log.debug("Background copy skipped: {}", e.getMessage());
        }
    }

    /**
     * Fallback shape-by-shape copy if importContent fails.
     */
    private void copyShapesManually(XSLFSlide src, XSLFSlide dst,
                                     XMLSlideShow srcPptx, XMLSlideShow dstPptx) {
        for (XSLFShape shape : src.getShapes()) {
            try {
                if (shape instanceof XSLFTextShape textShape) {
                    // Create a matching text box
                    XSLFTextBox copy = dst.createTextBox();
                    copy.setAnchor(textShape.getAnchor());

                    // Copy text content preserving runs
                    copy.clearText();
                    for (XSLFTextParagraph srcPara : textShape.getTextParagraphs()) {
                        XSLFTextParagraph dstPara = copy.addNewTextParagraph();
                        dstPara.setTextAlign(srcPara.getTextAlign());
                        for (XSLFTextRun srcRun : srcPara.getTextRuns()) {
                            XSLFTextRun dstRun = dstPara.addNewTextRun();
                            dstRun.setText(srcRun.getRawText());
                            try {
                                dstRun.setFontSize(srcRun.getFontSize());
                                dstRun.setBold(srcRun.isBold());
                                dstRun.setItalic(srcRun.isItalic());
                                dstRun.setFontColor(srcRun.getFontColor());
                                dstRun.setFontFamily(srcRun.getFontFamily());
                            } catch (Exception ignored) {
                                // Some properties may not be set
                            }
                        }
                    }

                } else if (shape instanceof XSLFPictureShape picShape) {
                    // Re-add picture data and create picture
                    XSLFPictureData srcPicData = picShape.getPictureData();
                    if (srcPicData != null) {
                        XSLFPictureData dstPicData = dstPptx.addPicture(
                                srcPicData.getData(), srcPicData.getType());
                        XSLFPictureShape copy = dst.createPicture(dstPicData);
                        copy.setAnchor(picShape.getAnchor());
                    }

                } else if (shape instanceof XSLFAutoShape autoShape) {
                    // Auto shapes (rectangles, etc.)
                    XSLFAutoShape copy = dst.createAutoShape();
                    copy.setAnchor(autoShape.getAnchor());
                    copy.setShapeType(autoShape.getShapeType());
                    // Copy text if any
                    if (autoShape.getText() != null && !autoShape.getText().isEmpty()) {
                        copy.setText(autoShape.getText());
                    }
                }
                // Note: XSLFGroupShape, XSLFTable, XSLFConnectorShape 
                // are handled by importContent in the primary path
            } catch (Exception e) {
                log.debug("Shape copy skipped ({}): {}", shape.getClass().getSimpleName(), e.getMessage());
            }
        }
    }

    /**
     * Copy slide background.
     */
    private void copyBackground(XSLFSlide src, XSLFSlide dst) {
        try {
            // Use OOXML direct access for background
            var srcBg = src.getXmlObject().getCSld().getBg();
            if (srcBg != null) {
                var dstCsld = dst.getXmlObject().getCSld();
                if (dstCsld.getBg() == null) {
                    dstCsld.addNewBg();
                }
                dstCsld.getBg().set(srcBg.copy());
            }
        } catch (Exception e) {
            log.debug("OOXML background copy failed: {}", e.getMessage());
        }
    }

    /**
     * Replace all {{placeholder}} strings in a slide's text shapes.
     * Handles: direct shapes, grouped shapes, table cells.
     */
    private void replacePlaceholders(XSLFSlide slide, Map<String, String> placeholders) {
        for (XSLFShape shape : slide.getShapes()) {
            replaceInShape(shape, placeholders);
        }
    }

    /**
     * Recursively replace placeholders in any shape type.
     */
    private void replaceInShape(XSLFShape shape, Map<String, String> placeholders) {
        if (shape instanceof XSLFTextShape textShape) {
            replaceInTextShape(textShape, placeholders);

        } else if (shape instanceof XSLFGroupShape groupShape) {
            for (XSLFShape child : groupShape.getShapes()) {
                replaceInShape(child, placeholders);
            }

        } else if (shape instanceof XSLFTable table) {
            // Handle table cells
            for (int row = 0; row < table.getNumberOfRows(); row++) {
                for (int col = 0; col < table.getNumberOfColumns(); col++) {
                    XSLFTableCell cell = table.getCell(row, col);
                    if (cell != null) {
                        replaceInTextShape(cell, placeholders);
                    }
                }
            }
        }
    }

    /**
     * Replace {{placeholder}} in any text shape (textbox, autoshape, table cell).
     * Preserves original formatting of each run.
     */
    private void replaceInTextShape(XSLFTextShape textShape, Map<String, String> placeholders) {
        for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                String text = run.getRawText();
                if (text != null && text.contains("{{")) {
                    for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                        String key = "{{" + entry.getKey() + "}}";
                        String value = entry.getValue() != null ? entry.getValue() : "";
                        if (text.contains(key)) {
                            text = text.replace(key, value);
                        }
                    }
                    run.setText(text);
                }
            }
        }
    }

    /**
     * Fetch chart PNG from Python service and embed into the CHART_IMAGE slide.
     * Replaces the placeholder shape with the actual chart image.
     */
    private void embedChart(XMLSlideShow pptx, XSLFSlide slide,
                            String chartType, Map<String, Object> chartData) throws Exception {
        // Fetch PNG from Python Chart Service
        byte[] pngBytes;
        try {
            if ("kaplan-meier".equals(chartType)) {
                pngBytes = pythonCharts.fetchKaplanMeier(chartData);
            } else {
                log.warn("Chart type '{}' not yet supported, skipping", chartType);
                return;
            }
        } catch (Exception e) {
            log.error("Failed to fetch chart from Python service: {}", e.getMessage());
            return; // Don't crash — just leave placeholder
        }

        log.info("Received {} KB chart PNG for type '{}'", pngBytes.length / 1024, chartType);

        // Find the placeholder shape: by name or by {{chart_image}} text
        XSLFShape placeholderShape = null;
        for (XSLFShape shape : slide.getShapes()) {
            String name = shape.getShapeName();
            if (name != null && name.contains("Chart Placeholder")) {
                placeholderShape = shape;
                break;
            }
            if (shape instanceof XSLFTextShape textShape) {
                String text = textShape.getText();
                if (text != null && text.contains("{{chart_image}}")) {
                    placeholderShape = shape;
                    break;
                }
            }
        }

        // Get position (from placeholder or default)
        Rectangle2D anchor;
        if (placeholderShape != null) {
            anchor = placeholderShape.getAnchor();
            slide.removeShape(placeholderShape);
        } else {
            log.warn("No chart placeholder found, using default position");
            // Default: centered content area (in points)
            anchor = new Rectangle2D.Double(48, 108, 864, 360);
        }

        // Add chart image at the placeholder's position
        XSLFPictureData picData = pptx.addPicture(pngBytes, PictureData.PictureType.PNG);
        XSLFPictureShape pic = slide.createPicture(picData);
        pic.setAnchor(anchor);

        log.info("Chart embedded at ({}, {}) size ({} x {})",
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
