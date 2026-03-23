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

                // 3. CLONE STRATEGY (POI 5.3.0 compatible):
                //    Instead of creating a new empty deck (which loses masters/layouts),
                //    we work on the template copy itself:
                //    a) First, DUPLICATE the slides we need (appending clones at the end)
                //    b) Then remove ALL original 216 template slides
                //    c) Result: only our recipe slides remain, in order
                //
                //    This preserves slide masters, layouts, backgrounds, and all formatting
                //    because cloned slides inherit from the same masters.

                int originalCount = sourceSlides.size(); // 216
                log.info("Cloning {} recipe slides from {} template slides",
                        recipe.size(), originalCount);

                // 4a. Append clones for each recipe entry
                //     After this loop: slides [0..215] = original, [216..216+N-1] = our clones
                for (int i = 0; i < recipe.size(); i++) {
                    int srcIdx = sourceIndices.get(i);
                    XSLFSlide srcSlide = sourceSlides.get(srcIdx);

                    // Clone by importing the slide's XML — works in POI 5.3.0
                    XSLFSlide cloned = source.createSlide(srcSlide.getSlideLayout());
                    cloneSlideViaXml(srcSlide, cloned);

                    log.debug("Cloned template slide {} → position {} (layout: {})",
                            srcIdx + 1, originalCount + i + 1, recipe.get(i).get("layout"));
                }

                // 4b. Remove ALL original template slides (indices 0 to originalCount-1)
                //     Remove from end to start so indices don't shift
                for (int i = originalCount - 1; i >= 0; i--) {
                    source.removeSlide(i);
                }

                // 5. Replace placeholders and handle chart images
                //    Now we ONLY work with the remaining slides (our clones)
                List<XSLFSlide> outputSlides = source.getSlides();
                log.info("After cleanup: {} slides remain (expected {})",
                        outputSlides.size(), recipe.size());

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
                            embedChart(source, slide, chartType, chartData);
                        }
                    }
                }

                // 6. Write output
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                source.write(baos);

                log.info("Deck rendered: {} slides, {} KB",
                        outputSlides.size(), baos.size() / 1024);
                return baos.toByteArray();
            }
        } finally {
            Files.deleteIfExists(tempSource);
        }
    }

    /**
     * Clone slide content via XML node copy (POI 5.3.0 compatible).
     *
     * Copies all shapes, text, images, fills, and formatting from src to dst
     * by replacing the dst slide's XML content with a copy of src's XML.
     * Since both slides live in the same presentation, all references
     * (images, masters, layouts) resolve correctly.
     */
    private void cloneSlideViaXml(XSLFSlide src, XSLFSlide dst) {
        try {
            // Copy the common slide data (cSld) which contains all shapes
            var srcCsld = src.getXmlObject().getCSld();
            var dstXml = dst.getXmlObject();

            // Replace the shape tree (spTree) — this is where all content lives
            if (srcCsld.getSpTree() != null) {
                dstXml.getCSld().setSpTree(
                        (org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape)
                                srcCsld.getSpTree().copy()
                );
            }

            // Copy background if explicitly set on the slide
            if (srcCsld.getBg() != null) {
                dstXml.getCSld().setBg(srcCsld.getBg().copy());
            }

            // Copy slide-level color map override if present
            if (src.getXmlObject().getClrMapOvr() != null) {
                dstXml.setClrMapOvr(src.getXmlObject().getClrMapOvr().copy());
            }

        } catch (Exception e) {
            log.warn("XML clone failed, falling back to shape-by-shape: {}", e.getMessage());
            // Fallback: copy text shapes manually (at least placeholders will work)
            copyShapesManually(src, dst);
        }
    }

    /**
     * Fallback shape copy — copies text content for placeholder replacement.
     */
    private void copyShapesManually(XSLFSlide src, XSLFSlide dst) {
        for (XSLFShape shape : src.getShapes()) {
            try {
                if (shape instanceof XSLFTextShape textShape) {
                    XSLFTextBox copy = dst.createTextBox();
                    copy.setAnchor(textShape.getAnchor());
                    copy.clearText();
                    for (XSLFTextParagraph srcPara : textShape.getTextParagraphs()) {
                        XSLFTextParagraph dstPara = copy.addNewTextParagraph();
                        try { dstPara.setTextAlign(srcPara.getTextAlign()); } catch (Exception ignored) {}
                        for (XSLFTextRun srcRun : srcPara.getTextRuns()) {
                            XSLFTextRun dstRun = dstPara.addNewTextRun();
                            dstRun.setText(srcRun.getRawText());
                            try {
                                dstRun.setFontSize(srcRun.getFontSize());
                                dstRun.setBold(srcRun.isBold());
                                dstRun.setItalic(srcRun.isItalic());
                                dstRun.setFontColor(srcRun.getFontColor());
                                dstRun.setFontFamily(srcRun.getFontFamily());
                            } catch (Exception ignored) {}
                        }
                    }
                }
            } catch (Exception e) {
                log.debug("Shape copy skipped: {}", e.getMessage());
            }
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
