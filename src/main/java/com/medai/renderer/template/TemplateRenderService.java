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
 * Template-based PPTX renderer — CLONE strategy (POI 5.3.0 compatible).
 *
 * KEY FIX: The old "keep & reorder" strategy used a Set which collapsed
 * duplicate layouts (e.g. 3x SECTION_DIVIDER all pointing to same template slide).
 * New strategy: clone each recipe slide individually, then delete all originals.
 */
@Service
public class TemplateRenderService {

    private static final Logger log = LoggerFactory.getLogger(TemplateRenderService.class);
    private static final ObjectMapper mapper = new ObjectMapper();

    @Autowired
    private PythonChartClient pythonCharts;

    @Value("${TEMPLATE_PATH:classpath:MedAI_Template_v2_final.pptx}")
    private String templatePath;

    private Map<String, Integer> layoutSlideMap;

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

    @SuppressWarnings("unchecked")
    public byte[] renderDeck(List<Map<String, Object>> recipe) throws Exception {
        Map<String, Integer> layoutMap = getLayoutMap();

        ZipSecureFile.setMaxFileCount(5000);

        // 1. Resolve layout IDs to 0-indexed template slide numbers
        List<Integer> sourceIndices = new ArrayList<>();
        for (Map<String, Object> slideSpec : recipe) {
            String layoutId = (String) slideSpec.get("layout");
            Integer slideNum = layoutMap.get(layoutId);
            if (slideNum == null) {
                throw new IllegalArgumentException("Unknown layout: " + layoutId
                        + ". Available: " + layoutMap.keySet());
            }
            sourceIndices.add(slideNum - 1);
        }

        log.info("Deck recipe: {} slides, layouts: {}",
                recipe.size(),
                recipe.stream().map(s -> (String) s.get("layout")).toList());

        // 2. Open template copy
        Path tempSource = Files.createTempFile("medai_src_", ".pptx");
        try {
            copyTemplate(tempSource);

            try (XMLSlideShow source = new XMLSlideShow(new FileInputStream(tempSource.toFile()))) {

                List<XSLFSlide> sourceSlides = source.getSlides();
                int sourceTotal = sourceSlides.size();
                log.info("Template loaded: {} source slides", sourceTotal);

                // Validate indices
                for (int i = 0; i < sourceIndices.size(); i++) {
                    int idx = sourceIndices.get(i);
                    if (idx < 0 || idx >= sourceTotal) {
                        String layoutId = (String) recipe.get(i).get("layout");
                        throw new IllegalArgumentException(
                                "Layout '" + layoutId + "' maps to slide " + (idx + 1)
                                        + " but template only has " + sourceTotal + " slides");
                    }
                }

                // 3. CLONE STRATEGY:
                //    a) Append clones at end of deck
                //    b) Remove all original template slides
                //    c) Only our clones remain

                int originalCount = sourceSlides.size();
                log.info("Cloning {} recipe slides from {} template slides",
                        recipe.size(), originalCount);

                // 3a. Append clones
                for (int i = 0; i < recipe.size(); i++) {
                    int srcIdx = sourceIndices.get(i);
                    XSLFSlide srcSlide = sourceSlides.get(srcIdx);

                    XSLFSlide cloned = source.createSlide(srcSlide.getSlideLayout());
                    cloneSlideViaXml(srcSlide, cloned);

                    log.debug("Cloned template slide {} -> position {} (layout: {})",
                            srcIdx + 1, originalCount + i + 1, recipe.get(i).get("layout"));
                }

                // 3b. Remove ALL original template slides
                for (int i = originalCount - 1; i >= 0; i--) {
                    source.removeSlide(i);
                }

                // 4. Replace placeholders and embed charts
                List<XSLFSlide> outputSlides = source.getSlides();
                log.info("After cleanup: {} slides remain (expected {})",
                        outputSlides.size(), recipe.size());

                for (int i = 0; i < outputSlides.size(); i++) {
                    XSLFSlide slide = outputSlides.get(i);
                    Map<String, Object> slideSpec = recipe.get(i);
                    String layoutId = (String) slideSpec.get("layout");

                    // Replace {{placeholders}} — use Map<String, ?> to handle Integer values
                    Map<String, ?> placeholders = (Map<String, ?>) slideSpec.get("placeholders");
                    if (placeholders != null) {
                        replacePlaceholders(slide, placeholders);
                    }

                    // Handle CHART_IMAGE slides
                    if ("CHART_IMAGE".equals(layoutId)) {
                        Map<String, Object> chartData =
                                (Map<String, Object>) slideSpec.get("chartData");
                        String chartType = (String) slideSpec.getOrDefault("chartType", "kaplan-meier");
                        if (chartData != null) {
                            embedChart(source, slide, chartType, chartData);
                        }
                    }
                }

                // 5. Write output
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

    // ═══════════════════════════════════════════════════════════
    // CLONE HELPERS
    // ═══════════════════════════════════════════════════════════

    private void cloneSlideViaXml(XSLFSlide src, XSLFSlide dst) {
        try {
            var srcCsld = src.getXmlObject().getCSld();
            var dstXml = dst.getXmlObject();

            if (srcCsld.getSpTree() != null) {
                dstXml.getCSld().getSpTree().set(srcCsld.getSpTree().copy());
            }
        } catch (Exception e) {
            log.warn("XML clone failed, falling back to shape-by-shape: {}", e.getMessage());
            copyShapesManually(src, dst);
        }
    }

    private void copyShapesManually(XSLFSlide src, XSLFSlide dst) {
        for (XSLFShape shape : src.getShapes()) {
            try {
                if (shape instanceof XSLFTextShape textShape) {
                    XSLFTextBox copy = dst.createTextBox();
                    copy.setAnchor(textShape.getAnchor());
                    copy.clearText();
                    for (XSLFTextParagraph srcPara : textShape.getTextParagraphs()) {
                        XSLFTextParagraph dstPara = copy.addNewTextParagraph();
                        try {
                            dstPara.setTextAlign(srcPara.getTextAlign());
                        } catch (Exception ignored) {
                        }
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
                            }
                        }
                    }
                }
            } catch (Exception e) {
                log.debug("Shape copy skipped: {}", e.getMessage());
            }
        }
    }

    // ═══════════════════════════════════════════════════════════
    // PLACEHOLDER REPLACEMENT
    // ═══════════════════════════════════════════════════════════

    private void replacePlaceholders(XSLFSlide slide, Map<String, ?> placeholders) {
        for (XSLFShape shape : slide.getShapes()) {
            replaceInShape(shape, placeholders);
        }
    }

    private void replaceInShape(XSLFShape shape, Map<String, ?> placeholders) {
        if (shape instanceof XSLFTextShape textShape) {
            replaceInTextShape(textShape, placeholders);

        } else if (shape instanceof XSLFGroupShape groupShape) {
            for (XSLFShape child : groupShape.getShapes()) {
                replaceInShape(child, placeholders);
            }

        } else if (shape instanceof XSLFTable table) {
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

    private void replaceInTextShape(XSLFTextShape textShape, Map<String, ?> placeholders) {
        for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
            List<XSLFTextRun> runs = para.getTextRuns();
            if (runs.isEmpty()) continue;

            // 1. Concatenate ALL runs to get the full paragraph text
            StringBuilder sb = new StringBuilder();
            for (XSLFTextRun run : runs) {
                String t = run.getRawText();
                if (t != null) sb.append(t);
            }
            String fullText = sb.toString();

            // 2. Skip if no placeholders in this paragraph
            if (!fullText.contains("{{")) continue;

            // 3. Replace all known placeholders
            for (Map.Entry<String, ?> entry : placeholders.entrySet()) {
                String key = "{{" + entry.getKey() + "}}";
                String value = entry.getValue() != null ? String.valueOf(entry.getValue()) : "";
                fullText = fullText.replace(key, value);
            }

            // 4. Clean up any remaining unreplaced {{...}} placeholders
            //    (e.g. step_5 when only 4 steps exist)
            fullText = fullText.replaceAll("\\{\\{[^}]*\\}\\}", "");

            // 5. Put the result in the FIRST run, clear all others
            //    This preserves the formatting of the first run
            boolean first = true;
            for (XSLFTextRun run : runs) {
                if (first) {
                    run.setText(fullText);
                    first = false;
                } else {
                    run.setText("");
                }
            }
        }
    }

    // ═══════════════════════════════════════════════════════════
    // CHART EMBEDDING
    // ═══════════════════════════════════════════════════════════

    private void embedChart(XMLSlideShow pptx, XSLFSlide slide,
                            String chartType, Map<String, Object> chartData) throws Exception {
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
            return;
        }

        log.info("Received {} KB chart PNG for type '{}'", pngBytes.length / 1024, chartType);

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

        Rectangle2D anchor;
        if (placeholderShape != null) {
            anchor = placeholderShape.getAnchor();
            slide.removeShape(placeholderShape);
        } else {
            log.warn("No chart placeholder found, using default position");
            anchor = new Rectangle2D.Double(48, 108, 864, 360);
        }

        XSLFPictureData picData = pptx.addPicture(pngBytes, PictureData.PictureType.PNG);
        XSLFPictureShape pic = slide.createPicture(picData);
        pic.setAnchor(anchor);

        log.info("Chart embedded at ({}, {}) size ({} x {})",
                (int) anchor.getX(), (int) anchor.getY(),
                (int) anchor.getWidth(), (int) anchor.getHeight());
    }

    // ═══════════════════════════════════════════════════════════
    // TEMPLATE FILE COPY
    // ═══════════════════════════════════════════════════════════

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
