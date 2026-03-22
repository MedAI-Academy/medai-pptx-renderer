package com.medai.renderer.util;

import com.medai.renderer.template.ThemeConfig;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackground;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackgroundProperties;

import java.awt.*;
import java.awt.geom.Rectangle2D;

/**
 * Utility methods for Apache POI XSLF (PPTX) manipulation.
 *
 * v2.0 — Prezent-Premium Upgrade:
 * - addIconCircle() for card headers
 * - addConfidenceBadge() with colored dot
 * - Improved addHeader() with accent left-bar
 * - addRoundRect() for SWOT quadrants
 */
public class PptxUtils {

    private static final long EMU_PER_INCH = 914400L;
    private static final long EMU_PER_PT   = 12700L;

    public static long emu(double inches) { return Math.round(inches * EMU_PER_INCH); }
    public static long ptEmu(double points) { return Math.round(points * EMU_PER_PT); }

    // ═══════════════════════════════════════════════════════════
    // SLIDE BACKGROUND
    // ═══════════════════════════════════════════════════════════

    public static void setBackground(XSLFSlide slide, String hexColor) {
        CTBackground bg = slide.getXmlObject().getCSld().addNewBg();
        CTBackgroundProperties bgPr = bg.addNewBgPr();
        CTSolidColorFillProperties solidFill = bgPr.addNewSolidFill();
        CTSRgbColor srgb = solidFill.addNewSrgbClr();
        srgb.setVal(hexToBytes(hexColor));
        bgPr.addNewEffectLst();
    }

    // ═══════════════════════════════════════════════════════════
    // SHAPES
    // ═══════════════════════════════════════════════════════════

    /** Add a filled rectangle shape */
    public static XSLFAutoShape addRect(XSLFSlide slide,
                                         double x, double y, double w, double h,
                                         String fillHex) {
        XSLFAutoShape shape = slide.createAutoShape();
        shape.setShapeType(ShapeType.RECT);
        shape.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        shape.setFillColor(ThemeConfig.hex(fillHex));
        shape.setLineWidth(0.0);
        return shape;
    }

    /** Add a rounded rectangle (for SWOT quadrants, premium cards) */
    public static XSLFAutoShape addRoundRect(XSLFSlide slide,
                                              double x, double y, double w, double h,
                                              String fillHex) {
        XSLFAutoShape shape = slide.createAutoShape();
        shape.setShapeType(ShapeType.ROUND_RECT);
        shape.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        shape.setFillColor(ThemeConfig.hex(fillHex));
        shape.setLineWidth(0.0);
        return shape;
    }

    /** Add an ellipse/circle shape (for icon circles, timeline dots) */
    public static XSLFAutoShape addEllipse(XSLFSlide slide,
                                            double x, double y, double size,
                                            String fillHex) {
        XSLFAutoShape shape = slide.createAutoShape();
        shape.setShapeType(ShapeType.ELLIPSE);
        shape.setAnchor(new Rectangle2D.Double(x * 72, y * 72, size * 72, size * 72));
        shape.setFillColor(ThemeConfig.hex(fillHex));
        shape.setLineWidth(0.0);
        return shape;
    }

    // ═══════════════════════════════════════════════════════════
    // ICON CIRCLES (Prezent-Premium: replaces emoji icons)
    // ═══════════════════════════════════════════════════════════

    /**
     * Draw a colored circle with a Unicode symbol inside.
     * Used for card headers, SWOT labels, KPI indicators.
     *
     * @param slide   Target slide
     * @param x       X position in inches
     * @param y       Y position in inches
     * @param size    Diameter in inches (typically 0.28 - 0.35)
     * @param bgHex   Circle background color
     * @param symbol  Unicode symbol string
     */
    public static void addIconCircle(XSLFSlide slide, double x, double y,
                                      double size, String bgHex, String symbol) {
        // Circle background
        addEllipse(slide, x, y, size, bgHex);

        // Symbol text centered inside circle
        XSLFTextBox label = slide.createTextBox();
        label.setAnchor(new Rectangle2D.Double(
            x * 72, y * 72, size * 72, size * 72));
        label.setVerticalAlignment(org.apache.poi.sl.usermodel.VerticalAlignment.MIDDLE);
        label.setWordWrap(false);
        // Remove default insets so text centers properly
        label.setTopInset(0); label.setBottomInset(0);
        label.setLeftInset(0); label.setRightInset(0);

        XSLFTextParagraph p = label.getTextParagraphs().get(0);
        p.setTextAlign(TextParagraph.TextAlign.CENTER);
        XSLFTextRun r = p.addNewTextRun();
        r.setText(symbol);
        r.setFontSize(size * 28);  // Scale font relative to circle
        r.setFontColor(Color.WHITE);
        r.setFontFamily(ThemeConfig.FONT_ICON);
    }

    // ═══════════════════════════════════════════════════════════
    // HEADER — Improved with left accent bar
    // ═══════════════════════════════════════════════════════════

    /** Add header bar with left accent stripe (Prezent-Premium style) */
    public static void addHeader(XSLFSlide slide, String title, String subtitle) {
        // Navy header background
        addRect(slide, 0, 0, ThemeConfig.SLIDE_W, ThemeConfig.HEADER_H, ThemeConfig.HEX_NAVY);

        // Left accent bar (teal, 5px wide, full header height)
        addRect(slide, 0, 0, 0.07, ThemeConfig.HEADER_H, ThemeConfig.HEX_TEAL);

        // Thin top accent line (accent purple, 2px)
        addRect(slide, 0, 0, ThemeConfig.SLIDE_W, 0.03, ThemeConfig.HEX_ACCENT);

        // Title text
        addText(slide, title,
            0.30, 0, 9.5, ThemeConfig.HEADER_H,
            ThemeConfig.SIZE_HEADING - 2, ThemeConfig.FONT_TITLE, ThemeConfig.HEX_WHITE,
            true, TextParagraph.TextAlign.LEFT);

        // Subtitle (right-aligned, section label)
        if (subtitle != null && !subtitle.isEmpty()) {
            addText(slide, subtitle,
                10.0, 0, 3.1, ThemeConfig.HEADER_H,
                ThemeConfig.SIZE_SMALL - 1, ThemeConfig.FONT_BODY, ThemeConfig.HEX_ACCENT,
                false, TextParagraph.TextAlign.RIGHT);
        }
    }

    // ═══════════════════════════════════════════════════════════
    // FOOTER — with Confidence Badge
    // ═══════════════════════════════════════════════════════════

    /** Add footer with AI disclaimer + optional confidence badge */
    public static void addFooter(XSLFSlide slide, String confidenceBadge) {
        String disclaimer = "\u26A0 AI-generated \u2014 Verify before external use \u00B7 MedAI Suite";
        addText(slide, disclaimer,
            ThemeConfig.CONTENT_X, ThemeConfig.FOOTER_Y, ThemeConfig.CONTENT_W, ThemeConfig.FOOTER_H,
            ThemeConfig.SIZE_FOOTER, ThemeConfig.FONT_BODY, ThemeConfig.HEX_DIM,
            false, TextParagraph.TextAlign.LEFT);

        if (confidenceBadge != null && !confidenceBadge.isEmpty()) {
            addText(slide, confidenceBadge,
                ThemeConfig.CONTENT_X, ThemeConfig.FOOTER_Y - 0.22, ThemeConfig.CONTENT_W, 0.22,
                ThemeConfig.SIZE_FOOTER, ThemeConfig.FONT_MONO, ThemeConfig.HEX_TEAL,
                false, TextParagraph.TextAlign.RIGHT);
        }
    }

    /**
     * Add a visual confidence score badge with colored indicator dot.
     * Shows: [●] Sources: X (Y Gold) | Liability: Z% | MedAI Suite
     *
     * @param slide        Target slide
     * @param totalSources Number of sources cited
     * @param goldSources  Number of Tier-1 (Gold) sources
     * @param score        Confidence score 0-100
     */
    public static void addConfidenceBadgeVisual(XSLFSlide slide,
                                                 int totalSources, int goldSources, double score) {
        String badgeColorHex = ThemeConfig.confColor(score);
        String badgeText = String.format(
            "Sources: %d (%d Gold) | Liability: %.0f%% | MedAI Suite",
            totalSources, goldSources, score);

        double dotX = 10.0;
        double dotY = ThemeConfig.FOOTER_Y - 0.18;
        double dotSize = 0.12;

        // Colored indicator dot
        addEllipse(slide, dotX, dotY, dotSize, badgeColorHex);

        // Badge text next to dot
        addText(slide, badgeText,
            dotX + 0.18, dotY - 0.03, 3.0, 0.20,
            7.0, ThemeConfig.FONT_MONO, badgeColorHex,
            false, TextParagraph.TextAlign.RIGHT);
    }

    // ═══════════════════════════════════════════════════════════
    // TEXT
    // ═══════════════════════════════════════════════════════════

    /** Add a text box with full formatting control */
    public static XSLFTextBox addText(XSLFSlide slide, String text,
                                       double x, double y, double w, double h,
                                       double fontSize, String fontName, String colorHex,
                                       boolean bold, TextParagraph.TextAlign align) {
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        textBox.setWordWrap(true);
        textBox.setVerticalAlignment(org.apache.poi.sl.usermodel.VerticalAlignment.MIDDLE);
        textBox.setTextAutofit(XSLFTextShape.TextAutofit.NORMAL);

        XSLFTextParagraph para = textBox.getTextParagraphs().get(0);
        para.setTextAlign(align);

        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        run.setFontSize(fontSize);
        run.setFontFamily(fontName);
        run.setFontColor(ThemeConfig.hex(colorHex));
        run.setBold(bold);

        return textBox;
    }

    /**
     * Add text with mixed formatting (e.g., icon + label in same box).
     * Returns the text box for further manipulation.
     */
    public static XSLFTextBox addTextBox(XSLFSlide slide,
                                          double x, double y, double w, double h) {
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        textBox.setWordWrap(true);
        return textBox;
    }

    /** Add multi-line text with bullet points */
    public static XSLFTextBox addBulletText(XSLFSlide slide, String[] items,
                                             double x, double y, double w, double h,
                                             double fontSize, String colorHex) {
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
        textBox.setWordWrap(true);
        textBox.setTextAutofit(XSLFTextShape.TextAutofit.NORMAL);

        textBox.getTextParagraphs().get(0).getTextRuns().forEach(r -> r.setText(""));

        for (int i = 0; i < items.length; i++) {
            XSLFTextParagraph para;
            if (i == 0) {
                para = textBox.getTextParagraphs().get(0);
            } else {
                para = textBox.addNewTextParagraph();
            }
            para.setTextAlign(TextParagraph.TextAlign.LEFT);
            para.setBullet(true);
            para.setIndentLevel(0);

            XSLFTextRun run = para.addNewTextRun();
            run.setText(items[i]);
            run.setFontSize(fontSize);
            run.setFontFamily(ThemeConfig.FONT_BODY);
            run.setFontColor(ThemeConfig.hex(colorHex));
        }
        return textBox;
    }

    // ═══════════════════════════════════════════════════════════
    // TABLES — with zebra striping
    // ═══════════════════════════════════════════════════════════

    /** Create a styled data table with zebra striping and smart column widths */
    public static XSLFTable addTable(XSLFSlide slide,
                                      String[] headers, String[][] rows,
                                      double x, double y, double w, double rowH) {
        int numRows = rows.length + 1;
        int numCols = headers.length;
        double[] colWidths = computeColumnWidths(headers, rows, w);

        XSLFTable table = slide.createTable(numRows, numCols);
        table.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, rowH * numRows * 72));

        for (int c = 0; c < numCols; c++) {
            table.setColumnWidth(c, colWidths[c] * 72);
        }

        // Header row — accent purple background
        for (int c = 0; c < numCols; c++) {
            XSLFTableCell cell = table.getCell(0, c);
            setCellText(cell, headers[c], ThemeConfig.SIZE_SMALL, ThemeConfig.HEX_WHITE, true);
            setCellFill(cell, ThemeConfig.HEX_ACCENT);
        }

        // Data rows — zebra striping
        for (int r = 0; r < rows.length; r++) {
            String bgHex = (r % 2 == 0) ? ThemeConfig.HEX_ZEBRA_EVEN : ThemeConfig.HEX_ZEBRA_ODD;
            for (int c = 0; c < Math.min(rows[r].length, numCols); c++) {
                XSLFTableCell cell = table.getCell(r + 1, c);
                setCellText(cell, rows[r][c], ThemeConfig.SIZE_SMALL - 1, ThemeConfig.HEX_TEXT, false);
                setCellFill(cell, bgHex);
            }
            // Fill remaining cells if row is shorter than headers
            for (int c = rows[r].length; c < numCols; c++) {
                XSLFTableCell cell = table.getCell(r + 1, c);
                setCellText(cell, "", ThemeConfig.SIZE_SMALL - 1, ThemeConfig.HEX_TEXT, false);
                setCellFill(cell, bgHex);
            }
        }

        return table;
    }

    /** Compute proportional column widths based on content length */
    private static double[] computeColumnWidths(String[] headers, String[][] rows, double totalW) {
        int n = headers.length;
        double[] weights = new double[n];

        for (int c = 0; c < n; c++) {
            weights[c] = Math.max(3.0, headers[c].length());
        }

        for (int r = 0; r < Math.min(rows.length, 5); r++) {
            for (int c = 0; c < Math.min(rows[r].length, n); c++) {
                double contentLen = rows[r][c] != null ? rows[r][c].length() : 0;
                weights[c] = Math.max(weights[c], contentLen * 0.8);
            }
        }

        // Cap narrow columns by known header names
        for (int c = 0; c < n; c++) {
            String h = headers[c].toLowerCase();
            if (h.equals("n") || h.equals("phase") || h.equals("status") || h.equals("si")
                || h.equals("budget") || h.equals("orr") || h.equals("year")) {
                weights[c] = Math.min(weights[c], 8.0);
            }
        }

        // Normalize with minimum column width
        double sum = 0;
        for (double w2 : weights) sum += w2;
        double[] widths = new double[n];
        double minCol = 0.8;
        double remaining = totalW;

        for (int c = 0; c < n; c++) {
            double proposed = (weights[c] / sum) * totalW;
            if (proposed < minCol) {
                widths[c] = minCol;
                remaining -= minCol;
            }
        }

        double uncappedSum = 0;
        for (int c = 0; c < n; c++) {
            if (widths[c] == 0) uncappedSum += weights[c];
        }
        for (int c = 0; c < n; c++) {
            if (widths[c] == 0) {
                widths[c] = uncappedSum > 0 ? (weights[c] / uncappedSum) * remaining : remaining / n;
            }
        }

        return widths;
    }

    private static void setCellText(XSLFTableCell cell, String text,
                                     double fontSize, String colorHex, boolean bold) {
        cell.clearText();
        cell.setVerticalAlignment(org.apache.poi.sl.usermodel.VerticalAlignment.TOP);
        cell.setTopInset(4.0);
        cell.setBottomInset(4.0);
        cell.setLeftInset(6.0);
        cell.setRightInset(6.0);
        XSLFTextParagraph para = cell.addNewTextParagraph();
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text != null ? text : "");
        run.setFontSize(fontSize);
        run.setFontFamily(ThemeConfig.FONT_BODY);
        run.setFontColor(ThemeConfig.hex(colorHex));
        run.setBold(bold);
    }

    private static void setCellFill(XSLFTableCell cell, String hexColor) {
        cell.setFillColor(ThemeConfig.hex(hexColor));
    }

    // ═══════════════════════════════════════════════════════════
    // IMAGES
    // ═══════════════════════════════════════════════════════════

    public static void addImage(XSLFSlide slide, XMLSlideShow pptx, byte[] imageData,
                                 double x, double y, double w, double h) {
        XSLFPictureData pictureData = pptx.addPicture(imageData,
            org.apache.poi.sl.usermodel.PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(pictureData);
        picture.setAnchor(new Rectangle2D.Double(x * 72, y * 72, w * 72, h * 72));
    }

    // ═══════════════════════════════════════════════════════════
    // UTILITY
    // ═══════════════════════════════════════════════════════════

    public static byte[] hexToBytes(String hex) {
        return new byte[] {
            (byte) Integer.parseInt(hex.substring(0, 2), 16),
            (byte) Integer.parseInt(hex.substring(2, 4), 16),
            (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
    }
}
