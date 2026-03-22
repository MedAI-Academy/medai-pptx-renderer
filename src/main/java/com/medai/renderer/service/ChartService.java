package com.medai.renderer.service;

import com.medai.renderer.template.ThemeConfig;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.renderer.xy.XYStepRenderer;
import org.jfree.chart.title.LegendTitle;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.text.NumberFormat;
import java.util.List;
import java.util.Map;

/**
 * Generates high-resolution chart images (PNG) using JFreeChart.
 *
 * v2.0 — Prezent-Premium Upgrade:
 * - generateBarChart() for ORR comparisons, efficacy data
 * - generateSwotMatrix() marked @Deprecated (replaced by native POI shapes)
 * - Improved chart styling for dark theme
 */
@Service
public class ChartService {

    private static final int CHART_W = 1600;
    private static final int CHART_H = 900;

    // ═══════════════════════════════════════════════════════════
    // KAPLAN-MEIER SURVIVAL CURVES
    // ═══════════════════════════════════════════════════════════

    @SuppressWarnings("unchecked")
    public byte[] generateKaplanMeier(Map<String, Object> chartData) throws Exception {
        List<Map<String, Object>> arms = (List<Map<String, Object>>) chartData.get("arms");
        String xlabel = (String) chartData.getOrDefault("xlabel", "Time (months)");
        String ylabel = (String) chartData.getOrDefault("ylabel", "Survival Probability");
        String hrText = (String) chartData.getOrDefault("hazardRatio", "");

        XYSeriesCollection dataset = new XYSeriesCollection();

        for (Map<String, Object> arm : arms) {
            String name = (String) arm.get("name");
            List<Number> timepoints = (List<Number>) arm.get("timepoints");
            List<Number> survival = (List<Number>) arm.get("survival");
            Number median = (Number) arm.get("median");

            XYSeries series = new XYSeries(name + (median != null ?
                " (median: " + median + " mo)" : ""));

            for (int i = 0; i < Math.min(timepoints.size(), survival.size()); i++) {
                series.add(timepoints.get(i).doubleValue(), survival.get(i).doubleValue());
            }
            dataset.addSeries(series);
        }

        JFreeChart chart = ChartFactory.createXYLineChart(
            null, xlabel, ylabel, dataset
        );

        XYPlot plot = chart.getXYPlot();
        plot.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
        plot.setDomainGridlinePaint(new Color(74, 106, 154, 80));
        plot.setRangeGridlinePaint(new Color(74, 106, 154, 80));
        plot.setOutlinePaint(ThemeConfig.CLR_SURFACE);

        XYStepRenderer renderer = new XYStepRenderer();
        for (int i = 0; i < arms.size(); i++) {
            Map<String, Object> arm = arms.get(i);
            String colorHex = ((String) arm.getOrDefault("color", "#7C6FFF")).replace("#", "");
            renderer.setSeriesPaint(i, ThemeConfig.hex(colorHex));
            renderer.setSeriesStroke(i, new BasicStroke(2.5f));
        }
        plot.setRenderer(renderer);

        NumberAxis domainAxis = (NumberAxis) plot.getDomainAxis();
        domainAxis.setLabelPaint(ThemeConfig.CLR_MUTED);
        domainAxis.setTickLabelPaint(ThemeConfig.CLR_MUTED);
        domainAxis.setLabelFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 14));

        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setRange(0.0, 1.05);
        rangeAxis.setNumberFormatOverride(NumberFormat.getPercentInstance());
        rangeAxis.setLabelPaint(ThemeConfig.CLR_MUTED);
        rangeAxis.setTickLabelPaint(ThemeConfig.CLR_MUTED);
        rangeAxis.setLabelFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 14));

        chart.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
        LegendTitle legend = chart.getLegend();
        if (legend != null) {
            legend.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
            legend.setItemPaint(ThemeConfig.CLR_TEXT);
            legend.setItemFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 13));
        }

        BufferedImage image = chart.createBufferedImage(CHART_W, CHART_H);

        if (hrText != null && !hrText.isEmpty()) {
            Graphics2D g2 = image.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setFont(new Font(ThemeConfig.FONT_BODY, Font.BOLD, 16));
            g2.setColor(ThemeConfig.CLR_TEAL);
            g2.drawString(hrText, CHART_W - g2.getFontMetrics().stringWidth(hrText) - 20, 30);
            g2.dispose();
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }

    // ═══════════════════════════════════════════════════════════
    // BAR CHART — for ORR comparisons, efficacy data
    // ═══════════════════════════════════════════════════════════

    /**
     * Generate a horizontal or vertical bar chart as PNG.
     *
     * @param chartData Map containing:
     *   - bars: List of {label, value, color?}
     *   - xlabel, ylabel
     *   - title (optional — usually set via slide header)
     *   - orientation: "horizontal" or "vertical" (default: vertical)
     */
    @SuppressWarnings("unchecked")
    public byte[] generateBarChart(Map<String, Object> chartData) throws Exception {
        List<Map<String, Object>> bars = (List<Map<String, Object>>) chartData.get("bars");
        String xlabel = (String) chartData.getOrDefault("xlabel", "");
        String ylabel = (String) chartData.getOrDefault("ylabel", "");
        String orientation = (String) chartData.getOrDefault("orientation", "vertical");

        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        for (Map<String, Object> bar : bars) {
            String label = (String) bar.getOrDefault("label", "");
            Number value = (Number) bar.getOrDefault("value", 0);
            dataset.addValue(value, "Value", label);
        }

        PlotOrientation orient = "horizontal".equalsIgnoreCase(orientation)
            ? PlotOrientation.HORIZONTAL : PlotOrientation.VERTICAL;

        JFreeChart chart = ChartFactory.createBarChart(
            null, xlabel, ylabel, dataset, orient, false, false, false
        );

        // Style for dark theme
        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(ThemeConfig.CLR_SURFACE);
        plot.setDomainGridlinePaint(new Color(74, 106, 154, 80));
        plot.setRangeGridlinePaint(new Color(74, 106, 154, 80));
        plot.setOutlinePaint(ThemeConfig.CLR_SURFACE);

        // Bar renderer — flat style (no gradient), accent colors
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBarPainter(new StandardBarPainter()); // Flat bars, no shine
        renderer.setShadowVisible(false);
        renderer.setMaximumBarWidth(0.12);

        // Color each bar individually
        for (int i = 0; i < bars.size(); i++) {
            Map<String, Object> bar = bars.get(i);
            String colorHex = ((String) bar.getOrDefault("color",
                ThemeConfig.ACCENT_CYCLE[i % ThemeConfig.ACCENT_CYCLE.length]))
                .replace("#", "");
            renderer.setSeriesPaint(i, ThemeConfig.hex(colorHex));
        }

        // If all bars are in same series, color by item
        if (bars.size() > 1 && dataset.getRowCount() == 1) {
            for (int i = 0; i < bars.size(); i++) {
                String colorHex = ((String) bars.get(i).getOrDefault("color",
                    ThemeConfig.ACCENT_CYCLE[i % ThemeConfig.ACCENT_CYCLE.length]))
                    .replace("#", "");
                renderer.setItemPaint(0, i, ThemeConfig.hex(colorHex));
            }
        }

        // Axis styling
        CategoryAxis domainAxis = plot.getDomainAxis();
        domainAxis.setLabelPaint(ThemeConfig.CLR_MUTED);
        domainAxis.setTickLabelPaint(ThemeConfig.CLR_TEXT);
        domainAxis.setTickLabelFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 12));

        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setLabelPaint(ThemeConfig.CLR_MUTED);
        rangeAxis.setTickLabelPaint(ThemeConfig.CLR_MUTED);
        rangeAxis.setLabelFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 14));

        chart.setBackgroundPaint(ThemeConfig.CLR_SURFACE);

        // Add value labels on top of bars
        BufferedImage image = chart.createBufferedImage(CHART_W, CHART_H);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }

    // ═══════════════════════════════════════════════════════════
    // SWOT MATRIX — @Deprecated: replaced by native POI shapes
    // ═══════════════════════════════════════════════════════════

    /**
     * @deprecated Since v2.0 — SWOT is now rendered as native POI shapes
     *             in PptxRenderService.buildSwot() for better quality.
     *             Kept for backward compatibility only.
     */
    @Deprecated
    public byte[] generateSwotMatrix(Map<String, Object> swotData) throws Exception {
        int W = 1600, H = 1000;
        BufferedImage image = new BufferedImage(W, H, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = image.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        g2.setColor(ThemeConfig.CLR_SURFACE);
        g2.fillRect(0, 0, W, H);

        Color[] colors = {
            ThemeConfig.hex(ThemeConfig.HEX_SWOT_S), ThemeConfig.hex(ThemeConfig.HEX_SWOT_W),
            ThemeConfig.hex(ThemeConfig.HEX_SWOT_O), ThemeConfig.hex(ThemeConfig.HEX_SWOT_T)
        };
        String[] labels = {"STRENGTHS", "WEAKNESSES", "OPPORTUNITIES", "THREATS"};
        String[] icons  = {"\u2714", "\u26A0", "\u25B2", "\u25CF"};
        String[] keys   = {"strengths", "weaknesses", "opportunities", "threats"};

        int halfW = W / 2 - 8, halfH = H / 2 - 8;
        int[][] positions = {{4, 4}, {W/2 + 4, 4}, {4, H/2 + 4}, {W/2 + 4, H/2 + 4}};

        for (int i = 0; i < 4; i++) {
            int qx = positions[i][0], qy = positions[i][1];
            g2.setColor(colors[i]);
            g2.fillRoundRect(qx, qy, halfW, halfH, 12, 12);

            g2.setFont(new Font(ThemeConfig.FONT_TITLE, Font.BOLD, 18));
            g2.setColor(ThemeConfig.CLR_WHITE);
            g2.drawString(icons[i] + "  " + labels[i], qx + 16, qy + 34);

            @SuppressWarnings("unchecked")
            List<String> items = (List<String>) swotData.getOrDefault(keys[i], List.of());
            g2.setFont(new Font(ThemeConfig.FONT_BODY, Font.PLAIN, 13));
            g2.setColor(ThemeConfig.CLR_TEXT);
            int ty = qy + 60;
            for (int j = 0; j < Math.min(items.size(), 6); j++) {
                String item = items.get(j);
                if (item.length() > 80) item = item.substring(0, 77) + "...";
                g2.drawString("\u2022 " + item, qx + 20, ty);
                ty += 22;
            }
        }

        g2.dispose();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }
}
