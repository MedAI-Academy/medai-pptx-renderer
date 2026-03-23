package com.medai.renderer.service;

import com.medai.renderer.template.ThemeConfig;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.List;

/**
 * Chart rendering service — delegates to Python Chart Service for publication-quality output.
 * Falls back to Java2D if Python service is unavailable.
 * v4.0 — Hybrid: Python primary, Java2D fallback. ALL original methods preserved.
 */
@Service
public class ChartService {

    private static final Logger log = LoggerFactory.getLogger(ChartService.class);

    @Autowired
    private PythonChartClient pythonCharts;

    private static final int W = 1800, H = 1050;
    private static final int ML = 120, MR = 40, MT = 50, MB_PLOT = 60;

    private static final Color BG = ThemeConfig.CLR_SURFACE;
    private static final Color GRID = new Color(74, 106, 154, 50);
    private static final Color AXIS_CLR = ThemeConfig.CLR_MUTED;
    private static final Color TEXT_CLR = ThemeConfig.CLR_TEXT;

    private static final Color[] ARM_COLORS = {
        new Color(66, 133, 244), new Color(234, 67, 53),
        new Color(52, 168, 83), new Color(156, 39, 176), new Color(255, 152, 0)
    };

    private static final Font F_AXIS = new Font("Calibri", Font.PLAIN, 16);
    private static final Font F_TICK = new Font("Calibri", Font.PLAIN, 13);
    private static final Font F_LEG  = new Font("Calibri", Font.PLAIN, 14);
    private static final Font F_LEGB = new Font("Calibri", Font.BOLD, 14);
    private static final Font F_HR   = new Font("Calibri", Font.PLAIN, 13);
    private static final Font F_HRB  = new Font("Calibri", Font.BOLD, 14);
    private static final Font F_ARH  = new Font("Calibri", Font.BOLD, 11);
    private static final Font F_AR   = new Font("Calibri", Font.PLAIN, 11);

    // ═════════════════════════════════════════════
    // 1. KAPLAN-MEIER — Python primary, Java2D fallback
    // ═════════════════════════════════════════════
    @SuppressWarnings("unchecked")
    public byte[] generateKaplanMeier(Map<String, Object> data) throws Exception {
        // Try Python service first
        try {
            if (pythonCharts.isAvailable()) {
                log.info("Delegating KM chart to Python service");
                return pythonCharts.fetchKaplanMeier(data);
            } else {
                log.warn("Python chart service not available, using Java2D fallback");
            }
        } catch (Exception e) {
            log.error("Python chart service failed, falling back to Java2D: {}", e.getMessage());
        }
        return generateKaplanMeierJava2D(data);
    }

    @SuppressWarnings("unchecked")
    private byte[] generateKaplanMeierJava2D(Map<String, Object> data) throws Exception {
        List<Map<String, Object>> arms = (List<Map<String, Object>>) data.get("arms");
        String xlabel = s(data, "xlabel", "Time (months)");
        String ylabel = s(data, "ylabel", "Progression-Free Survival");
        String hrText = s(data, "hazardRatio", "");
        boolean showAR = b(data, "showAtRisk", true);
        boolean showMed = b(data, "showMedian", true);
        boolean showCens = b(data, "showCensoring", true);

        int arH = showAR ? arms.size() * 20 + 30 : 0;
        int pH = H - MT - MB_PLOT - arH;
        int pW = W - ML - MR;
        int pB = MT + pH;
        double xMax = n(data, "xmax", autoXMax(arms));

        BufferedImage img = new BufferedImage(W, H, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = img.createGraphics();
        aa(g);
        g.setColor(BG); g.fillRect(0, 0, W, H);

        // Grid
        g.setColor(GRID); g.setStroke(new BasicStroke(0.5f));
        double xStep = niceStep(xMax);
        for (double x = xStep; x < xMax; x += xStep) { int px = xPx(x, pW, xMax); g.drawLine(px, MT, px, pB); }
        for (double y = 0.1; y <= 1.0; y += 0.1) { int py = yPx(y, pH); g.drawLine(ML, py, ML + pW, py); }
        g.setColor(GRID); g.setStroke(new BasicStroke(1f)); g.drawRect(ML, MT, pW, pH);

        // Arms
        for (int a = 0; a < arms.size(); a++) {
            Map<String, Object> arm = arms.get(a);
            List<Number> tp = nl(arm, "timepoints"), sv = nl(arm, "survival"), cens = nl(arm, "censored");
            Color c = armClr(arm, a);
            g.setColor(c); g.setStroke(new BasicStroke(2.5f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));

            GeneralPath path = new GeneralPath(); boolean first = true;
            double px = 0, py = 1.0;
            for (int i = 0; i < Math.min(tp.size(), sv.size()); i++) {
                double tx = tp.get(i).doubleValue(), sy = sv.get(i).doubleValue();
                if (first) { path.moveTo(xPx(px, pW, xMax), yPx(py, pH)); first = false; }
                path.lineTo(xPx(tx, pW, xMax), yPx(py, pH));
                path.lineTo(xPx(tx, pW, xMax), yPx(sy, pH));
                px = tx; py = sy;
            }
            path.lineTo(xPx(px, pW, xMax), yPx(py, pH));
            g.draw(path);

            // Censoring ticks
            if (showCens && !cens.isEmpty()) {
                g.setStroke(new BasicStroke(1.5f));
                for (Number ct : cens) { double cx = ct.doubleValue(), cy = svAt(tp, sv, cx);
                    int cpx = xPx(cx, pW, xMax), cpy = yPx(cy, pH);
                    g.drawLine(cpx, cpy - 5, cpx, cpy + 5); }
            }

            // Median dashed lines
            if (showMed && arm.get("median") != null) {
                double med = ((Number) arm.get("median")).doubleValue();
                if (med > 0 && med <= xMax) {
                    g.setStroke(new BasicStroke(1.2f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER, 10f, new float[]{6f, 4f}, 0f));
                    g.setColor(new Color(c.getRed(), c.getGreen(), c.getBlue(), 120));
                    int mx = xPx(med, pW, xMax), my = yPx(0.5, pH);
                    g.drawLine(mx, my, mx, pB); g.drawLine(ML, my, mx, my);
                }
            }
        }

        // Axes
        g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f));
        g.drawLine(ML, pB, ML + pW, pB); g.drawLine(ML, MT, ML, pB);
        g.setFont(F_TICK);
        for (double x = 0; x <= xMax; x += xStep) { int px = xPx(x, pW, xMax); g.drawLine(px, pB, px, pB + 5); cStr(g, fmtI(x), px, pB + 20); }
        for (double y = 0.0; y <= 1.01; y += 0.2) { int py = yPx(y, pH); g.drawLine(ML - 5, py, ML, py); rStr(g, String.format("%.1f", y), ML - 10, py + 4); }
        g.setFont(F_AXIS); g.setColor(TEXT_CLR); cStr(g, xlabel, ML + pW / 2, pB + 45);
        AffineTransform ot = g.getTransform(); g.rotate(-Math.PI / 2); cStr(g, ylabel, -(MT + pH / 2), ML - 65); g.setTransform(ot);

        // Legend
        int lx = ML + pW - 280, ly = MT + 15, lH = arms.size() * 22 + 16;
        g.setColor(new Color(22, 48, 96, 200)); g.fillRoundRect(lx, ly, 260, lH, 8, 8);
        g.setColor(new Color(74, 106, 154, 100)); g.drawRoundRect(lx, ly, 260, lH, 8, 8);
        int ty = ly + 16;
        for (int i = 0; i < arms.size(); i++) {
            Map<String, Object> arm = arms.get(i); Color c = armClr(arm, i);
            String name = s(arm, "name", "Arm " + (i + 1));
            Number med = (Number) arm.get("median");
            g.setColor(c); g.setStroke(new BasicStroke(2.5f)); g.drawLine(lx + 12, ty, lx + 40, ty);
            g.setFont(F_LEGB); g.setColor(TEXT_CLR);
            String label = name + (med != null ? " (med: " + fmtN(med.doubleValue()) + " mo)" : "");
            if (label.length() > 35) label = label.substring(0, 32) + "...";
            g.drawString(label, lx + 48, ty + 4); ty += 22;
        }

        // HR box
        if (!hrText.isEmpty()) {
            String[] hrl = hrText.split(";\\s*");
            g.setFont(F_HR); FontMetrics fm = g.getFontMetrics();
            int bw = 0; for (String l : hrl) bw = Math.max(bw, fm.stringWidth(l.trim()));
            bw += 24; int bh = hrl.length * 18 + 14;
            int bx = ML + 15, by = pB - bh - 15;
            g.setColor(new Color(22, 48, 96, 220)); g.fillRoundRect(bx, by, bw, bh, 6, 6);
            g.setColor(ThemeConfig.CLR_TEAL); g.setStroke(new BasicStroke(1.5f)); g.drawLine(bx, by, bx, by + bh);
            int hty = by + 16;
            for (int i = 0; i < hrl.length; i++) {
                g.setFont(i == 0 ? F_HRB : F_HR); g.setColor(i == 0 ? ThemeConfig.CLR_TEAL : TEXT_CLR);
                g.drawString(hrl[i].trim(), bx + 10, hty); hty += 18;
            }
        }

        // At-risk table
        if (showAR) {
            g.setFont(F_ARH); g.setColor(TEXT_CLR); g.drawString("No. at risk", ML - 100, pB + MB_PLOT - 8);
            for (int a = 0; a < arms.size(); a++) {
                Map<String, Object> arm = arms.get(a); Color c = armClr(arm, a);
                List<Number> ar = nl(arm, "atRisk");
                int ry = pB + MB_PLOT + 7 + a * 20;
                g.setFont(F_AR); g.setColor(c);
                String sn = s(arm, "name", "Arm"); if (sn.length() > 18) sn = sn.substring(0, 15) + "...";
                rStr(g, sn, ML - 10, ry + 4);
                g.setColor(TEXT_CLR);
                int idx = 0; for (double x = 0; x <= xMax && idx < ar.size(); x += xStep) {
                    cStr(g, String.valueOf(ar.get(idx).intValue()), xPx(x, pW, xMax), ry + 4); idx++;
                }
            }
        }

        g.dispose(); return toPng(img);
    }

    // ═════════════════════════════════════════════
    // 2. BAR CHART (Java2D — no Python equivalent yet)
    // ═════════════════════════════════════════════
    @SuppressWarnings("unchecked")
    public byte[] generateBarChart(Map<String, Object> data) throws Exception {
        List<Map<String, Object>> bars = (List<Map<String, Object>>) data.getOrDefault("bars", List.of());
        String ylabel = s(data, "ylabel", "Response Rate (%)");
        boolean showVal = b(data, "showValues", true);
        int nn = bars.size(), pW = W - ML - MR, pH = H - MT - 100, pB = MT + pH;
        double yMax = 0; for (Map<String, Object> bb : bars) yMax = Math.max(yMax, n(bb, "value", 0));
        yMax = Math.ceil(yMax / 10) * 10 + 10; if (yMax > 100) yMax = 110;
        double bW = Math.min(80, (double) pW / nn * 0.6), gap = (pW - nn * bW) / (nn + 1);

        BufferedImage img = new BufferedImage(W, H, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = img.createGraphics(); aa(g);
        g.setColor(BG); g.fillRect(0, 0, W, H);
        g.setColor(GRID); g.setStroke(new BasicStroke(0.5f));
        for (double y = 10; y < yMax; y += 10) { int py = (int)(pB - (y / yMax) * pH); g.drawLine(ML, py, ML + pW, py); }

        for (int i = 0; i < nn; i++) {
            Map<String, Object> bb = bars.get(i); double val = n(bb, "value", 0); Color c = barClr(bb, i);
            double bx = ML + gap + i * (bW + gap); int barH = (int)((val / yMax) * pH), by = pB - barH;
            g.setColor(c); g.fillRoundRect((int) bx, by, (int) bW, barH, 4, 4); g.fillRect((int) bx, by + barH - 4, (int) bW, 4);
            double ciL = n(bb, "ci_low", 0), ciH = n(bb, "ci_high", 0);
            if (ciL > 0 && ciH > 0) { int cx = (int)(bx + bW / 2); g.setColor(TEXT_CLR); g.setStroke(new BasicStroke(1.5f));
                int cyL = (int)(pB - (ciL / yMax) * pH), cyH = (int)(pB - (ciH / yMax) * pH);
                g.drawLine(cx, cyL, cx, cyH); g.drawLine(cx - 4, cyL, cx + 4, cyL); g.drawLine(cx - 4, cyH, cx + 4, cyH); }
            if (showVal) { g.setColor(TEXT_CLR); g.setFont(F_LEGB); cStr(g, String.format("%.0f%%", val), (int)(bx + bW / 2), by - 10); }
            g.setColor(TEXT_CLR); g.setFont(F_TICK); String lb = s(bb, "label", ""); if (lb.length() > 15) lb = lb.substring(0, 12) + "...";
            cStr(g, lb, (int)(bx + bW / 2), pB + 18);
        }
        g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f)); g.drawLine(ML, pB, ML + pW, pB); g.drawLine(ML, MT, ML, pB);
        g.setFont(F_TICK); for (double y = 0; y <= yMax; y += 20) { int py = (int)(pB - (y / yMax) * pH); g.drawLine(ML - 5, py, ML, py); rStr(g, String.format("%.0f", y), ML - 10, py + 4); }
        g.setFont(F_AXIS); g.setColor(TEXT_CLR); AffineTransform ot = g.getTransform(); g.rotate(-Math.PI / 2); cStr(g, ylabel, -(MT + pH / 2), ML - 70); g.setTransform(ot);
        g.dispose(); return toPng(img);
    }

    // ═════════════════════════════════════════════
    // 3. WATERFALL PLOT (Java2D — no Python equivalent yet)
    // ═════════════════════════════════════════════
    @SuppressWarnings("unchecked")
    public byte[] generateWaterfallPlot(Map<String, Object> data) throws Exception {
        List<Map<String, Object>> pts = (List<Map<String, Object>>) data.get("patients");
        String ylabel = s(data, "ylabel", "Best Change from Baseline (%)");
        double thPR = n(data, "threshold_pr", -30), thPD = n(data, "threshold_pd", 20);
        pts.sort(Comparator.comparingDouble(p -> n(p, "change", 0)));
        int nn = pts.size(), pW = W - ML - MR, pH = H - MT - 120, pB = MT + pH;
        double yMin = -100, yMax = 100;
        for (Map<String, Object> p : pts) { double v = n(p, "change", 0); yMin = Math.min(yMin, v - 5); yMax = Math.max(yMax, v + 5); }
        yMin = Math.floor(yMin / 20) * 20; yMax = Math.ceil(yMax / 20) * 20;
        double bW = Math.max(2, (double) pW / nn - 1);

        BufferedImage img = new BufferedImage(W, H, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = img.createGraphics(); aa(g); g.setColor(BG); g.fillRect(0, 0, W, H);
        int zY = wfY(0, pH, yMin, yMax);
        g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f)); g.drawLine(ML, zY, ML + pW, zY);

        // Threshold lines
        g.setStroke(new BasicStroke(1f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER, 10f, new float[]{8f, 4f}, 0f));
        g.setColor(new Color(52, 168, 83, 150)); int prY = wfY(thPR, pH, yMin, yMax); g.drawLine(ML, prY, ML + pW, prY);
        g.setFont(F_TICK); g.drawString(String.format("%.0f%%", thPR), ML + pW + 5, prY + 4);
        g.setColor(new Color(234, 67, 53, 150)); int pdY = wfY(thPD, pH, yMin, yMax); g.drawLine(ML, pdY, ML + pW, pdY);
        g.drawString(String.format("+%.0f%%", thPD), ML + pW + 5, pdY + 4);

        for (int i = 0; i < nn; i++) {
            Map<String, Object> p = pts.get(i); double ch = n(p, "change", 0);
            String rsp = s(p, "response", ch <= thPR ? "PR" : ch >= thPD ? "PD" : "SD");
            Color c = rspClr(rsp); g.setColor(c);
            double bx = ML + i * ((double) pW / nn); int py = wfY(ch, pH, yMin, yMax);
            if (ch <= 0) g.fillRect((int) bx, zY, (int) Math.max(bW, 2), py - zY);
            else g.fillRect((int) bx, py, (int) Math.max(bW, 2), zY - py);
        }
        g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f)); g.drawLine(ML, MT, ML, pB);
        g.setFont(F_TICK); for (double y = yMin; y <= yMax; y += 20) { int py = wfY(y, pH, yMin, yMax); g.drawLine(ML - 5, py, ML, py); rStr(g, String.format("%.0f", y), ML - 10, py + 4); }
        g.setFont(F_AXIS); g.setColor(TEXT_CLR); AffineTransform ot = g.getTransform(); g.rotate(-Math.PI / 2); cStr(g, ylabel, -(MT + pH / 2), ML - 70); g.setTransform(ot);
        cStr(g, "Patients", ML + pW / 2, pB + 35);

        // Legend
        String[][] leg = {{"CR", "34A853"}, {"PR", "4285F4"}, {"SD", "FFA726"}, {"PD", "EA4335"}};
        int llx = ML + pW - 200, lly = MT + 10;
        g.setColor(new Color(22, 48, 96, 200)); g.fillRoundRect(llx, lly, 120, leg.length * 22 + 10, 8, 8);
        int lty = lly + 18; g.setFont(F_LEG);
        for (String[] it : leg) { g.setColor(ThemeConfig.hex(it[1])); g.fillRect(llx + 10, lty - 10, 14, 10); g.setColor(TEXT_CLR); g.drawString(it[0], llx + 30, lty); lty += 22; }

        g.dispose(); return toPng(img);
    }

    // ═════════════════════════════════════════════
    // 4. FOREST PLOT (Java2D — Python coming soon)
    // ═════════════════════════════════════════════
    @SuppressWarnings("unchecked")
    public byte[] generateForestPlot(Map<String, Object> data) throws Exception {
        List<Map<String, Object>> subs = (List<Map<String, Object>>) data.get("subgroups");
        String favD = s(data, "favorsDrug", "Favors Drug"), favC = s(data, "favorsControl", "Favors Control");
        int nn = subs.size(), rH = 28, pSX = W / 2 + 40, pW2 = W / 2 - 80;
        int imgH = Math.max(H, MT + nn * rH + 80);
        double hrMin = 0.1, hrMax = 3.0;

        BufferedImage img = new BufferedImage(W, imgH, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = img.createGraphics(); aa(g); g.setColor(BG); g.fillRect(0, 0, W, imgH);

        int hr1X = fPx(1.0, pSX, pW2, hrMin, hrMax);
        g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f)); g.drawLine(hr1X, MT, hr1X, MT + nn * rH);

        for (int i = 0; i < nn; i++) {
            if (i % 2 == 0) { g.setColor(new Color(22, 48, 96, 40)); g.fillRect(20, MT + i * rH, W - 40, rH); }
            Map<String, Object> sub = subs.get(i); double hr = n(sub, "hr", 1.0), ciL = n(sub, "ciLow", hr * 0.5), ciH = n(sub, "ciHigh", hr * 2.0);
            int ry = MT + i * rH + rH / 2;
            g.setFont(F_LEG); g.setColor(TEXT_CLR); String nm = s(sub, "name", "Subgroup"); if (nm.length() > 45) nm = nm.substring(0, 42) + "..."; g.drawString(nm, 30, ry + 4);
            g.setFont(F_TICK); rStr(g, String.format("%.2f (%.2f\u2013%.2f)", hr, ciL, ciH), pSX - 15, ry + 4);
            int x1 = fPx(ciL, pSX, pW2, hrMin, hrMax), x2 = fPx(ciH, pSX, pW2, hrMin, hrMax), hx = fPx(hr, pSX, pW2, hrMin, hrMax);
            Color lc = hr < 1.0 ? new Color(66, 133, 244) : new Color(234, 67, 53);
            g.setColor(lc); g.setStroke(new BasicStroke(2f)); g.drawLine(x1, ry, x2, ry);
            int ds = 5; g.fillPolygon(new int[]{hx - ds, hx, hx + ds, hx}, new int[]{ry, ry - ds, ry, ry + ds}, 4);
        }
        int xAY = MT + nn * rH + 10; g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f)); g.drawLine(pSX, xAY, pSX + pW2, xAY);
        g.setFont(F_TICK); for (double t : new double[]{0.2, 0.5, 1.0, 2.0}) { int tx = fPx(t, pSX, pW2, hrMin, hrMax); g.drawLine(tx, xAY, tx, xAY + 5); cStr(g, fmtN(t), tx, xAY + 20); }
        g.setFont(new Font("Calibri", Font.BOLD, 15)); g.setColor(new Color(66, 133, 244)); cStr(g, "\u2190 " + favD, pSX + pW2 / 4, xAY + 42);
        g.setColor(new Color(234, 67, 53)); cStr(g, favC + " \u2192", pSX + 3 * pW2 / 4, xAY + 42);

        g.dispose(); return toPng(img);
    }

    // ═════════════════════════════════════════════
    // 5. SWIMMER PLOT (Java2D — Python coming soon)
    // ═════════════════════════════════════════════
    @SuppressWarnings("unchecked")
    public byte[] generateSwimmerPlot(Map<String, Object> data) throws Exception {
        List<Map<String, Object>> pts = (List<Map<String, Object>>) data.get("patients");
        String xlabel = s(data, "xlabel", "Time (months)");
        pts.sort(Comparator.comparingDouble(p -> -n(p, "duration", 0)));
        int nn = Math.min(pts.size(), 40), barH = Math.max(12, Math.min(20, (H - MT - 100) / nn));
        int imgH = MT + nn * barH + 120, pW = W - ML - MR - 120;
        double xMax = 0; for (Map<String, Object> p : pts) xMax = Math.max(xMax, n(p, "duration", 0));
        xMax = Math.ceil(xMax / 6) * 6;

        BufferedImage img = new BufferedImage(W, Math.max(imgH, H), BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = img.createGraphics(); aa(g); g.setColor(BG); g.fillRect(0, 0, W, Math.max(imgH, H));

        g.setColor(GRID); g.setStroke(new BasicStroke(0.5f)); double xS = niceStep(xMax);
        for (double x = xS; x < xMax; x += xS) { int px = ML + (int)((x / xMax) * pW); g.drawLine(px, MT, px, MT + nn * barH); }

        for (int i = 0; i < nn; i++) {
            Map<String, Object> p = pts.get(i); double dur = n(p, "duration", 0); Color c = rspClr(s(p, "response", "SD"));
            int by = MT + i * barH + 2, bw = (int)((dur / xMax) * pW);
            g.setColor(c); g.fillRoundRect(ML, by, bw, barH - 4, 4, 4);
            g.setColor(AXIS_CLR); g.setFont(F_TICK); rStr(g, s(p, "id", String.valueOf(i + 1)), ML - 8, by + barH / 2 + 3);
            List<Map<String, Object>> evts = (List<Map<String, Object>>) p.getOrDefault("events", List.of());
            for (Map<String, Object> ev : evts) { double et = n(ev, "time", dur); int ex = ML + (int)((et / xMax) * pW), ey = by + barH / 2 - 2;
                String evT = s(ev, "type", "").toLowerCase();
                if (evT.contains("pd") || evT.contains("progress")) { g.setColor(new Color(234, 67, 53)); g.fillPolygon(new int[]{ex - 5, ex + 5, ex}, new int[]{ey + 5, ey + 5, ey - 5}, 3); }
                else if (evT.contains("death")) { g.setColor(Color.WHITE); g.setStroke(new BasicStroke(2f)); g.drawLine(ex - 4, ey - 4, ex + 4, ey + 4); g.drawLine(ex - 4, ey + 4, ex + 4, ey - 4); }
                else if (evT.contains("cr")) { g.setColor(new Color(52, 168, 83)); g.fillPolygon(new int[]{ex, ex + 5, ex, ex - 5}, new int[]{ey - 6, ey, ey + 6, ey}, 4); }
            }
        }
        int xAY = MT + nn * barH + 5; g.setColor(AXIS_CLR); g.setStroke(new BasicStroke(1.5f)); g.drawLine(ML, xAY, ML + pW, xAY);
        g.setFont(F_TICK); for (double x = 0; x <= xMax; x += xS) { int px = ML + (int)((x / xMax) * pW); g.drawLine(px, xAY, px, xAY + 5); cStr(g, fmtI(x), px, xAY + 20); }
        g.setFont(F_AXIS); g.setColor(TEXT_CLR); cStr(g, xlabel, ML + pW / 2, xAY + 42);

        // Legend
        String[][] leg = {{"CR", "34A853"}, {"PR", "4285F4"}, {"SD", "FFA726"}, {"PD", "EA4335"}};
        int llx = ML + pW + 20, lly = MT + 10; g.setFont(F_LEG);
        for (String[] it : leg) { g.setColor(ThemeConfig.hex(it[1])); g.fillRoundRect(llx, lly - 10, 20, 10, 3, 3); g.setColor(TEXT_CLR); g.drawString(it[0], llx + 26, lly); lly += 22; }

        g.dispose(); return toPng(img);
    }

    // ═════════════════════════════════════════════
    // UTILITIES
    // ═════════════════════════════════════════════
    private void aa(Graphics2D g) { g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON); g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON); g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY); }
    private byte[] toPng(BufferedImage img) throws Exception { ByteArrayOutputStream b = new ByteArrayOutputStream(); ImageIO.write(img, "png", b); return b.toByteArray(); }
    private int xPx(double x, int pW, double xMax) { return ML + (int)((x / xMax) * pW); }
    private int yPx(double y, int pH) { return MT + pH - (int)(y * pH); }
    private int wfY(double y, int pH, double yMin, double yMax) { return MT + (int)(((yMax - y) / (yMax - yMin)) * pH); }
    private int fPx(double hr, int sX, int pW, double hrMin, double hrMax) { double lMin = Math.log(hrMin), lMax = Math.log(hrMax); return sX + (int)(((Math.log(Math.max(hrMin, Math.min(hrMax, hr))) - lMin) / (lMax - lMin)) * pW); }
    private void cStr(Graphics2D g, String t, int x, int y) { FontMetrics fm = g.getFontMetrics(); g.drawString(t, x - fm.stringWidth(t) / 2, y); }
    private void rStr(Graphics2D g, String t, int x, int y) { FontMetrics fm = g.getFontMetrics(); g.drawString(t, x - fm.stringWidth(t), y); }
    private Color armClr(Map<String, Object> a, int i) { Object c = a.get("color"); if (c instanceof String) { String h = ((String) c).replace("#", ""); if (h.length() == 6) return ThemeConfig.hex(h); } return ARM_COLORS[i % ARM_COLORS.length]; }
    private Color barClr(Map<String, Object> b, int i) { return armClr(b, i); }
    private Color rspClr(String r) { return switch (r.toUpperCase()) { case "CR" -> new Color(52, 168, 83); case "PR" -> new Color(66, 133, 244); case "SD" -> new Color(255, 167, 38); case "PD" -> new Color(234, 67, 53); default -> AXIS_CLR; }; }
    private double svAt(List<Number> tp, List<Number> sv, double t) { double v = 1.0; for (int i = 0; i < Math.min(tp.size(), sv.size()); i++) { if (tp.get(i).doubleValue() <= t) v = sv.get(i).doubleValue(); else break; } return v; }
    private double autoXMax(List<Map<String, Object>> arms) { double m = 12; for (Map<String, Object> a : arms) for (Number t : nl(a, "timepoints")) m = Math.max(m, t.doubleValue()); return Math.ceil(m / 6) * 6; }
    private double niceStep(double r) { if (r <= 6) return 1; if (r <= 12) return 2; if (r <= 24) return 3; if (r <= 36) return 6; if (r <= 60) return 12; return Math.ceil(r / 6); }
    private String fmtN(double v) { return v == (int) v ? String.valueOf((int) v) : String.format("%.1f", v); }
    private String fmtI(double v) { return v == (int) v ? String.valueOf((int) v) : String.format("%.1f", v); }
    private String s(Map<String, Object> m, String k, String d) { Object v = m.get(k); return v != null ? v.toString() : d; }
    private double n(Map<String, Object> m, String k, double d) { Object v = m.get(k); if (v instanceof Number) return ((Number) v).doubleValue(); if (v instanceof String) { try { return Double.parseDouble((String) v); } catch (NumberFormatException e) { return d; } } return d; }
    private boolean b(Map<String, Object> m, String k, boolean d) { Object v = m.get(k); return v instanceof Boolean ? (Boolean) v : d; }
    @SuppressWarnings("unchecked") private List<Number> nl(Map<String, Object> m, String k) { Object v = m.get(k); return v instanceof List ? (List<Number>) v : List.of(); }

    @Deprecated public byte[] generateSwotMatrix(Map<String, Object> d) throws Exception {
        BufferedImage img = new BufferedImage(800, 400, BufferedImage.TYPE_INT_ARGB); Graphics2D g = img.createGraphics(); aa(g);
        g.setColor(BG); g.fillRect(0, 0, 800, 400); g.setFont(F_AXIS); g.setColor(TEXT_CLR); g.drawString("SWOT: use native POI", 50, 50); g.dispose(); return toPng(img);
    }
}
