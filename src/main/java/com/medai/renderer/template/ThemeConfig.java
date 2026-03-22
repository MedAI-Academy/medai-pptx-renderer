package com.medai.renderer.template;

import java.awt.Color;

/**
 * MedAI Brand Design System — colors, fonts, and spacing constants.
 * All slide builders use this single source of truth.
 *
 * v2.0 — Prezent-Premium Upgrade:
 * - SWOT quadrant colors (deep saturated)
 * - Zebra-stripe table colors
 * - Confidence badge colors (green/yellow/red)
 * - Icon circle colors for CONTENT_CARDS
 * - Lighter surface variant for more visual depth
 */
public class ThemeConfig {

    // ═══════════════════════════════════════════════════════════
    // COLORS (as java.awt.Color for JFreeChart + as hex for POI)
    // ═══════════════════════════════════════════════════════════

    // Primary backgrounds
    public static final String HEX_DARK    = "0B1A3B";  // Title, dividers
    public static final String HEX_NAVY    = "0D2B4E";  // Content dark BG
    public static final String HEX_SURFACE = "163060";  // Cards, boxes
    public static final String HEX_SURFACE_LT = "1C3A6E"; // Cards lighter (hover/depth)
    public static final String HEX_LIGHT   = "F0F4FF";  // Content light BG

    // Accent colors
    public static final String HEX_ACCENT  = "7C6FFF";  // Purple accent
    public static final String HEX_TEAL    = "22D3A5";  // Positive, success
    public static final String HEX_GOLD    = "F5C842";  // Warnings, KPIs
    public static final String HEX_ROSE    = "FF5F7E";  // Negative, alerts
    public static final String HEX_ORANGE  = "FF9F43";  // Secondary accent

    // Text colors
    public static final String HEX_WHITE   = "FFFFFF";
    public static final String HEX_TEXT    = "EAF0FF";  // Primary text on dark
    public static final String HEX_MUTED   = "7B9FD4";  // Secondary text
    public static final String HEX_DIM     = "4A6A9A";  // Tertiary text
    public static final String HEX_INK     = "1E293B";  // Text on light BG
    public static final String HEX_SLATE   = "64748B";  // Muted on light BG

    // ── SWOT Quadrant Colors (deep, saturated, elegant) ──
    public static final String HEX_SWOT_S  = "0F4D3F";  // Strengths — deep teal
    public static final String HEX_SWOT_W  = "4A1942";  // Weaknesses — deep rose/plum
    public static final String HEX_SWOT_O  = "1A2F6F";  // Opportunities — deep blue
    public static final String HEX_SWOT_T  = "4A2F1A";  // Threats — deep amber/brown

    // ── Zebra-Stripe Table Colors ──
    public static final String HEX_ZEBRA_EVEN = "152C50";  // Slightly lighter than navy
    public static final String HEX_ZEBRA_ODD  = "0D2B4E";  // Same as navy (base)

    // ── Confidence Score Badge Colors ──
    public static final String HEX_CONF_GREEN  = "22C55E";  // Score ≥ 95%
    public static final String HEX_CONF_YELLOW = "F59E0B";  // Score 85–94%
    public static final String HEX_CONF_RED    = "EF4444";  // Score < 85%

    // ── Icon Circle Colors (for CONTENT_CARDS headers) ──
    public static final String HEX_ICON_TEAL   = "0D9488";  // Medical / Disease
    public static final String HEX_ICON_GOLD   = "D97706";  // Data / Statistics
    public static final String HEX_ICON_VIOLET = "7C3AED";  // Science / Molecular
    public static final String HEX_ICON_ROSE   = "DB2777";  // Narrative / Story
    public static final String HEX_ICON_ACCENT = "6366F1";  // Strategy / Target
    public static final String HEX_ICON_CYAN   = "0891B2";  // Timeline / Calendar

    // ── Unicode Icon Symbols (PPTX-safe via Segoe UI Symbol) ──
    public static final String ICON_MEDICAL  = "\u2720";  // ✠ Maltese Cross
    public static final String ICON_DATA     = "\u25A3";  // ▣ White Square with Square
    public static final String ICON_SCIENCE  = "\u2318";  // ⌘ Place of Interest
    public static final String ICON_BOOK     = "\u2756";  // ❖ Black Diamond Minus White X
    public static final String ICON_TARGET   = "\u25CE";  // ◎ Bullseye
    public static final String ICON_CALENDAR = "\u25A3";  // ▣ Square
    public static final String ICON_CHECK    = "\u2714";  // ✔ Check Mark
    public static final String ICON_WARNING  = "\u26A0";  // ⚠ Warning
    public static final String ICON_ARROW_UP = "\u25B2";  // ▲ Up Triangle
    public static final String ICON_CIRCLE   = "\u25CF";  // ● Filled Circle

    // Icon cycle for card headers (matches ACCENT_CYCLE order)
    public static final String[] ICON_CYCLE = {
        ICON_TARGET, ICON_DATA, ICON_SCIENCE, ICON_BOOK,
        ICON_MEDICAL, ICON_CALENDAR, ICON_CHECK, ICON_ARROW_UP
    };
    // Icon circle BG colors matching ICON_CYCLE
    public static final String[] ICON_BG_CYCLE = {
        HEX_ICON_ACCENT, HEX_ICON_GOLD, HEX_ICON_VIOLET, HEX_ICON_ROSE,
        HEX_ICON_TEAL, HEX_ICON_CYAN, HEX_ICON_ACCENT, HEX_ICON_TEAL
    };

    // java.awt.Color equivalents (for JFreeChart)
    public static final Color CLR_DARK    = hex(HEX_DARK);
    public static final Color CLR_NAVY    = hex(HEX_NAVY);
    public static final Color CLR_SURFACE = hex(HEX_SURFACE);
    public static final Color CLR_ACCENT  = hex(HEX_ACCENT);
    public static final Color CLR_TEAL    = hex(HEX_TEAL);
    public static final Color CLR_GOLD    = hex(HEX_GOLD);
    public static final Color CLR_ROSE    = hex(HEX_ROSE);
    public static final Color CLR_WHITE   = hex(HEX_WHITE);
    public static final Color CLR_TEXT    = hex(HEX_TEXT);
    public static final Color CLR_MUTED   = hex(HEX_MUTED);

    // Accent color array for cycling through charts/items
    public static final String[] ACCENT_CYCLE = {
        HEX_ACCENT, HEX_TEAL, HEX_GOLD, HEX_ROSE,
        HEX_ORANGE, "4DA6FF", "E879F9", "34D399"
    };

    // ═══════════════════════════════════════════════════════════
    // FONTS
    // ═══════════════════════════════════════════════════════════

    public static final String FONT_TITLE = "Calibri";
    public static final String FONT_BODY  = "Calibri";
    public static final String FONT_MONO  = "Consolas";
    public static final String FONT_ICON  = "Segoe UI Symbol";  // Fallback: Arial Unicode MS

    // Font sizes (in points)
    public static final double SIZE_TITLE     = 36.0;
    public static final double SIZE_SUBTITLE  = 20.0;
    public static final double SIZE_HEADING   = 18.0;
    public static final double SIZE_BODY      = 13.0;
    public static final double SIZE_SMALL     = 10.0;
    public static final double SIZE_CAPTION   = 8.0;
    public static final double SIZE_FOOTER    = 7.0;

    // ═══════════════════════════════════════════════════════════
    // SLIDE DIMENSIONS (inches — widescreen 13.33" × 7.5")
    // ═══════════════════════════════════════════════════════════

    public static final double SLIDE_W = 13.333;  // Widescreen width
    public static final double SLIDE_H = 7.5;     // Widescreen height

    public static final double MARGIN   = 0.50;  // Outer margin
    public static final double CONTENT_X = 0.50;
    public static final double CONTENT_W = 12.33; // SLIDE_W - 2*MARGIN
    public static final double HEADER_H  = 0.82;  // Header bar height
    public static final double FOOTER_Y  = 7.05;  // Footer Y position
    public static final double FOOTER_H  = 0.35;  // Footer height
    public static final double CONTENT_Y = 1.10;  // Content start Y (below header)
    public static final double CONTENT_H = 5.80;  // Available content height

    // ═══════════════════════════════════════════════════════════
    // HELPERS
    // ═══════════════════════════════════════════════════════════

    /** Convert 6-char hex string to java.awt.Color */
    public static Color hex(String hex) {
        if (hex == null || hex.length() < 6) return Color.WHITE;
        return new Color(
            Integer.parseInt(hex.substring(0, 2), 16),
            Integer.parseInt(hex.substring(2, 4), 16),
            Integer.parseInt(hex.substring(4, 6), 16)
        );
    }

    /** Convert 6-char hex + alpha (0-255) to Color with transparency */
    public static Color hexAlpha(String hex, int alpha) {
        if (hex == null || hex.length() < 6) return new Color(255, 255, 255, alpha);
        return new Color(
            Integer.parseInt(hex.substring(0, 2), 16),
            Integer.parseInt(hex.substring(2, 4), 16),
            Integer.parseInt(hex.substring(4, 6), 16),
            alpha
        );
    }

    /** Get background color for a layout type */
    public static String bgColorFor(String layout, String theme) {
        if (layout == null) return "dark".equals(theme) ? HEX_NAVY : HEX_LIGHT;
        return switch (layout) {
            case "TITLE", "DIVIDER", "CONFIDENCE" -> HEX_DARK;
            case "TOC", "REFERENCES" -> HEX_NAVY;
            default -> "dark".equals(theme) ? HEX_NAVY : HEX_LIGHT;
        };
    }

    /** Get text color based on background */
    public static String textColorFor(String bgHex) {
        return switch (bgHex) {
            case "F0F4FF", "FFFFFF" -> HEX_INK;
            default -> HEX_TEXT;
        };
    }

    /** Get confidence badge color based on score */
    public static String confColor(double score) {
        if (score >= 95) return HEX_CONF_GREEN;
        if (score >= 85) return HEX_CONF_YELLOW;
        return HEX_CONF_RED;
    }
}
