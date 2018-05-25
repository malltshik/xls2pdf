package ru.malltshik.xls2pdf.helpers.impl;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.malltshik.xls2pdf.helpers.HtmlHelper;

import java.util.Formatter;
import java.util.Map;

public class XSSFHtmlHelper implements HtmlHelper {
    private final XSSFWorkbook wb;

    private static final Map<Integer, HSSFColor> colors = HSSFColor.getIndexHash();

    public XSSFHtmlHelper(XSSFWorkbook wb) {
        this.wb = wb;
    }

    public void colorStyles(CellStyle style, Formatter out) {
        XSSFCellStyle cs = (XSSFCellStyle) style;
        styleColor(out, "background-color", cs.getFillForegroundXSSFColor());
        styleColor(out, "text-color", cs.getFont().getXSSFColor());
        styleColor(out, "color", cs.getFont().getXSSFColor());
    }

    private void styleColor(Formatter out, String attr, XSSFColor color) {
        if (color == null || color.isAuto())
            return;
        byte[] rgb = color.getRGBWithTint();
        if (rgb == null) {
            rgb = color.getRGB();
            if (rgb == null) {
                return;
            }
        }

        out.format("  %s: #%02x%02x%02x;%n", attr, rgb[0], rgb[1], rgb[2]);
        byte[] argb = color.getARGB();
        if (argb == null) {
            return;
        }
        out.format("  %s: rgba(0x%02x, 0x%02x, 0x%02x, 0x%02x);%n", attr,
                argb[3], argb[0], argb[1], argb[2]);
    }

    public void borderStyles(CellStyle style, Formatter out) {
        XSSFCellStyle xstyle = (XSSFCellStyle) style;
        styleOut("border-left", style.getBorderLeftEnum(), BORDER, xstyle.getLeftBorderXSSFColor(), out);
        styleOut("border-right", style.getBorderRightEnum(), BORDER, xstyle.getRightBorderXSSFColor(), out);
        styleOut("border-top", style.getBorderTopEnum(), BORDER, xstyle.getTopBorderXSSFColor(), out);
        styleOut("border-bottom", style.getBorderBottomEnum(), BORDER, xstyle.getBottomBorderXSSFColor(), out);
    }

    private <K> void styleOut(String attr, K key, Map<K, String> mapping, Color color, Formatter out) {
        String value = mapping.get(key);
        String c = borderStyle(color);
        value = String.format(value, c);
        if (value != null) {
            out.format("  %s: %s;%n", attr, value);
        }
    }

    private String borderStyle(Color color) {
        XSSFColor xcolor = (XSSFColor) color;
        if (xcolor != null && !xcolor.isAuto() && xcolor.getARGBHex() != null) {
            return "#" + xcolor.getARGBHex().substring(2);
        }
        //TODO return for color indexed
        if (xcolor != null && xcolor.getIndexed() > 0) {
            return AUTO_COLOR;
        }
        return AUTO_COLOR;
    }
}