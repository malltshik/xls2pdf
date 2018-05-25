package ru.malltshik.xls2pdf.helpers.impl;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import ru.malltshik.xls2pdf.helpers.HtmlHelper;

import java.util.Formatter;
import java.util.Map;
import java.util.stream.Stream;

public class HSSFHtmlHelper implements HtmlHelper {

    private final HSSFWorkbook wb;
    private final HSSFPalette colors;

    private static final HSSFColorPredefined HSSF_AUTO = HSSFColorPredefined.AUTOMATIC;

    public HSSFHtmlHelper(HSSFWorkbook wb) {
        this.wb = wb;
        colors = wb.getCustomPalette();
    }

    public void colorStyles(CellStyle style, Formatter out) {
        HSSFCellStyle cs = (HSSFCellStyle) style;
        out.format("  /* fill pattern = %d */%n", cs.getFillPatternEnum().getCode());
        styleColor(out, "background-color", cs.getFillForegroundColor());
        styleColor(out, "color", cs.getFont(wb).getColor());
        styleColor(out, "border-left-color", cs.getLeftBorderColor());
        styleColor(out, "border-right-color", cs.getRightBorderColor());
        styleColor(out, "border-top-color", cs.getTopBorderColor());
        styleColor(out, "border-bottom-color", cs.getBottomBorderColor());
    }

    private void styleColor(Formatter out, String attr, short index) {
        HSSFColor color = colors.getColor(index);
        if (index == HSSF_AUTO.getIndex() || color == null) {
            out.format("  /* %s: index = %d */%n", attr, index);
        } else {
            short[] rgb = color.getTriplet();
            out.format("  %s: #%02x%02x%02x; /* index = %d */%n", attr, rgb[0],
                    rgb[1], rgb[2], index);
        }
    }

    public void borderStyles(CellStyle style, Formatter out) {
        HSSFCellStyle xstyle = (HSSFCellStyle) style;
        styleOut("border-left", style.getBorderLeftEnum(), BORDER, xstyle.getLeftBorderColor(), out);
        styleOut("border-right", style.getBorderRightEnum(), BORDER, xstyle.getRightBorderColor(), out);
        styleOut("border-top", style.getBorderTopEnum(), BORDER, xstyle.getTopBorderColor(), out);
        styleOut("border-bottom", style.getBorderBottomEnum(), BORDER, xstyle.getBottomBorderColor(), out);
    }

    private <K> void styleOut(String attr, K key, Map<K, String> mapping, short color, Formatter out) {
        String value = mapping.get(key);
        String c = borderStyle(color);
        value = String.format(value, c);
        if (value != null) {
            out.format("  %s: %s;%n", attr, value);
        }
    }

    private String borderStyle(short color) {
        HSSFColorPredefined xcolor = Stream.of(HSSFColorPredefined.values())
                .filter(x -> x.getIndex() == color).findFirst().orElse(HSSFColorPredefined.AUTOMATIC);
        if (!xcolor.equals(HSSFColorPredefined.AUTOMATIC) && xcolor.getHexString() != null) {
            return "#" + xcolor.getHexString().substring(2);
        }
        return AUTO_COLOR;
    }
}