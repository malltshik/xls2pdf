package ru.malltshik.xls2pdf.helpers.impl;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.malltshik.xls2pdf.helpers.HtmlHelper;

import java.util.Formatter;

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

    //TODO implement this
    public void borderStyles(CellStyle style, Formatter out) {

    }
}