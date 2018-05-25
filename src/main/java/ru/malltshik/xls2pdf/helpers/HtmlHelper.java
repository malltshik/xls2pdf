package ru.malltshik.xls2pdf.helpers;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.malltshik.xls2pdf.utils.ExcelUtils;

import java.util.Formatter;
import java.util.Map;

import static org.apache.poi.ss.usermodel.BorderStyle.*;


public interface HtmlHelper {

    String AUTO_COLOR = "#000";

    Map<BorderStyle, String> BORDER = ExcelUtils.mapFor(
            DASH_DOT, "%s dashed 1pt", DASH_DOT_DOT, "%s dashed 1pt",
            DASHED, "%s dashed 1pt", DOTTED, "%s dotted 1pt",
            DOUBLE, "%s double 3pt", HAIR, "%s solid 1px",
            MEDIUM, "%s solid 2pt", MEDIUM_DASH_DOT, "%s dashed 2pt",
            MEDIUM_DASH_DOT_DOT, "%s dashed 2pt", MEDIUM_DASHED, "%s dashed 2pt",
            NONE, "none", SLANTED_DASH_DOT, "%s dashed 2pt",
            THICK, "%s solid 3pt", THIN, "%s solid 1pt");

    void colorStyles(CellStyle style, Formatter out);

    void borderStyles(CellStyle style, Formatter out);

}
