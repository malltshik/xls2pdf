package ru.malltshik.xls2pdf.converter;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.malltshik.xls2pdf.helpers.HtmlHelper;
import ru.malltshik.xls2pdf.helpers.impl.HSSFHtmlHelper;
import ru.malltshik.xls2pdf.helpers.impl.XSSFHtmlHelper;
import ru.malltshik.xls2pdf.utils.ExcelUtils;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import static com.itextpdf.text.Element.*;
import static org.apache.poi.hssf.record.ExtendedFormatRecord.*;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;

public class Xls2HtmlConverter {

    private final Workbook wb;
    private final Appendable output;
    private final OutputStream target;
    private boolean completeHTML;
    private Formatter out;
    private boolean gotBounds;
    private int firstColumn;
    private int endColumn;
    private HtmlHelper helper;

    private static final String DEFAULTS_CLASS = "excelDefaults";
    private static final String COL_HEAD_CLASS = "colHeader";

    private static final Map<HorizontalAlignment, String> ALIGN = ExcelUtils.mapFor(ALIGN_LEFT, "left",
            ALIGN_CENTER, "center", ALIGN_RIGHT, "right",
            ALIGN_JUSTIFIED, "left", ALIGN_CENTER, "center");

    private static final Map<VerticalAlignment, String> VERTICAL_ALIGN = ExcelUtils.mapFor(
            VERTICAL_BOTTOM, "bottom", VERTICAL_CENTER, "middle", VERTICAL_TOP, "top");


    public OutputStream convert() throws IOException {
        printPage();
        return this.target;
    }

    public Xls2HtmlConverter(InputStream in, OutputStream out) {
        Objects.requireNonNull(in, "Input source must be non null");
        Objects.requireNonNull(out, "Output target must be non null");
        try {
            this.wb = WorkbookFactory.create(in);
            this.target = out;
            this.output = new OutputStreamWriter(this.target);
        } catch (IOException | InvalidFormatException e) {
            throw new IllegalArgumentException("Unable to initialize converter", e);
        }
        setupColorMap();
        completeHTML = true;
    }

    public Xls2HtmlConverter(File in, OutputStream out) throws FileNotFoundException {
        this(new FileInputStream(in), out);
    }

    private void setupColorMap() {
        if (wb instanceof HSSFWorkbook)
            helper = new HSSFHtmlHelper((HSSFWorkbook) wb);
        else if (wb instanceof XSSFWorkbook)
            helper = new XSSFHtmlHelper((XSSFWorkbook) wb);
        else
            throw new IllegalArgumentException("unknown workbook type: " + wb.getClass().getSimpleName());
    }

    private void printPage() throws IOException {
        try {
            ensureOut();
            if (completeHTML) {
                out.format("<?xml version=\"1.0\" encoding=\"utf-8\" ?>%n");
                out.format("<html>%n");
                out.format("<head>%n");
                out.format("<meta http-equiv=\"content-type\" content=\"application/xhtml+xml; charset=UTF-8\"/>%n");
                out.format("</head>%n");
                out.format("<body>%n");
            }

            print();

            if (completeHTML) {
                out.format("</body>%n");
                out.format("</html>%n");
            }
        } finally {
            if (out != null)
                out.close();
            if (output instanceof Closeable) {
                Closeable closeable = (Closeable) output;
                closeable.close();
            }
        }
    }

    private void print() {
        printInlineStyle();
        printSheets();
    }

    private void printInlineStyle() {
        out.format("<style type=\"text/css\">%n");
        printStyles();
        out.format("</style>%n");
    }

    private void ensureOut() {
        if (out == null)
            out = new Formatter(output);
    }

    private void printStyles() {
        ensureOut();
        BufferedReader in = null;
        try {
            in = new BufferedReader(new InputStreamReader(
                    this.getClass().getClassLoader().getResourceAsStream("excelStyle.css")
            ));
            String line;
            while ((line = in.readLine()) != null) {
                out.format("%s%n", line);
            }
        } catch (IOException e) {
            throw new IllegalStateException("Reading standard css", e);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    throw new IllegalStateException("Reading standard css", e);
                }
            }
        }

        Set<Short> seen = new HashSet<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            Iterator<Row> rows = sheet.rowIterator();
            while (rows.hasNext()) {
                Row row = rows.next();
                for (Cell cell : row) {
                    CellStyle style = cell.getCellStyle();
                    if (!seen.contains(style.getIndex())) {
                        printStyle(style);
                        seen.add(style.getIndex());
                    }
                }
            }
        }
    }

    private void printStyle(CellStyle style) {
        out.format(".%s .%s {%n", DEFAULTS_CLASS, styleName(style));
        styleContents(style);
        out.format("}%n");
    }

    private void styleContents(CellStyle style) {
        styleOut("text-align", style.getAlignmentEnum(), ALIGN);
        styleOut("vertical-align", style.getVerticalAlignmentEnum(), VERTICAL_ALIGN);
        fontStyle(style);
        helper.borderStyles(style, out);
        helper.colorStyles(style, out);
    }

    private void fontStyle(CellStyle style) {
        Font font = wb.getFontAt(style.getFontIndex());

        if (font.getBold())
            out.format("  font-weight: bold;%n");
        if (font.getItalic())
            out.format("  font-style: italic;%n");

        int fontheight = font.getFontHeightInPoints();
        if (fontheight == 9) {
            fontheight = 10;
        }
        out.format("  font-size: %dpt;%n", fontheight);

        if (!StringUtils.isEmpty(font.getFontName())) {
            out.format("  font-family: %s;%n", font.getFontName());
        }
    }

    private String styleName(CellStyle style) {
        if (style == null)
            style = wb.getCellStyleAt((short) 0);
        StringBuilder sb = new StringBuilder();
        try (Formatter fmt = new Formatter(sb)) {
            fmt.format("style_%02d", style.getIndex());
            return fmt.toString();
        }
    }

    private <K> void styleOut(String attr, K key, Map<K, String> mapping) {
        String value = mapping.get(key);
        if (value != null) {
            out.format("  %s: %s;%n", attr, value);
        }
    }

    private static CellType ultimateCellType(Cell c) {
        CellType type = c.getCellTypeEnum();
        if (type == FORMULA)
            type = c.getCachedFormulaResultTypeEnum();
        return type;
    }

    private void printSheets() {
        ensureOut();
        Sheet sheet = wb.getSheetAt(0);
        printSheet(sheet);
    }

    private void printSheet(Sheet sheet) {
        ensureOut();
        out.format("<table class=%s>%n", DEFAULTS_CLASS);
        printCols(sheet);
        printSheetContent(sheet);
        out.format("</table>%n");
    }

    private void printCols(Sheet sheet) {
        out.format("<col/>%n");
        ensureColumnBounds(sheet);
        for (int i = firstColumn; i < endColumn; i++) {
            out.format("<col/>%n");
        }
    }

    private void ensureColumnBounds(Sheet sheet) {
        if (gotBounds)
            return;

        Iterator<Row> iter = sheet.rowIterator();
        firstColumn = (iter.hasNext() ? Integer.MAX_VALUE : 0);
        endColumn = 0;
        while (iter.hasNext()) {
            Row row = iter.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
        gotBounds = true;
    }

    private void printColumnHeads() {
        out.format("<thead>%n");
        out.format("  <tr class=%s>%n", COL_HEAD_CLASS);
        out.format("    <th class=%s>&#x25CA;</th>%n", COL_HEAD_CLASS);
        // noinspection UnusedDeclaration
        StringBuilder colName = new StringBuilder();
        for (int i = firstColumn; i < endColumn; i++) {
            colName.setLength(0);
            int cnum = i;
            do {
                colName.insert(0, (char) ('A' + cnum % 26));
                cnum /= 26;
            } while (cnum > 0);
            out.format("    <th class=%s>%s</th>%n", COL_HEAD_CLASS, colName);
        }
        out.format("  </tr>%n");
        out.format("</thead>%n");
    }

    private static CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, int cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    private void printSheetContent(Sheet sheet) {
        //Print column header
        //printColumnHeads();

        int splitTopRow = -1;
        boolean isSplit = false;
        PaneInformation pi = sheet.getPaneInformation();
        if (pi != null && pi.getActivePane() > 1 && pi.isFreezePane()) {
            splitTopRow = pi.getHorizontalSplitTopRow();
        }

        if (splitTopRow < 0) {
            out.format("<tbody>%n");
        } else {
            out.format("<thead>%n");
            isSplit = true;
        }

        Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            Row row = rows.next();
            int rowNum = row.getRowNum();

            out.format("  <tr>%n");
            for (int i = firstColumn; i < endColumn; i++) {
                CellRangeAddress mergeRegion = getMergedRegion(sheet, rowNum, i);
                String spanCol = "";
                String spanRow = "";
                if (mergeRegion != null) {
                    int mergeCol = mergeRegion.getLastColumn() - mergeRegion.getFirstColumn();
                    int mergeRow = mergeRegion.getLastRow() - mergeRegion.getFirstRow();
                    if (mergeCol > 0) {
                        int colSpan = mergeCol + 1;
                        spanCol = " colspan=\"" + colSpan + "\"";
                    }
                    if (mergeRow > 0 && rowNum == mergeRegion.getFirstRow()) {
                        int rowSpan = mergeRow + 1;
                        spanRow = " rowspan=\"" + rowSpan + "\"";
                    } else if (mergeRow > 0) {
                        continue;
                    }
                }

                String content = "&nbsp;";
                String attrs = "";
                CellStyle style = null;
                if (i >= row.getFirstCellNum() && i < row.getLastCellNum()) {
                    Cell cell = row.getCell(i);
                    if (cell != null) {
                        style = cell.getCellStyle();
                        attrs = tagStyle(style);
                        CellFormat cf = CellFormat.getInstance(style.getDataFormatString());
                        CellFormatResult result = cf.apply(cell);
                        if (cell.getCellTypeEnum() == CellType.FORMULA) {
                            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                            CellValue cellValue = evaluator.evaluate(cell);
                            DecimalFormat df2 = new DecimalFormat("##0.000");

                            content = String.valueOf(df2.format(cellValue.getNumberValue()));
                        } else
                            content = result.text;

                        if (!StringUtils.isEmpty(content)) {
                            content = StringUtils.stripEnd(content, null);
                        }

                        if (content.equals(""))
                            content = "&nbsp;";
                        if (!StringUtils.isEmpty(content) && content.trim().equals("- 0"))
                            content = "-";
                    }
                }

                if (style != null && style.getRotation() == 90) {
                    out.format("    <td class=\"%s rotate\" %s %s><div><span>%s</span></div></td>%n",
                            styleName(style), attrs, spanCol + spanRow, content);
                } else
                    out.format("    <td class=%s %s %s>%s</td>%n",
                            styleName(style), attrs, spanCol + spanRow, content);

                if (mergeRegion != null) {
                    int col = mergeRegion.getLastColumn() - mergeRegion.getFirstColumn();
                    i += col;
                }

                if (style != null) {

                }
            }
            out.format("  </tr>%n");

            if (isSplit && rowNum == splitTopRow - 1) {
                out.format("</thead>%n");
                out.format("<tbody>%n");
            }

        }
        out.format("</tbody>%n");
    }

    private String tagStyle(CellStyle style) {

        switch (style.getAlignmentEnum()) {
            case RIGHT:
                return "style=\"text-align: right;\"";
            case CENTER:
            case CENTER_SELECTION:
                return "style=\"text-align: center;\"";
            case GENERAL:
            case LEFT:
            case JUSTIFY:
            case FILL:
            case DISTRIBUTED:
            default:
                return "style=\"text-align: left;\"";

        }
    }
}
