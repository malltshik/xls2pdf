package ru.malltshik.xls2pdf.converter;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.ElementList;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Objects;

import static com.itextpdf.tool.xml.XMLWorkerHelper.*;

public class Xls2PdfConverter {

    private static final Logger LOGGER = LoggerFactory.getLogger(Xls2PdfConverter.class);

    private final Xls2HtmlConverter xls2HtmlConverter;
    private final OutputStream target;

    public OutputStream convert() throws IOException, DocumentException {
        ByteArrayOutputStream htmlOutput = (ByteArrayOutputStream) xls2HtmlConverter.convert();
        // TODO dynamical scale
        Document doc = new Document(PageSize.A2);
        PdfWriter writer = PdfWriter.getInstance(doc, target);
        doc.open();
        getInstance().parseXHtml(writer, doc, new ByteArrayInputStream(htmlOutput.toByteArray()));
        doc.close();
        return target;
    }

    public Xls2PdfConverter(InputStream in, OutputStream out) {
        Objects.requireNonNull(in, "Input source must be non null");
        Objects.requireNonNull(out, "Output target must be non null");
        OutputStream htmlOutput = new ByteArrayOutputStream();
        this.xls2HtmlConverter = new Xls2HtmlConverter(in, htmlOutput);
        target = out;
    }

}
