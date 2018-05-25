package ru.malltshik.xls2pdf.converter;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.UUID;

import static java.lang.String.format;
import static org.hamcrest.CoreMatchers.*;
import static org.hamcrest.MatcherAssert.assertThat;

public class Xls2PdfConverterTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(Xls2HtmlConverterTest.class);

    @Test
    public void convert() throws Exception {
        String filename = format("%s/%s.pdf", System.getProperty("java.io.tmpdir"), UUID.randomUUID().toString());
        File target = new File(filename);

        LOGGER.debug("Test target file name is: {} ", filename);

        OutputStream result = new Xls2PdfConverter(getClass().getClassLoader().getResourceAsStream("test.xlsx"),
                new FileOutputStream(target)).convert();

        assertThat(result, notNullValue());
        assertThat(new String(Files.readAllBytes(target.toPath())), notNullValue());
        assertThat(new String(Files.readAllBytes(target.toPath())), not(equalTo("")));

//        LOGGER.debug("Target file {} going to be deleted", filename);
//        target.delete();
//        LOGGER.debug("Target file {} has bean deleted", filename);
    }

}