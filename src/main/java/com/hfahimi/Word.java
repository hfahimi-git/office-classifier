package com.hfahimi;

import org.apache.poi.POIDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Optional;

public class Word extends Document {
    private POIXMLDocument office = null;
    private POIDocument docOffice = null;

    public Word(String file) throws IOException {
        if (file.endsWith(".docx"))
            office = new XWPFDocument(Files.newInputStream(Paths.get(file)));
        else if (file.endsWith(".doc"))
            docOffice = new HSSFWorkbook(Files.newInputStream(Paths.get(file)));
    }

    public String getAuthor() {
        try {
            if (office != null)
                return office.getPackage().getPackageProperties().getCreatorProperty().orElse("UNKNOWN");
            else if (docOffice != null)
                return Optional.ofNullable(docOffice.getSummaryInformation().getAuthor()).orElse("UNKNOWN");
        } catch (InvalidFormatException ignored) {
        }
        return "UNKNOWN";
    }

    public void close() throws IOException {
        if (office != null)
            office.close();
        else if(docOffice != null)
            docOffice.close();
    }

    public String getLastModifier() {
        try {
            if (office != null)
                return office.getPackage().getPackageProperties().getLastModifiedByProperty().orElse("UNKNOWN");
            else if (docOffice != null)
                return Optional.ofNullable(docOffice.getSummaryInformation().getLastAuthor()).orElse("UNKNOWN");
        } catch (InvalidFormatException ignored) {
        }
        return "UNKNOWN";
    }
}
