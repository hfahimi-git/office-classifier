package com.hfahimi;

import org.apache.poi.POIDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Optional;

public class Excel extends Document {
    private POIXMLDocument office = null;
    private POIDocument docOffice = null;

    public Excel(String file) throws IOException {
        if (file.endsWith(".xlsx"))
            office = new XSSFWorkbook(Files.newInputStream(Paths.get(file)));
        else if (file.endsWith(".xls"))
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
