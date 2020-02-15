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
    private POIXMLDocument xlsx = null;
    private POIDocument xls = null;

    public Excel(String file) throws IOException {
        if (file.endsWith(".xlsx")) {
            xlsx = new XSSFWorkbook(Files.newInputStream(Paths.get(file)));
        }
        else if (file.endsWith(".xls")) {
            xls = new HSSFWorkbook(Files.newInputStream(Paths.get(file)));
        }
        else {
            throw new IOException("invalid type");
        }
    }

    public String getAuthor() {
        try {
            if (xlsx != null) {
                return xlsx.getPackage().getPackageProperties().getCreatorProperty().orElse("UNKNOWN");
            }
            else if (xls != null) {
                return Optional.ofNullable(xls.getSummaryInformation().getAuthor()).orElse("UNKNOWN");
            }
        } catch (InvalidFormatException ignored) {}
        return "UNKNOWN";
    }

    public void close() throws IOException {
        if (xlsx != null) {
            xlsx.close();
        }
        else if(xls != null) {
            xls.close();
        }
    }

    public String getLastModifier() {
        try {
            if (xlsx != null) {
                return xlsx.getPackage().getPackageProperties().getLastModifiedByProperty().orElse("UNKNOWN");
            }
            else if (xls != null) {
                return Optional.ofNullable(xls.getSummaryInformation().getLastAuthor()).orElse("UNKNOWN");
            }
        } catch (InvalidFormatException ignored) {}
        return "UNKNOWN";
    }
}
