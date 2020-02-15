package com.hfahimi;

import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.eventfilesystem.POIFSReader;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Optional;

public class Word extends Document {
    private POIXMLDocument docx = null;
    private SummaryInformation doc;

    public Word(String file) throws IOException {
        if (file.endsWith(".doc")) {
            POIFSReader r = new POIFSReader();
            r.registerListener(poifsReaderEvent -> {
                        SummaryInformation si = null;
                        try {
                            doc = (SummaryInformation) PropertySetFactory.create(poifsReaderEvent.getStream());
                        } catch (NoPropertySetStreamException | MarkUnsupportedException | IOException e) {
                            e.printStackTrace();
                        }
                    },
                    "\005SummaryInformation");
            r.read(new FileInputStream(file));
        } else if (file.endsWith(".docx")) {
            docx = new XWPFDocument(Files.newInputStream(Paths.get(file)));
        }
        else {
            throw new IOException("invalid type");
        }
    }

    public String getAuthor() {
        if (docx != null) {
            try {
                return docx.getPackage().getPackageProperties().getCreatorProperty().orElse("UNKNOWN");
            } catch (InvalidFormatException e) {
                return "UNKNOWN";
            }
        }
        else {
            return Optional.ofNullable(doc.getAuthor()).orElse("UNKNOWN");
        }
    }

    public String getLastModifier() {
        if (docx != null) {
            try {
                return docx.getPackage().getPackageProperties().getLastModifiedByProperty().orElse("UNKNOWN");
            } catch (InvalidFormatException e) {
                return "UNKNOWN";
            }
        }
        else {
            return Optional.ofNullable(doc.getLastAuthor()).orElse("UNKNOWN");
        }
    }

    public void close() {
    }

}
