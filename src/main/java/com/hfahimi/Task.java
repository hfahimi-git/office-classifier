package com.hfahimi;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.concurrent.BlockingQueue;

public class Task implements Runnable {
    private static final Logger LOGGER = LoggerFactory.getLogger(Task.class);
    private BlockingQueue<String> blockingQueue;

    public Task(BlockingQueue<String> blockingQueue) {
        this.blockingQueue = blockingQueue;
    }

    @Override
    public void run() {
        try {
            String file = blockingQueue.take();
            String folder = "UNKNOWN" + App.SEP + "UNKNOWN";
            POIXMLDocument office;
            if (file.endsWith(".docx")) {
                office = new XWPFDocument(Files.newInputStream(Paths.get(file)));
            } else if(file.endsWith(".xlsx")) {
                office = new XSSFWorkbook(Files.newInputStream(Paths.get(file)));
            }
            else {
                return;
            }
            LOGGER.info("{} just opened", file);

            try {
                String author = office.getPackage().getPackageProperties().getCreatorProperty().orElse("UNKNOWN");
                String lastModifier = office.getPackage().getPackageProperties().getLastModifiedByProperty().orElse("UNKNOWN");
                folder = author + App.SEP + lastModifier;
            } catch (InvalidFormatException e) {
                LOGGER.error("error {}", e.getMessage());
            }
            LOGGER.info("{} open file, author" + App.SEP + "lastModifier: {}", file, folder);

            Path destFolder = Paths.get(App.ROOT + folder + File.separator);
            if(!Files.isDirectory(destFolder)) {
                Files.createDirectories(destFolder);
            }
            office.close();

            Files.copy(
                    Paths.get(file),
                    Paths.get(destFolder + File.separator + Paths.get(file).getFileName()),
                    StandardCopyOption.REPLACE_EXISTING
            );
            LOGGER.info("{} file just closed", file);
        } catch (IOException | InterruptedException e) {
            LOGGER.error("error {}", e.getMessage());
        }
    }
}
