package com.hfahimi;

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
            Document office = Document.get(file);
            LOGGER.info("{} just opened", file);
            String folder = office.getAuthor() + App.SEP + office.getLastModifier();
            office.close();
            LOGGER.info("{} file info, author" + App.SEP + "lastModifier: {}", file, folder);

            Path destFolder = Paths.get(App.ROOT + folder + File.separator);
            if(!Files.isDirectory(destFolder)) {
                Files.createDirectories(destFolder);
            }

            Files.copy(
                    Paths.get(file),
                    Paths.get(destFolder + File.separator + Paths.get(file).getFileName()),
                    StandardCopyOption.REPLACE_EXISTING
            );
            LOGGER.info("{} file just closed", file);
        } catch (IOException | InterruptedException e) {
            LOGGER.error("error {}", e.getMessage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
