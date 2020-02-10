package com.hfahimi;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class App {

    public static String ROOT;
    public static final String SEP = "___";
    public static int CPU_CORES_COUNT = Runtime.getRuntime().availableProcessors();
    private static final Logger LOGGER = LoggerFactory.getLogger(App.class);

    // fill ROOT with jar file's directory

    static {
        try {
            ROOT = new File(App.class.getProtectionDomain().getCodeSource().getLocation().toURI() + File.separator)
                    .getParentFile()
                    .getPath()
                    .replace("file:" + File.separator, "") + File.separator;
        } catch (URISyntaxException e) {
            System.out.println("root not found : " + ROOT);
            System.exit(-1);
        }
    }



    public static void main(String[] args) throws InterruptedException {
        List<String> files = getFiles(Paths.get(ROOT), ".xlsx", ".docx");

        if (files == null) {
            System.exit(-1);
        }
        if (files.isEmpty()) {
            LOGGER.info("file(s) not found!");
            System.exit(0);
        }

        BlockingQueue<String> blockingQueue = new LinkedBlockingQueue<>(files);
        ExecutorService executorService = Executors.newFixedThreadPool(CPU_CORES_COUNT * 2);
        for (int i = 0; i < files.size(); i++) {
            executorService.execute(new Task(blockingQueue));
        }

        executorService.shutdown();
        executorService.awaitTermination(30, TimeUnit.SECONDS);

    }

    private static List<String> getFiles(Path directory, String... extensions) {
        try (Stream<Path> walk = java.nio.file.Files.walk(directory)) {
            List<String> files = walk.filter(java.nio.file.Files::isRegularFile)
                    .map(Path::toString).collect(Collectors.toList());

            return Objects.requireNonNull(files).stream()
                    .filter(
                            file ->
                                    Stream.of(extensions).anyMatch(
                                            ext -> file.toLowerCase().endsWith(ext)
                                    )
                    )
                    .collect(Collectors.toCollection(ArrayList::new));

        } catch (IOException e) {
            LOGGER.error("", e);
        }
        return null;
    }


}
