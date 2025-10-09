package com.oracle.ofss.sanctions.tf.app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelProcessor {

    public static void main(String[] args) {
        Properties config = new Properties();
        try (FileInputStream fis = new FileInputStream(Constants.CONFIG_FILE_PATH)) {
            config.load(fis);
        } catch (IOException e) {
            System.err.println("Error loading config.properties: " + e.getMessage());
            return;
        }

        String inputDir = config.getProperty(Constants.PROP_INPUT_DIR);
        String outputDir = config.getProperty(Constants.PROP_OUTPUT_DIR, inputDir);


        if (osCol == null && otCol == null) {
            System.err.println("Error: At least one of TestStatusOS.column or TestStatusOT.column must be provided.");
            return;
        }

        // Ensure output directory exists
        try {
            Files.createDirectories(Paths.get(outputDir));
        } catch (IOException e) {
            System.err.println("Error creating output directory: " + e.getMessage());
            return;
        }


        // Process input files
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(inputDir), "*.xlsx")) {
            for (Path filePath : stream) {
                processFile();
            }
        } catch (IOException e) {
            System.err.println("Error processing input directory: " + e.getMessage());
            return;
        }

        // Generate output filename
        SimpleDateFormat sdf = new SimpleDateFormat(Constants.DATE_FORMAT);
        String timestamp = sdf.format(new Date());
        String outputFile = outputDir + File.separator + Constants.OUTPUT_PREFIX + timestamp + Constants.EXTENSION;

        // Write output
        writeOutput();
        System.out.println("Output written to: " + outputFile);
    }

}
