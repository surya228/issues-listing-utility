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
        String osColStr = config.getProperty(Constants.PROP_OS_COL);
        String otColStr = config.getProperty(Constants.PROP_OT_COL);
        Integer osCol = osColStr != null ? Integer.parseInt(osColStr) : null;
        Integer otCol = otColStr != null ? Integer.parseInt(otColStr) : null;

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

        // Collect unique FAIL entries with full data
        Map<String, String[]> osData = osCol != null ? new LinkedHashMap<>() : null;
        Map<String, String[]> otData = otCol != null ? new LinkedHashMap<>() : null;

        // Process input files
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(inputDir), "*.xlsx")) {
            for (Path filePath : stream) {
//                processFile(filePath.toString(), osCol, otCol, osCols, otCols, osData, otData);
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
        writeOutput(outputFile, osData, otData);
        System.out.println("Output written to: " + outputFile);
    }

    private static void writeOutput(String outputFile, Map<String, String[]> osData, Map<String, String[]> otData) {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFile)) {

            workbook.write(fos);
        } catch (IOException e) {
            System.err.println("Error writing output file: " + e.getMessage());
        }
    }

    private static void writeSheet(XSSFSheet sheet, Map<String, String[]> data, String[] headers) {
        // Header
        XSSFRow headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        // Data
        int rowNum = 1;
        for (String[] rowData : data.values()) {
            XSSFRow row = sheet.createRow(rowNum++);
            for (int i = 0; i < rowData.length; i++) {
                row.createCell(i).setCellValue(rowData[i] != null ? rowData[i] : "");
            }
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            if(i == 1 || i == 2 || i == 13) continue;
            sheet.autoSizeColumn(i);
        }
    }

    private static int[] parseColumns(String colsStr) {
        if (colsStr == null || colsStr.trim().isEmpty()) return null;
        String[] parts = colsStr.split(",");
        int[] cols = new int[parts.length];
        for (int i = 0; i < parts.length; i++) {
            cols[i] = Integer.parseInt(parts[i].trim());
        }
        return cols;
    }
}
