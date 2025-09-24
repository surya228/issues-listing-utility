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
        int[] osCols = parseColumns(Constants.PROP_OS_COLS);
        int[] otCols = parseColumns(Constants.PROP_OT_COLS);

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
                processFile(filePath.toString(), osCol, otCol, osCols, otCols, osData, otData);
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

    private static void processFile(String filePath, Integer osCol, Integer otCol, int[] osCols, int[] otCols, Map<String, String[]> osData, Map<String, String[]> otData) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assume first sheet
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header

                if (osCol != null && osCols != null) {
                    String status = getCellValue(row, osCol);
                    if (Constants.FAIL_STATUS.equalsIgnoreCase(status)) {
                        String[] values = new String[16];
                        for (int i = 0; i < osCols.length; i++) {
                            values[i] = getCellValue(row, osCols[i]);
                        }
                        values[15] = Constants.FAIL_STATUS;
                        String ruleName = values[0];
                        if (ruleName != null && !ruleName.trim().isEmpty()) {
                            osData.put(ruleName.trim(), values);
                        }
                    }
                }

                if (otCol != null && otCols != null) {
                    String status = getCellValue(row, otCol);
                    if (Constants.FAIL_STATUS.equalsIgnoreCase(status)) {
                        String[] values = new String[16];
                        for (int i = 0; i < otCols.length; i++) {
                            values[i] = getCellValue(row, otCols[i]);
                        }
                        values[15] = Constants.FAIL_STATUS;
                        String ruleName = values[0];
                        if (ruleName != null && !ruleName.trim().isEmpty()) {
                            otData.put(ruleName.trim(), values);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Error processing file " + filePath + ": " + e.getMessage());
        }
    }

    private static String getCellValue(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }

    private static void writeOutput(String outputFile, Map<String, String[]> osData, Map<String, String[]> otData) {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFile)) {

            // Sheet for Open Search Issues (if applicable)
            if (osData != null && !osData.isEmpty()) {
                XSSFSheet sheet1 = workbook.createSheet(Constants.SHEET_OS);
                writeSheet(sheet1, osData, Constants.HEADERS);
            }

            // Sheet for Oracle Text Issues (if applicable)
            if (otData != null && !otData.isEmpty()) {
                XSSFSheet sheet2 = workbook.createSheet(Constants.SHEET_OT);
                writeSheet(sheet2, otData, Constants.HEADERS);
            }

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
