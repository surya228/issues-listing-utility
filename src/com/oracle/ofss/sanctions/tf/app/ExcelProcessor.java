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
        try (FileInputStream fis = new FileInputStream("src/config.properties")) {
            config.load(fis);
        } catch (IOException e) {
            System.err.println("Error loading config.properties: " + e.getMessage());
            return;
        }

        String inputDir = config.getProperty("inputDirectory");
        String outputDir = config.getProperty("outputDirectory", inputDir);
        int statusCol1 = Integer.parseInt(config.getProperty("statusColumn1"));
        String statusCol2Str = config.getProperty("statusColumn2");
        Integer statusCol2 = statusCol2Str != null ? Integer.parseInt(statusCol2Str) : null;
        int ruleNameCol = Integer.parseInt(config.getProperty("ruleNameColumn"));

        // Ensure output directory exists
        try {
            Files.createDirectories(Paths.get(outputDir));
        } catch (IOException e) {
            System.err.println("Error creating output directory: " + e.getMessage());
            return;
        }

        // Collect unique FAIL entries
        Set<String> failsCol1 = new LinkedHashSet<>();
        Set<String> failsCol2 = statusCol2 != null ? new LinkedHashSet<>() : null;

        // Process input files
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(inputDir), "*.xlsx")) {
            for (Path filePath : stream) {
                processFile(filePath.toString(), statusCol1, statusCol2, ruleNameCol, failsCol1, failsCol2);
            }
        } catch (IOException e) {
            System.err.println("Error processing input directory: " + e.getMessage());
            return;
        }

        // Generate output filename
        SimpleDateFormat sdf = new SimpleDateFormat("ddMMyy_HHmmss");
        String timestamp = sdf.format(new Date());
        String outputFile = outputDir + File.separator + "Issues_" + timestamp + ".xlsx";

        // Write output
        writeOutput(outputFile, failsCol1, failsCol2);
        System.out.println("Output written to: " + outputFile);
    }

    private static void processFile(String filePath, int statusCol1, Integer statusCol2, int ruleNameCol, Set<String> failsCol1, Set<String> failsCol2) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assume first sheet
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header

                String ruleName = getCellValue(row, ruleNameCol);
                String status1 = getCellValue(row, statusCol1);

                if ("FAIL".equalsIgnoreCase(status1) && ruleName != null && !ruleName.trim().isEmpty()) {
                    failsCol1.add(ruleName.trim());
                }

                if (statusCol2 != null) {
                    String status2 = getCellValue(row, statusCol2);
                    if ("FAIL".equalsIgnoreCase(status2) && ruleName != null && !ruleName.trim().isEmpty()) {
                        failsCol2.add(ruleName.trim());
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

    private static void writeOutput(String outputFile, Set<String> failsCol1, Set<String> failsCol2) {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFile)) {

            // Sheet 1: StatusColumn1
            XSSFSheet sheet1 = workbook.createSheet("StatusColumn1");
            writeSheet(sheet1, failsCol1, "RuleName", "Status");

            // Sheet 2: StatusColumn2 (if applicable)
            if (failsCol2 != null) {
                XSSFSheet sheet2 = workbook.createSheet("StatusColumn2");
                writeSheet(sheet2, failsCol2, "RuleName", "Status");
            }

            workbook.write(fos);
        } catch (IOException e) {
            System.err.println("Error writing output file: " + e.getMessage());
        }
    }

    private static void writeSheet(XSSFSheet sheet, Set<String> data, String header1, String header2) {
        // Header
        XSSFRow headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue(header1);
        headerRow.createCell(1).setCellValue(header2);

        // Data
        int rowNum = 1;
        for (String ruleName : data) {
            XSSFRow row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(ruleName);
            row.createCell(1).setCellValue("FAIL");
        }

        // Auto-size columns
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
    }
}
