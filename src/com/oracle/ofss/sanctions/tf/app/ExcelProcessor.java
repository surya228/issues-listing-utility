package com.oracle.ofss.sanctions.tf.app;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.sql.*;
import java.sql.SQLTimeoutException;
import java.sql.SQLRecoverableException;

public class ExcelProcessor {

    private static final Logger log = LoggerFactory.getLogger(ExcelProcessor.class);
    static List<String> osHeaders;
    static List<String> otHeaders;
    static List<List<Object>> osData;
    static List<List<Object>> otData;
    static List<String> allHeaders;
    static List<List<Object>> filteredData;
    static int threadPoolSize;
    static Object headerLock = new Object();

    public static void main(String[] args) throws Exception {
        log.info("=============================================================");
        log.info("               ISSUES LISTING UTILITY STARTED                ");
        log.info("=============================================================");

        // Configure Apache POI zip security to prevent zip bomb attacks
        ZipSecureFile.setMinInflateRatio(Constants.DEFAULT_MIN_INFLATE_RATIO);
        ZipSecureFile.setMaxEntrySize(Constants.DEFAULT_MAX_EXCEL_FILE_SIZE);

        long executionStartMillis = System.currentTimeMillis();
        Properties config = new Properties();
        try (FileInputStream fis = new FileInputStream(Constants.CONFIG_FILE_PATH)) {
            config.load(fis);
        } catch (IOException e) {
            System.err.println("Error loading config.properties: " + e.getMessage());
            return;
        }

        String inputDir = config.getProperty(Constants.PROP_INPUT_DIR);
        String outputDir = config.getProperty(Constants.PROP_OUTPUT_DIR, inputDir);
        threadPoolSize = Integer.parseInt(config.getProperty(Constants.PROP_THREAD_POOL_SIZE, "4"));
        String analysisEnabled = config.getProperty(Constants.PROP_ANALYSIS_ENABLED, "Y");
        String extractionEnabled = config.getProperty(Constants.PROP_EXTRACTION_ENABLED, "N");
        String osStatusFilter = config.getProperty(Constants.PROP_OS_STATUS_FILTER, "PASS");
        String otStatusFilter = config.getProperty(Constants.PROP_OT_STATUS_FILTER, "FAIL");
        String filters = config.getProperty(Constants.PROP_FILTERS, "");

        // Generate output filename
        SimpleDateFormat sdf = new SimpleDateFormat(Constants.DATE_FORMAT);
        String timestamp = sdf.format(new Date());

        // Ensure output directory exists
        try {
            Files.createDirectories(Paths.get(outputDir));
        } catch (IOException e) {
            System.err.println("Error creating output directory: " + e.getMessage());
            return;
        }


        // init
        osHeaders = new ArrayList<>();
        otHeaders = new ArrayList<>();
        osData = Collections.synchronizedList(new ArrayList<>());
        otData = Collections.synchronizedList(new ArrayList<>());

        // Collect input files
        List<Path> files = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(inputDir), "*.xlsx")) {
            for (Path filePath : stream) {
                files.add(filePath);
            }
        } catch (IOException e) {
            System.err.println("Error processing input directory: " + e.getMessage());
            return;
        }

        // Process input files concurrently
        if("Y".equalsIgnoreCase(analysisEnabled)) {
            ExecutorService fileExecutor = Executors.newFixedThreadPool(threadPoolSize);
            List<Future<Void>> fileFutures = new ArrayList<>();
            for (Path filePath : files) {
                fileFutures.add(fileExecutor.submit(() -> {
                    try {
                        processFile(filePath);
                    } catch (Exception e) {
                        System.err.println("Error processing file " + filePath + ": " + e.getMessage());
                    }
                    return null;
                }));
            }
            for (Future<Void> f : fileFutures) {
                try {
                    f.get();
                } catch (Exception e) {
                    log.error("Error in file processing future: {}", e.getMessage());
                }
            }
            fileExecutor.shutdown();

            String outputFile = outputDir + File.separator + Constants.OUTPUT_PREFIX + timestamp + Constants.EXTENSION;

            // Write output
            writeOutput(outputFile);
            System.out.println("Output written to: " + outputFile);
        }

        if ("Y".equalsIgnoreCase(extractionEnabled)) {
            List<String[]> filterList = new ArrayList<>();
            if (!filters.isEmpty()) {
                // Parse filters: OS:PASS,OT:FAIL;OS:FAIL,OT:PASS
                String[] filterGroups = filters.split(";");
                for (String group : filterGroups) {
                    String[] parts = group.split(",");
                    if (parts.length == 2) {
                        String os = parts[0].substring(parts[0].indexOf(":") + 1);
                        String ot = parts[1].substring(parts[1].indexOf(":") + 1);
                        filterList.add(new String[]{os, ot});
                    }
                }
            } else {
                // Fallback to individual properties
                filterList.add(new String[]{osStatusFilter, otStatusFilter});
            }

            for (String[] filter : filterList) {
                String osFilter = filter[0];
                String otFilter = filter[1];
                processFilteredExtraction(files, osFilter, otFilter);
                String filteredOutputFile = outputDir + File.separator + "OS " + osFilter + " OT " + otFilter + " " + timestamp + Constants.EXTENSION;
                writeFilteredOutput(filteredOutputFile, osFilter, otFilter);
                System.out.println("Filtered output written to: " + filteredOutputFile);
            }
        }

        log.info("=============================================================");
        log.info("              ISSUES LISTING UTILITY COMPLETED               ");
        log.info("=============================================================");
        long executionEndMillis = System.currentTimeMillis();
        log.info("Total time taken by utility: {} seconds", (executionEndMillis - executionStartMillis) / 1000L);
    }

    static void processFile(Path filePath) {
        try (Workbook wb = WorkbookFactory.create(filePath.toFile())) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            List<String> allColumns = new ArrayList<>();
            Map<String, Integer> colIndices = new HashMap<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    String colName = cell.getStringCellValue();
                    allColumns.add(colName);
                    colIndices.put(colName, i);
                }
            }

            // build headers if first file
            synchronized(headerLock) {
                if (osHeaders.isEmpty()) {
                // OS headers: common + OS + "Input to MS" + "Comment"
                for (String col : allColumns) {
                    if (!col.startsWith(Constants.OT_PREFIX)) {  // include common and OS
                        osHeaders.add(col);
                    }
                }
                osHeaders.add(Constants.INPUT_TO_MS);
                osHeaders.add(Constants.COMMENT);

                // OT headers: common + OT + "Input to MS" + "Candidates present" + "Comment"
                for (String col : allColumns) {
                    if (!col.startsWith(Constants.OS_PREFIX)) {  // include common and OT
                        otHeaders.add(col);
                    }
                }
                otHeaders.add(Constants.INPUT_TO_MS);
                otHeaders.add(Constants.CANDIDATES_PRESENT);
                otHeaders.add(Constants.COMMENT);
                }
            }

            List<List<Object>> localOsData = Collections.synchronizedList(new ArrayList<>());
            List<List<Object>> localOtData = Collections.synchronizedList(new ArrayList<>());
            ExecutorService executor = Executors.newFixedThreadPool(threadPoolSize);
            List<Future<Void>> futures = new ArrayList<>();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                final int rowNum = r;
                futures.add(executor.submit(() -> {
                    try {
                        processRow(rowNum, sheet, colIndices, localOsData, localOtData);
                    } catch (Exception e) {
                        log.error("Error processing row " + rowNum + ": " + e.getMessage());
                    }
                    return null;
                }));
            }
            for (Future<Void> f : futures) {
                f.get();
            }
            executor.shutdown();
            osData.addAll(localOsData);
            otData.addAll(localOtData);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void processFilteredExtraction(List<Path> files, String osFilter, String otFilter) {
        allHeaders = new ArrayList<>();
        filteredData = Collections.synchronizedList(new ArrayList<>());
        ExecutorService fileExecutor = Executors.newFixedThreadPool(threadPoolSize);
        List<Future<Void>> fileFutures = new ArrayList<>();
        for (Path filePath : files) {
            fileFutures.add(fileExecutor.submit(() -> {
                try {
                    processFilteredFile(filePath, osFilter, otFilter);
                } catch (Exception e) {
                    System.err.println("Error processing file " + filePath + ": " + e.getMessage());
                }
                return null;
            }));
        }
        for (Future<Void> f : fileFutures) {
            try {
                f.get();
            } catch (Exception e) {
                log.error("Error in filtered file processing future: {}", e.getMessage());
            }
        }
        fileExecutor.shutdown();
    }

    static void processFilteredFile(Path filePath, String osFilter, String otFilter) {
        try (Workbook wb = WorkbookFactory.create(filePath.toFile())) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            List<String> allColumns = new ArrayList<>();
            Map<String, Integer> colIndices = new HashMap<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    String colName = cell.getStringCellValue();
                    allColumns.add(colName);
                    colIndices.put(colName, i);
                }
            }

            // build headers if first file
            synchronized(headerLock) {
                if (allHeaders.isEmpty()) {
                    allHeaders.addAll(allColumns);
                }
            }

            List<List<Object>> localFilteredData = Collections.synchronizedList(new ArrayList<>());
            ExecutorService executor = Executors.newFixedThreadPool(threadPoolSize);
            List<Future<Void>> futures = new ArrayList<>();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                final int rowNum = r;
                futures.add(executor.submit(() -> {
                    try {
                        processFilteredRow(rowNum, sheet, colIndices, localFilteredData, osFilter, otFilter);
                    } catch (Exception e) {
                        log.error("Error processing filtered row " + rowNum + ": " + e.getMessage());
                    }
                    return null;
                }));
            }
            for (Future<Void> f : futures) {
                f.get();
            }
            executor.shutdown();
            filteredData.addAll(localFilteredData);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void processFilteredRow(int r, Sheet sheet, Map<String, Integer> colIndices, List<List<Object>> localFilteredData, String osFilter, String otFilter) {
        Row row = sheet.getRow(r);
        if (row == null) return;
        String osStatus = getCellValue(row, colIndices.get(Constants.OS_TEST_STATUS));
        String otStatus = getCellValue(row, colIndices.get(Constants.OT_TEST_STATUS));
        if (!osFilter.equals(osStatus) || !otFilter.equals(otStatus)) return;

        List<Object> rowData = new ArrayList<>();
        for (String header : allHeaders) {
            rowData.add(getCellValue(row, colIndices.get(header)));
        }
        localFilteredData.add(rowData);
    }

    static String computeRequestId(Row row, Map<String, Integer> colIndices, String type) {
        String transactionToken = getCellValue(row, colIndices.get(type + Constants.TRANSACTION_TOKEN_SUFFIX));
        String messageType = "";

        for (String col : colIndices.keySet()) {
            if (col.startsWith(Constants.MESSAGE_PREFIX)) {
                String val = getCellValue(row, colIndices.get(col));
                if (val != null && !val.isEmpty()) {
                    messageType = col.substring((Constants.MESSAGE_PREFIX).length()).trim();
                    break;
                }
            }
        }
        int suffix = 0;
        if (Constants.SWIFT.equals(messageType)) suffix = 1;
        else if (Constants.FEDWIRE.equals(messageType)) suffix = 2;
        else if (Constants.ISO20022.equals(messageType)) suffix = 3;
        return transactionToken + suffix;
    }

    static void processRow(int r, Sheet sheet, Map<String, Integer> colIndices, List<List<Object>> localOsData, List<List<Object>> localOtData) {
        Row row = sheet.getRow(r);
        if (row == null) return;
        String osStatus = getCellValue(row, colIndices.get(Constants.OS_TEST_STATUS));
        String otStatus = getCellValue(row, colIndices.get(Constants.OT_TEST_STATUS));
        boolean isOsFail = Constants.FAIL_STATUS.equals(osStatus);
        boolean isOtFail = Constants.FAIL_STATUS.equals(otStatus);
        if (!isOsFail && !isOtFail) return;

        // compute requestId
        String requestId = null;
        if (isOsFail) {
            requestId = computeRequestId(row, colIndices, "OS");
        } else {
            requestId = computeRequestId(row, colIndices, "OT");
        }

        if (isOsFail) {
            String inputToMs = checker1(requestId, row, colIndices, "OS");
            List<Object> rowData = new ArrayList<>();
            for (String header : osHeaders) {
                if (Constants.INPUT_TO_MS.equals(header)) {
                    rowData.add(inputToMs);
                } else if (Constants.COMMENT.equals(header)) {
                    String comment = Constants.NO.equals(inputToMs) ? Constants.TF_ISSUE : Constants.MATCHING_ISSUE;
                    rowData.add(comment);
                } else {
                    rowData.add(getCellValue(row, colIndices.get(header)));
                }
            }
            localOsData.add(rowData);
        }

        if (isOtFail) {
            String inputToMs = checker1(requestId, row, colIndices, "OT");
            String candidates;
            if (Constants.YES.equals(inputToMs)) {
                candidates = checker2(requestId, row, colIndices);
            } else {
                candidates = Constants.NA;
            }
            List<Object> rowData = new ArrayList<>();
            for (String header : otHeaders) {
                if (Constants.INPUT_TO_MS.equals(header)) {
                    rowData.add(inputToMs);
                } else if (Constants.CANDIDATES_PRESENT.equals(header)) {
                    rowData.add(candidates);
                } else if (Constants.COMMENT.equals(header)) {
                    String comment;
                    if (Constants.YES.equals(candidates)) {
                        comment = Constants.SCORING_ENGINE_ISSUE;
                    } else if (Constants.NO.equals(candidates)) {
                        comment = Constants.OT_ISSUE;
                    } else {
                        comment = Constants.TF_ISSUE;
                    }
                    rowData.add(comment);
                } else {
                    rowData.add(getCellValue(row, colIndices.get(header)));
                }
            }
            localOtData.add(rowData);
        }
    }

    static String checker1(String requestId, Row row, Map<String, Integer> colIndices, String type) {
        // find webservice
        String webservice = null;
        String prefix = type + Constants.WEBSERVICE_PREFIX;
        String suffix = Constants.MATCHES_SUFFIX;
        for (String col : colIndices.keySet()) {
            if (col.startsWith(prefix) && col.endsWith(suffix)) {
                String val = getCellValue(row, colIndices.get(col));
                if (val != null && !val.isEmpty()) {
                    webservice = col.substring(prefix.length(), col.length() - suffix.length());
                    break;
                }
            }
        }
        if (webservice == null) return Constants.NA;

        String searchText = getSearchText(webservice);
        String sourceInput = getCellValue(row, colIndices.get(Constants.SOURCE_INPUT));
        String query = Constants.CHECKER1_QUERY;
        try (Connection conn = SQLUtility.getDbConnection();
             PreparedStatement ps = conn.prepareStatement(query)) {
            ps.setString(1, requestId);
            ps.setString(2, "%" + searchText + "%");
            log.info("Executing checker1 query: {} with params: requestId={}, likePattern={}", query, requestId, "%" + searchText + "%");
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    String json = rs.getString(1);
                    if (json != null && json.contains(sourceInput)) {
                        return Constants.YES;
                    }
                }
            }
        } catch (SQLTimeoutException | SQLRecoverableException e) {
            log.error("Database timeout/recoverable error in checker1: {}", e.getMessage());
            return Constants.NA;
        } catch (Exception e) {
            log.error("Unexpected error in checker1: {}", e.getMessage(), e);
            return Constants.NA;
        }
        return Constants.NO;
    }

    static String getSearchText(String webservice) {
        switch (webservice) {
            case "NameAndAddress": return Constants.RULE_FULL_NAME_AND_ADDRESS;
            case "Identifier": return Constants.RULE_IDENTIFIER;
            case "City": return Constants.RULE_CITY_NAME;
            case "Country": return Constants.RULE_COUNTRY_NAME;
            case "Port": return Constants.RULE_PORT_NAME;
            case "Goods": return Constants.RULE_GOODS_NAME;
            case "Narrative NameAndAddress": return Constants.RULE_NARRATIVE_FULL_NAME;
            case "Narrative Identifier": return Constants.RULE_NARRATIVE_IDENTIFIER;
            case "Narrative City": return Constants.RULE_NARRATIVE_CITY;
            case "Narrative Country": return Constants.RULE_NARRATIVE_COUNTRY;
            case "Narrative Port": return Constants.RULE_NARRATIVE_PORT;
            case "Narrative Goods": return Constants.RULE_NARRATIVE_GOODS;
            case "Stopkeywords": return Constants.RULE_STOP_KEYWORDS;
            default: return "";
        }
    }

    static String checker2(String requestId, Row row, Map<String, Integer> colIndices) {
        String nUid = getCellValue(row, colIndices.get(Constants.N_UID));
        String watchlist = getCellValue(row, colIndices.get(Constants.WATCHLIST));
        String targetCol = getCellValue(row, colIndices.get(Constants.TARGET_COLUMN));
        String table = Constants.OT_TABLE_WL_MAP.get(watchlist);
        if (table == null) return Constants.NA;
        String query = Constants.CHECKER2_QUERY_PREFIX + targetCol + " is not null";
        try (Connection conn = SQLUtility.getDbConnection();
             PreparedStatement ps = conn.prepareStatement(query)) {
            ps.setString(1, requestId);
            ps.setString(2, table);
            ps.setString(3, nUid);
            log.info("Executing checker2 query: {} with params: requestId={}, nUid={}, table={}, targetCol={}", query, requestId, nUid, table, targetCol);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt(1) > 0 ? Constants.YES : Constants.NO;
                }
            }
        } catch (SQLTimeoutException | SQLRecoverableException e) {
            log.error("Database timeout/recoverable error in checker2: {}", e.getMessage());
            return Constants.NA;
        } catch (Exception e) {
            log.error("Unexpected error in checker2: {}", e.getMessage(), e);
            return Constants.NA;
        }
        return Constants.NO;
    }

    static String getCellValue(Row row, Integer colIndex) {
        if (colIndex == null) return "";
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";
        return cell.toString().trim();
    }

    static void writeOutput(String outputFile) {
        try (Workbook wb = new XSSFWorkbook()) {
            // OS sheet
            Sheet osSheet = wb.createSheet(Constants.SHEET_OS);
            Row osHeaderRow = osSheet.createRow(0);
            for (int i = 0; i < osHeaders.size(); i++) {
                osHeaderRow.createCell(i).setCellValue(osHeaders.get(i));
            }
            for (int r = 0; r < osData.size(); r++) {
                Row row = osSheet.createRow(r + 1);
                List<Object> data = osData.get(r);
                for (int c = 0; c < data.size(); c++) {
                    row.createCell(c).setCellValue(data.get(c).toString());
                }
            }

            // OT sheet
            Sheet otSheet = wb.createSheet(Constants.SHEET_OT);
            Row otHeaderRow = otSheet.createRow(0);
            for (int i = 0; i < otHeaders.size(); i++) {
                otHeaderRow.createCell(i).setCellValue(otHeaders.get(i));
            }
            for (int r = 0; r < otData.size(); r++) {
                Row row = otSheet.createRow(r + 1);
                List<Object> data = otData.get(r);
                for (int c = 0; c < data.size(); c++) {
                    row.createCell(c).setCellValue(data.get(c).toString());
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void writeFilteredOutput(String outputFile, String osFilter, String otFilter) {
        try (Workbook wb = new XSSFWorkbook()) {
            String sheetName = "OS " + osFilter + " OT " + otFilter;
            Sheet sheet = wb.createSheet(sheetName);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < allHeaders.size(); i++) {
                headerRow.createCell(i).setCellValue(allHeaders.get(i));
            }
            for (int r = 0; r < filteredData.size(); r++) {
                Row row = sheet.createRow(r + 1);
                List<Object> data = filteredData.get(r);
                for (int c = 0; c < data.size(); c++) {
                    row.createCell(c).setCellValue(data.get(c).toString());
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
