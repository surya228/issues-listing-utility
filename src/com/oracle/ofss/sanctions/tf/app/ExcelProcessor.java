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
import java.sql.*;

public class ExcelProcessor {

    static List<String> osHeaders;
    static List<String> otHeaders;
    static List<List<Object>> osData;
    static List<List<Object>> otData;
    static Connection conn;

    public static void main(String[] args) throws Exception {
        Properties config = new Properties();
        try (FileInputStream fis = new FileInputStream(Constants.CONFIG_FILE_PATH)) {
            config.load(fis);
        } catch (IOException e) {
            System.err.println("Error loading config.properties: " + e.getMessage());
            return;
        }

        String inputDir = config.getProperty(Constants.PROP_INPUT_DIR);
        String outputDir = config.getProperty(Constants.PROP_OUTPUT_DIR, inputDir);

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
        osData = new ArrayList<>();
        otData = new ArrayList<>();

        conn = SQLUtility.getDbConnection();

        // Process input files
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(inputDir), "*.xlsx")) {
            for (Path filePath : stream) {
                try {
                    processFile(filePath);
                } catch (Exception e) {
                    System.err.println("Error processing file " + filePath + ": " + e.getMessage());
                    // Continue with next file
                }
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
        writeOutput(outputFile);
        System.out.println("Output written to: " + outputFile);

        // close db
        try {
            if (conn != null) conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void processFile(Path filePath) {
        try (Workbook wb = WorkbookFactory.create(filePath.toFile())) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            Map<String, Integer> colIndices = new HashMap<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    colIndices.put(cell.getStringCellValue(), i);
                }
            }

            // build headers if first file
            if (osHeaders.isEmpty()) {
                // common
                for (String col : colIndices.keySet()) {
                    if (!col.startsWith("OS ") && !col.startsWith("OT ")) {
                        osHeaders.add(col);
                        otHeaders.add(col);
                    }
                }
                // OS specific
                for (String col : colIndices.keySet()) {
                    if (col.startsWith("OS ")) {
                        osHeaders.add(col);
                    }
                }
                osHeaders.add("Input to MS");
                // OT specific
                for (String col : colIndices.keySet()) {
                    if (col.startsWith("OT ")) {
                        otHeaders.add(col);
                    }
                }
                otHeaders.add("Input to MS");
                otHeaders.add("Candidates present");
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String osStatus = getCellValue(row, colIndices.get("OS Test Status"));
                String otStatus = getCellValue(row, colIndices.get("OT Test Status"));
                boolean isOsFail = "FAIL".equals(osStatus);
                boolean isOtFail = "FAIL".equals(otStatus);
                if (!isOsFail && !isOtFail) continue;

                // compute requestId
                String requestId = null;
                if (isOsFail) {
                    requestId = computeRequestId(row, colIndices, "OS");
                } else {
                    requestId = computeRequestId(row, colIndices, "OT");
                }

                List<Object> rowData = new ArrayList<>();
                // common
                for (String col : colIndices.keySet()) {
                    if (!col.startsWith("OS ") && !col.startsWith("OT ")) {
                        rowData.add(getCellValue(row, colIndices.get(col)));
                    }
                }

                if (isOsFail) {
                    // OS specific
                    for (String col : colIndices.keySet()) {
                        if (col.startsWith("OS ")) {
                            rowData.add(getCellValue(row, colIndices.get(col)));
                        }
                    }
                    String inputToMs = checker1(requestId, row, colIndices, "OS");
                    rowData.add(inputToMs);
                    osData.add(rowData);
                }

                if (isOtFail) {
                    // OT specific
                    for (String col : colIndices.keySet()) {
                        if (col.startsWith("OT ")) {
                            rowData.add(getCellValue(row, colIndices.get(col)));
                        }
                    }
                    String inputToMs = checker1(requestId, row, colIndices, "OT");
                    rowData.add(inputToMs);
                    String candidates = checker2(requestId, row, colIndices);
                    rowData.add(candidates);
                    otData.add(rowData);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static String computeRequestId(Row row, Map<String, Integer> colIndices, String type) {
        String transactionToken = getCellValue(row, colIndices.get(type + " Transaction Token"));
        String messageType = getCellValue(row, colIndices.get("Message Type"));
        int suffix = 0;
        if ("SWIFT".equals(messageType)) suffix = 1;
        else if ("FEDWIRE".equals(messageType)) suffix = 2;
        else if ("ISO20022".equals(messageType)) suffix = 3;
        return transactionToken + suffix;
    }

    static String checker1(String requestId, Row row, Map<String, Integer> colIndices, String type) {
        // find webservice
        String webservice = null;
        for (String col : colIndices.keySet()) {
            if (col.startsWith(type + " # ")) {
                String val = getCellValue(row, colIndices.get(col));
                if (val != null && !val.isEmpty()) {
                    webservice = col.substring((type + " # ").length());
                    break;
                }
            }
        }
        if (webservice == null) return "NA";

        String searchText = getSearchText(webservice);
        String sourceInput = getCellValue(row, colIndices.get("Source Input"));
        String query = "select c_request_json from FCC_MR_MATCHED_RESULT_RT WHERE N_REQUEST_ID = ? and c_request_json like ?";
        try (PreparedStatement ps = conn.prepareStatement(query)) {
            ps.setString(1, requestId);
            ps.setString(2, "%" + searchText + "%");
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    String json = rs.getString(1);
                    if (json != null && json.contains(sourceInput)) {
                        return "Yes";
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "No";
    }

    static String getSearchText(String webservice) {
        switch (webservice) {
            case "NameAndAddress": return "\"ruleName\":\"Full Name And Address\"";
            case "Identifier": return "\"ruleName\":\"Identifier\"";
            case "City": return "\"ruleName\":\"City Name\"";
            case "Country": return "\"ruleName\":\"Country Name\"";
            case "Port": return "\"ruleName\":\"Port Name\"";
            case "Goods": return "\"ruleName\":\"Goods Name\"";
            case "Narrative NameAndAddress": return "\"ruleName\":\"Narrative Full Name\"";
            case "Narrative Identifier": return "\"ruleName\":\"Narrative Identifier\"";
            case "Narrative City": return "\"ruleName\":\"Narrative City\"";
            case "Narrative Country": return "\"ruleName\":\"Narrative Country\"";
            case "Narrative Port": return "\"ruleName\":\"Narrative Port\"";
            case "Narrative Goods": return "\"ruleName\":\"Narrative Goods\"";
            case "Stopkeywords": return "\"ruleName\":\"Stop Keywords\"";
            default: return "";
        }
    }

    static String checker2(String requestId, Row row, Map<String, Integer> colIndices) {
        String nUid = getCellValue(row, colIndices.get("N_UID"));
        String watchlist = getCellValue(row, colIndices.get("Watchlist value"));
        String targetCol = getCellValue(row, colIndices.get("Target Column"));
        String table = Constants.TABLE_WL_MAP.get(watchlist);
        if (table == null) return "NA";
        String query = "select count(*) from rt_candidates where n_run_skey = (select n_run_skey from fcc_mr_matched_result_rt where rownum=1 and n_request_id=?) and n_uid=? and V_WATCHLIST_TYPE = ? and " + targetCol + " is not null";
        try (PreparedStatement ps = conn.prepareStatement(query)) {
            ps.setString(1, requestId);
            ps.setString(2, nUid);
            ps.setString(3, table);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt(1) > 0 ? "Yes" : "No";
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "No";
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

}
