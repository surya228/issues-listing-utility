package com.oracle.ofss.sanctions.tf.app;

import java.io.File;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class Constants {

    public static String CONFIG_FILE_NAME = "config";
    public static String BIN_FOLDER_NAME = "bin";
    public static String CURRENT_DIRECTORY = System.getProperty("user.dir");
    public static File PARENT_DIRECTORY = new File(CURRENT_DIRECTORY).getParentFile();
    public static String CONFIG_FILE_PATH = PARENT_DIRECTORY+File.separator+BIN_FOLDER_NAME+File.separator+CONFIG_FILE_NAME+".properties";
    public static final String FAIL_STATUS = "FAIL";
    public static final String SHEET_OS = "Open Search Analysis";
    public static final String SHEET_OT = "Oracle Text Analysis";
    public static final String OUTPUT_PREFIX = "Issues_";
    public static final String DATE_FORMAT = "ddMMyy_HHmmss";
    public static final String EXTENSION = ".xlsx";
    public static final String PROP_INPUT_DIR = "inputDirectory";
    public static final String PROP_OUTPUT_DIR = "outputDirectory";
    public static final String JDBC_DRIVER = "jdbcdriver";
    public static final String JDBC_URL = "jdbcurl";
    public static final String WALLET_NAME = "walletName";
    public static final String TNS_ADMIN = "oracle.net.tns_admin";
    public static final String CONNECTION_ESTABLISHED = "Database connection established successfully";

    public static final Map<String, String> OT_TABLE_WL_MAP;
    static {
        Map<String, String> map = new HashMap<>();
        map.put("COUNTRY", "FCC_TF_DIM_COUNTRY_OT");
        map.put("CITY", "FCC_TF_DIM_CITY_OT");
        map.put("GOODS", "FCC_TF_DIM_GOODS_OT");
        map.put("PORT", "FCC_TF_DIM_PORT_OT");
        map.put("STOP_KEYWORDS", "FCC_TF_DIM_STOPKEYWORDS_OT");
        map.put("IDENTIFIER", "FCC_DIM_IDENTIFIER_OT");
        map.put("WCPREM", "FCC_WL_WC_PREMIUM_OT");
        map.put("WCSTANDARD", "FCC_WL_WC_STANDARD_OT");
        map.put("DJW", "FCC_WL_DJW_OT");
        map.put("OFAC", "FCC_WL_OFAC_OT");
        map.put("HMT", "FCC_WL_HMT_OT");
        map.put("EU", "FCC_WL_EUROPEAN_UNION_OT");
        map.put("UN", "FCC_WL_UN_OT");
        map.put("PRV_WL1", "FCC_WL_PRIVATELIST_OT");

        OT_TABLE_WL_MAP = Collections.unmodifiableMap(map);
    }

}
