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
    public static final String PROP_THREAD_POOL_SIZE = "threads";
    public static final String PROP_EXTRACTION_ENABLED = "extractionEnabled";
    public static final String PROP_ANALYSIS_ENABLED = "analysisEnabled";
    public static final String PROP_OS_STATUS_FILTER = "osStatusFilter";
    public static final String PROP_OT_STATUS_FILTER = "otStatusFilter";
    public static final String JDBC_DRIVER = "jdbcdriver";
    public static final String JDBC_URL = "jdbcurl";
    public static final String WALLET_NAME = "walletName";
    public static final String TNS_ADMIN = "oracle.net.tns_admin";
    public static final String CONNECTION_ESTABLISHED = "Database connection established successfully";

    // Excel Column Names
    public static final String OS_TEST_STATUS = "OS Test Status";
    public static final String OT_TEST_STATUS = "OT Test Status";
    public static final String SOURCE_INPUT = "Source Input";
    public static final String N_UID = "N_UID";
    public static final String WATCHLIST = "Watchlist";
    public static final String TARGET_COLUMN = "Target Column";

    // Prefixes and Suffixes
    public static final String OS_PREFIX = "OS ";
    public static final String OT_PREFIX = "OT ";
    public static final String MESSAGE_PREFIX = "Message ";
    public static final String WEBSERVICE_PREFIX = " # ";
    public static final String MATCHES_SUFFIX = " matches";
    public static final String TRANSACTION_TOKEN_SUFFIX = " Transaction Token";

    // Status Values
    public static final String SWIFT = "SWIFT";
    public static final String FEDWIRE = "FEDWIRE";
    public static final String ISO20022 = "ISO20022";
    public static final String YES = "Yes";
    public static final String NO = "No";
    public static final String NA = "NA";
    public static final String INPUT_TO_MS = "Input to MS";
    public static final String CANDIDATES_PRESENT = "Candidates present";
    public static final String COMMENT = "Comment";

    // Query Strings
    public static final String CHECKER1_QUERY = "select c_request_json from FCC_MR_MATCHED_RESULT_RT WHERE N_REQUEST_ID = ? and c_request_json like ?";
    public static final String CHECKER2_QUERY_PREFIX = "select count(*) from rt_candidates where n_run_skey = (select n_run_skey from fcc_mr_matched_result_rt where rownum=1 and n_request_id=?) and V_WATCHLIST_TYPE = ? and n_uid=? and ";

    // Rule Names (preserving spaces)
    public static final String RULE_FULL_NAME_AND_ADDRESS = "\"ruleName\":\"Full Name And Address";
    public static final String RULE_IDENTIFIER = "\"ruleName\":\"Identifier";
    public static final String RULE_CITY_NAME = "\"ruleName\":\"City Name";
    public static final String RULE_COUNTRY_NAME = "\"ruleName\":\"Country Name";
    public static final String RULE_PORT_NAME = "\"ruleName\":\"Port Name";
    public static final String RULE_GOODS_NAME = "\"ruleName\":\"Goods Name";
    public static final String RULE_NARRATIVE_FULL_NAME = "\"ruleName\":\"Narrative Full Name";
    public static final String RULE_NARRATIVE_IDENTIFIER = "\"ruleName\":\"Narrative Identifier";
    public static final String RULE_NARRATIVE_CITY = "\"ruleName\":\"Narrative City";
    public static final String RULE_NARRATIVE_COUNTRY = "\"ruleName\":\"Narrative Country";
    public static final String RULE_NARRATIVE_PORT = "\"ruleName\":\"Narrative Port";
    public static final String RULE_NARRATIVE_GOODS = "\"ruleName\":\"Narrative Goods";
    public static final String RULE_STOP_KEYWORDS = "\"ruleName\":\"Stop Keywords";



    public static final String SCORING_ENGINE_ISSUE = "Scoring Engine Issue";
    public static final String TF_ISSUE = "TF Issue";
    public static final String MATCHING_ISSUE = "Matching Issue";
    public static final String OT_ISSUE = "Oracle Text Issue";

    // Zip bomb protection constants
    public static final long DEFAULT_MAX_EXCEL_FILE_SIZE = 100L * 1024 * 1024; // 100MB
    public static final double DEFAULT_MIN_INFLATE_RATIO = 0.0d; // 0%

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
