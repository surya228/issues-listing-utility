package com.oracle.ofss.sanctions.tf.app;

public class Constants {

    public static final String CONFIG_FILE_PATH = "src/config.properties";
    public static final String FAIL_STATUS = "FAIL";
    public static final String SHEET_OS = "Open Search Issues";
    public static final String SHEET_OT = "Oracle Text Issues";
    public static final String OUTPUT_PREFIX = "Issues_";
    public static final String DATE_FORMAT = "ddMMyy_HHmmss";
    public static final String EXTENSION = ".xlsx";
    public static final String PROP_INPUT_DIR = "inputDirectory";
    public static final String PROP_OUTPUT_DIR = "outputDirectory";
    public static final String PROP_OS_COL = "TestStatusOS.column";
    public static final String PROP_OT_COL = "TestStatusOT.column";
    public static final String PROP_OS_COLS = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,16";
    public static final String PROP_OT_COLS = "1,2,3,4,5,6,7,8,17,18,19,20,21,22,24";
    public static final String PROP_RULE_NAME_COL = "ruleNameColumn";
    public static final String[] HEADERS = {
        "Rule Name", "Raw Message", "Tag", "Source Input", "Target Input",
        "Target Column", "Watchlist", "N_UID", "Transaction Token",
        "Match Count", "Status", "Feedback Status", "Specific Match Count",
        "Feedback", "Comments", "Test Status"
    };

}
