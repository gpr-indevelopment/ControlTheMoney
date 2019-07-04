package com.MoneyControl;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;

public class Utils {

    public static final String ANSI_RESET = "\u001B[0m";
    public static final String ANSI_BLACK = "\u001B[30m";
    public static final String ANSI_RED = "\u001B[31m";
    public static final String ANSI_GREEN = "\u001B[32m";
    public static final String ANSI_YELLOW = "\u001B[33m";
    public static final String ANSI_BLUE = "\u001B[34m";
    public static final String ANSI_PURPLE = "\u001B[35m";
    public static final String ANSI_CYAN = "\u001B[36m";
    public static final String ANSI_WHITE = "\u001B[37m";

    public static HSSFSheet getSheet(String sheetName){
        HSSFSheet sheet = null;
        try {
            String path = MoneyControlApplication.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
            InputStream input = new BufferedInputStream(new FileInputStream(path + sheetName));
            POIFSFileSystem fileSystem = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
            sheet = workbook.getSheetAt(0);
        } catch (Exception e) {
            System.out.println("Failed to getSheet from xls.");
        }
        return sheet;
    }

    public static String formatBankReportDisplay(BankReport bankReport){
        String date = bankReport.getDate().toString();
        String description = bankReport.getDescription();
        String debit = String.valueOf(bankReport.getDebit());
        return ANSI_BLUE + "DATE: " + ANSI_RESET + date + " - " + ANSI_BLUE + "DESCRIPTION: " + ANSI_RESET + description + " - " + ANSI_BLUE + "DEBIT: " + ANSI_RESET + debit;
    }

    public static String trimDescription(String description){
        while(description.contains("  ")){
            description = description.replace("  ", " ");
        }
        return description;
    }
}
