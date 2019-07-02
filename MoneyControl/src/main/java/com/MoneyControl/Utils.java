package com.MoneyControl;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;

public class Utils {

    public static HSSFSheet getSheet(){
        HSSFSheet sheet = null;
        try {
            String path = MoneyControlApplication.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
            InputStream input = new BufferedInputStream(new FileInputStream(path + "/sample.xls"));
            POIFSFileSystem fileSystem = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
            sheet = workbook.getSheetAt(0);
        } catch (Exception e) {
            System.out.println("Failed to getSheet from xls.");
        }
        return sheet;
    }
}
