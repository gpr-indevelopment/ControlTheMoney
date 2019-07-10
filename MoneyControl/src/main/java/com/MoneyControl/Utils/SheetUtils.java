package com.MoneyControl.Utils;

import com.MoneyControl.MoneyControlApplication;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

public class SheetUtils {

    public static final String EXCEL_SUM = "SUM";

    public static HSSFSheet getSheetFromLocalFile(String fileName){
        HSSFSheet sheet = null;
        try {
            String path = MoneyControlApplication.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
            InputStream input = new BufferedInputStream(new FileInputStream(path + fileName));
            POIFSFileSystem fileSystem = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
            sheet = workbook.getSheetAt(0);
        } catch (Exception e) {
            System.out.println("Failed to get sheet from xls.");
        }
        return sheet;
    }

    public static void buildAndStyleHeaders(Sheet sheet, String[] headers, short border, short headerAlign, short fontColor){
        Row headerRow = sheet.getRow(0);
        if(headerRow == null) headerRow = sheet.createRow(0);
        Font font = sheet.getWorkbook().createFont();
        font.setFontHeightInPoints((short) 13);
        font.setColor(fontColor);
        font.setItalic(true);
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(headerAlign);
        cellStyle.setFont(font);
        cellStyle.setBorderTop(border);
        cellStyle.setBorderBottom(border);
        cellStyle.setBorderLeft(border);
        cellStyle.setBorderRight(border);
        for(int i=0;i<headers.length;i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(cellStyle);
        }
    }

    public static void createEmptyRows(Sheet sheet, Integer size){
        for(int i=0;i<size;i++){
            sheet.createRow(i);
        }
    }

    public static void applyStylesToSheet(Sheet sheet, short border){
        short columnCount = sheet.getRow(0).getLastCellNum();
        for(int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(border);
        cellStyle.setBorderBottom(border);
        cellStyle.setBorderLeft(border);
        cellStyle.setBorderRight(border);
        Iterator rowIterator = sheet.rowIterator();
        rowIterator.next();
        while(rowIterator.hasNext()){
            Row currentRow = (Row) rowIterator.next();
            Iterator cellIterator = currentRow.cellIterator();
            while(cellIterator.hasNext()){
                Cell currentCell = (Cell) cellIterator.next();
                currentCell.setCellStyle(cellStyle);
            }
        }
    }

    public static String buildFormulaOfLastColumn(Sheet sheet, String excelFormula){
        StringBuilder stringBuilder = new StringBuilder(excelFormula + "(");
        String sheetName = "'" + sheet.getSheetName() + "'!";
        String referenceColumn = null;
        if(sheet.getRow(0) == null){
            new Exception("Formula builder: The first row of the sheet must exist for formula building.");
        }
        else{
            referenceColumn = CellReference.convertNumToColString(sheet.getRow(0).getLastCellNum()-1);
        }
        stringBuilder.append(sheetName).append(referenceColumn).append(sheet.getFirstRowNum()+1);
        stringBuilder.append(":");
        stringBuilder.append(sheetName).append(referenceColumn).append(sheet.getLastRowNum()+1);
        stringBuilder.append(")");
        return stringBuilder.toString();
    }
}
