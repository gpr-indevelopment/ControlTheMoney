package com.MoneyControl;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.sql.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

@Component
public class SheetBuilder {

    public static final String FIXED_EXPENSES_KEY = "Fixed Expenses";
    public static final String VARIABLE_EXPENSES_KEY = "Variable Expenses";
    public static final String MOM_EXPENSES_KEY = "Mom Expenses";
    public static final String MOM_CREDIT_KEY = "Mom Credit";
    public static final String[] HEADERS = {"Date", "Description", "Debit"};
    private final short BORDER = CellStyle.BORDER_THIN;

    public Workbook buildExpensesSheet(Map<String, List> expensesMap){
        Workbook workbook = new HSSFWorkbook();
        expensesMap.forEach((key, bankReportList) ->{
            Sheet sheet = workbook.createSheet(key);
            createEmptyRows(sheet, bankReportList.size()+1);
            buildAndStyleHeaders(sheet, HEADERS);
            buildSheetBody(sheet, bankReportList);
            applyStylesToSheet(sheet);
        });
        return workbook;
    }

    private void createEmptyRows(Sheet sheet, Integer size){
        for(int i=0;i<size;i++){
            sheet.createRow(i);
        }
    }

    public void buildAndStyleHeaders(Sheet sheet, String[] headers){
        Row headerRow = sheet.getRow(0);
        if(headerRow == null) headerRow = sheet.createRow(0);
        Font font = sheet.getWorkbook().createFont();
        font.setFontHeightInPoints((short) 13);
        font.setColor(HSSFColor.ROYAL_BLUE.index);
        font.setItalic(true);
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setFont(font);
        cellStyle.setBorderTop(BORDER);
        cellStyle.setBorderBottom(BORDER);
        cellStyle.setBorderLeft(BORDER);
        cellStyle.setBorderRight(BORDER);
        for(int i=0;i<headers.length;i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(cellStyle);
        }
    }

    private void buildSheetBody(Sheet sheet, List<BankReport> bankReportList){
        Iterator rowIterator = sheet.rowIterator();
        rowIterator.next();
        bankReportList.forEach((bankReport) -> {
            HSSFRow currentRow = (HSSFRow) rowIterator.next();
            HSSFCell dateCell = currentRow.createCell(0);
            HSSFCell descriptionCell = currentRow.createCell(1);
            HSSFCell moneyCell = currentRow.createCell(2);
            dateCell.setCellValue(Date.valueOf(bankReport.getDate()));
            descriptionCell.setCellValue(bankReport.getDescription());
            if(bankReport.isDebit()){
                moneyCell.setCellValue(bankReport.getDebit());
            }
            else if(!bankReport.isDebit()){
                sheet.getRow(0).getCell(2).setCellValue("Credit");
                moneyCell.setCellValue(bankReport.getCredit());
            }
        });
    }

    public void applyStylesToSheet(Sheet sheet){
        Integer columnCount = HEADERS.length;
        for(int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(BORDER);
        cellStyle.setBorderBottom(BORDER);
        cellStyle.setBorderLeft(BORDER);
        cellStyle.setBorderRight(BORDER);
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
}
