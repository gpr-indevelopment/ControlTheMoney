package com.MoneyControl.SheetBuilding;

import com.MoneyControl.Model.BankReport;
import com.MoneyControl.Utils.SheetUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.sql.Date;
import java.util.*;

@Component
public class SheetBuilder {

    @Getter @Setter
    public String salary;

    public static final String FIXED_EXPENSES_KEY = "Fixed Expenses";
    public static final String VARIABLE_EXPENSES_KEY = "Variable Expenses";
    public static final String MOM_EXPENSES_KEY = "Mom Expenses";
    public static final String MOM_CREDIT_KEY = "Mom Credit";
    public static final String[] HEADERS = {"Date", "Description", "Debit"};

    private final short BORDER = CellStyle.BORDER_THIN;
    private final short HEADER_ALIGN = CellStyle.ALIGN_CENTER;
    private final short HEADER_FONT_COLOR = HSSFColor.ROYAL_BLUE.index;

    public Workbook buildExpensesSheet(Map<String, List> expensesMap){
        Workbook workbook = new HSSFWorkbook();
        expensesMap.forEach((key, bankReportList) ->{
            Sheet sheet = workbook.createSheet(key);
            SheetUtils.createEmptyRows(sheet, bankReportList.size()+1);
            SheetUtils.buildAndStyleHeaders(sheet, HEADERS, BORDER, HEADER_ALIGN, HEADER_FONT_COLOR);
            buildSheetBody(sheet, bankReportList);
            SheetUtils.applyStylesToSheet(sheet, BORDER);
        });
        return workbook;
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
}
