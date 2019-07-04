package com.MoneyControl.SheetOutput;

import com.MoneyControl.BankReport;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.sql.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

@Component
public class MomCredit {

    public static final String MOM_CREDIT_KEY = "momCredit";
    private final String SHEET_NAME = "Mom Credit";
    private static final String[] HEADERS = {"Date", "Description", "Credit"};

    public Workbook buildMomCredit(Map<String, List> expensesMap, Workbook workbook){
        List<BankReport> bankReportList = expensesMap.get(MOM_CREDIT_KEY);
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        sheet.createRow(0); Integer rowCount = 1;
        for(int i=0;i<bankReportList.size();i++){
            sheet.createRow(rowCount);
            rowCount++;
        }
        Iterator rowIterator = sheet.rowIterator();
        HSSFRow headerRow = (HSSFRow) rowIterator.next();
        for(int i=0;i<HEADERS.length;i++){
            headerRow.createCell(i).setCellValue(HEADERS[i]);
        }

        bankReportList.forEach((bankReport) -> {
            HSSFRow currentRow = (HSSFRow) rowIterator.next();
            HSSFCell dateCell = currentRow.createCell(0);
            HSSFCell descriptionCell = currentRow.createCell(1);
            HSSFCell moneyCell = currentRow.createCell(2);
            dateCell.setCellValue(Date.valueOf(bankReport.getDate()));
            descriptionCell.setCellValue(bankReport.getDescription());
            moneyCell.setCellValue(bankReport.getCredit());
        });
        for(int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }
        return workbook;
    }
}
