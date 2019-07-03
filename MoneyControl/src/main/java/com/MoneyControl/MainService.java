package com.MoneyControl;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

@Service
public class MainService {

    private final String SEARCH_REFERENCE = "Data";
    private final String STOP_REFERENCE = "TOTAL";
    private final String SHEET_NAME = "/sample.xls";

    public void init() {

        HSSFSheet sheet = Utils.getSheet(SHEET_NAME);
        HSSFCell referenceCell = findReferenceCell(sheet);
        List<BankReport> bankReportList = fillBankReports(referenceCell, sheet);
        routeExpenses(bankReportList);
    }

    public void routeExpenses(List<BankReport> bankReportsList) {
        Scanner scanner = new Scanner(System.in);
        List<BankReport> fixedExpenses = new ArrayList<>();
        List<BankReport> variableExpenses = new ArrayList<>();
        List<BankReport> momExpenses = new ArrayList<>();
        bankReportsList.forEach((report) -> {
            if (report.getDebit() != 0) {
                System.out.println(Utils.formatBankReportDisplay(report));
                System.out.println("Insert: f = fixed; v = variable; m = mom.");
                String line = scanner.nextLine();
                while (!(line.equals("f") || line.equals("v") || line.equals("m"))) {
                    System.out.println("Incorrect insert.");
                    System.out.println("Insert: f = fixed; v = variable; m = mom.");
                    line = scanner.nextLine();
                }
                switch (line) {
                    case ("f"): {
                        fixedExpenses.add(report);
                        break;
                    }
                    case ("v"): {
                        variableExpenses.add(report);
                        break;
                    }
                    case ("m"): {
                        momExpenses.add(report);
                        break;
                    }
                    default: {
                        System.out.println("Unexpected error. Inserted line is not an expected type.");
                    }
                }
            }
        });
        System.out.println("teste");
    }

    public List<BankReport> fillBankReports(HSSFCell cell, HSSFSheet sheet) {
        List<BankReport> bankReportList = new ArrayList<>();
        HSSFRow referenceRow = cell.getRow();
        HSSFRow currentRow = sheet.getRow(referenceRow.getRowNum() + 2);
        String[] reportLine = new String[7];
        while (!currentRow.getCell(0).getStringCellValue().trim().equals(STOP_REFERENCE)) {
            Iterator cellIterator = currentRow.cellIterator();
            while (cellIterator.hasNext()) {
                cell = (HSSFCell) cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    reportLine[cell.getColumnIndex()] = String.valueOf(cell.getNumericCellValue()).trim();
                } else {
                    reportLine[cell.getColumnIndex()] = cell.getStringCellValue().trim();
                }

            }
            BankReport bankReport = new BankReport(reportLine);
            bankReportList.add(bankReport);
            Integer currentRowNumber = currentRow.getRowNum();
            currentRow = sheet.getRow(currentRowNumber + 1);
        }
        return bankReportList;
    }

    public HSSFCell findReferenceCell(HSSFSheet sheet) {
        Iterator rowIterator = sheet.rowIterator();
        HSSFRow row = null;
        HSSFCell cell = null;
        dateFieldFindingLoop:
        while (rowIterator.hasNext()) {
            row = (HSSFRow) rowIterator.next();
            Iterator cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                cell = (HSSFCell) cellIterator.next();
                if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().trim().equals(SEARCH_REFERENCE))
                    break dateFieldFindingLoop;
            }
        }
        return cell;
    }
}

