package com.MoneyControl;

import com.MoneyControl.SheetOutput.FixedExpenses;
import com.MoneyControl.SheetOutput.MomCredit;
import com.MoneyControl.SheetOutput.MomExpenses;
import com.MoneyControl.SheetOutput.VariableExpenses;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.util.*;

@Service
public class MainService {

    @Autowired
    private FixedExpenses fixedExpenses;

    @Autowired
    private MomCredit momCredit;

    @Autowired
    private MomExpenses momExpenses;

    @Autowired VariableExpenses variableExpenses;

    private final String SEARCH_REFERENCE = "Data";
    private final String STOP_REFERENCE = "TOTAL";
    private final String SHEET_NAME = "/sample.xls";
    private final Integer SEARCH_OFFSET_FROM_REFERENCE = 2;
    private final Integer REFERENCE_SALARY = 2000;

    public void init() {

        HSSFSheet sheet = Utils.getSheet(SHEET_NAME);
        Integer referenceRow = findReferenceRow(sheet);
        List<BankReport> bankReportList = fillBankReports(referenceRow, sheet);
        Map<String, List> expensesMap = routeExpenses(bankReportList);
        Workbook workbook = buildWorkbook(expensesMap);
    }

    public Workbook buildWorkbook(Map<String, List> expensesMap){
        Workbook workbook = new HSSFWorkbook();
        fixedExpenses.buildFixedExpensesSheet(expensesMap, workbook);
        momCredit.buildMomCredit(expensesMap, workbook);
        momExpenses.buildMomExpensesSheet(expensesMap, workbook);
        variableExpenses.buildVariableExpensesSheet(expensesMap, workbook);
        try{
            FileOutputStream fileOut = new FileOutputStream("finance_workbook.xls");
            workbook.write(fileOut);
            fileOut.close();
        }
        catch (Exception e){

        }
        return workbook;
    }

    public Map<String, List> routeExpenses(List<BankReport> bankReportsList) {
        Scanner scanner = new Scanner(System.in);
        Map<String, List> expensesMap = new HashMap<>();
        List<BankReport> fixedExpenses = new ArrayList<>();
        List<BankReport> variableExpenses = new ArrayList<>();
        List<BankReport> momExpenses = new ArrayList<>();
        List<BankReport> momCredit = new ArrayList<>();
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
            else{
                if(report.getCredit() > 0 && report.getCredit() < REFERENCE_SALARY) momCredit.add(report);
            }
        });
        expensesMap.put(FixedExpenses.FIXED_EXPENSES_KEY, fixedExpenses);
        expensesMap.put(VariableExpenses.VARIABLE_EXPENSES_KEY, variableExpenses);
        expensesMap.put(MomExpenses.MOM_EXPENSES_KEY, momExpenses);
        expensesMap.put(MomCredit.MOM_CREDIT_KEY, momCredit);
        return expensesMap;
    }

    public List<BankReport> fillBankReports(Integer referenceRow, HSSFSheet sheet) {
        List<BankReport> bankReportList = new ArrayList<>();
        HSSFRow currentRow = sheet.getRow(referenceRow);
        String[] reportLine = new String[7];
        while (!currentRow.getCell(0).getStringCellValue().trim().equals(STOP_REFERENCE)) {
            Iterator cellIterator = currentRow.cellIterator();
            while (cellIterator.hasNext()) {
                HSSFCell cell = (HSSFCell) cellIterator.next();
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

    public Integer findReferenceRow(HSSFSheet sheet) {
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
        return cell.getRow().getRowNum() + SEARCH_OFFSET_FROM_REFERENCE;
    }
}

