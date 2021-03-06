package com.MoneyControl;

import com.MoneyControl.Model.BankReport;
import com.MoneyControl.SheetBuilding.CompiledData;
import com.MoneyControl.SheetBuilding.SheetBuilder;
import com.MoneyControl.Utils.SheetUtils;
import com.MoneyControl.Utils.Utils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.util.*;

@Service
public class MainService {

    // TODO: Write tests, setup code to run for more than one month.
    @Autowired
    private SheetBuilder sheetBuilder;

    @Autowired
    private CompiledData compiledData;

    private final String STOP_REFERENCE = "TOTAL";
    private final String OUTPUT_SHEET_EXTENSION = ".xls";
    private final Integer REFERENCE_SALARY = 2000;
    private final String ORIGIN_SHEET_NAME = "/report.xls";
    private final String OUTPUT_SHEET_NAME = "Finance_Workbook_";
    private final Integer SEARCH_OFFSET_FROM_REFERENCE = 1;
    private String outputSheetMonth;

    public void init() {
        HSSFSheet sheet = null;
        while(sheet == null){
            sheet = SheetUtils.getSheetFromLocalFile(ORIGIN_SHEET_NAME);
        }
        Integer referenceRow = findReferenceRow(sheet);
        List<BankReport> bankReportList = fillBankReports(referenceRow, sheet);
        Map<String, List> expensesMap = routeExpenses(bankReportList);
        buildAndSaveWorkbook(expensesMap);

    }

    public Integer findReferenceRow(HSSFSheet sheet) {
        Iterator rowIterator = sheet.rowIterator();
        HSSFRow row;
        HSSFCell cell = null;
        dateFieldFindingLoop:
        while (rowIterator.hasNext()) {
            row = (HSSFRow) rowIterator.next();
            boolean isSalaryRow = isSalaryRow(row);
            Iterator cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                cell = (HSSFCell) cellIterator.next();
                if (isSalaryRow) {
                    sheetBuilder.setSalary(String.valueOf(row.getCell(4).getNumericCellValue()));
                    outputSheetMonth = Utils.convertDateStringToMonth(row.getCell(0).getStringCellValue());
                    break dateFieldFindingLoop;
                }
            }
        }
        return cell.getRow().getRowNum() + SEARCH_OFFSET_FROM_REFERENCE;
    }

    public List<BankReport> fillBankReports(Integer referenceRow, HSSFSheet sheet) {
        List<BankReport> bankReportList = new ArrayList<>();
        HSSFRow currentRow = sheet.getRow(referenceRow);
        String[] reportLine = new String[7];
        while (!isStopReferenceRow(currentRow) && !isSalaryRow(currentRow)) {
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
                System.out.println("Format: [branchedExpenses];[originalBranch]. branchedExpenses must be a positive number.");
                String[] lineArray;
                String expensesRouter;
                Double valueToSubtract = 0d;
                Boolean success = false;
                while(!success){
                    success = true;
                    lineArray = scanner.nextLine().split(";");
                    if(lineArray.length == 2 && lineArray[0].chars().allMatch(Character::isDigit)){
                        valueToSubtract = Double.parseDouble(lineArray[0]);
                    }
                    BankReport subtractedReport = Utils.subtractFromDebit(report, valueToSubtract);
                    expensesRouter = lineArray[lineArray.length-1];
                    switch (expensesRouter) {
                        case "f": {
                            fixedExpenses.add(report);
                            if(subtractedReport.getDebit() < 0) momExpenses.add(subtractedReport);
                            break;
                        }
                        case "v": {
                            variableExpenses.add(report);
                            if(subtractedReport.getDebit() < 0) momExpenses.add(subtractedReport);
                            break;
                        }
                        case "m": {
                            momExpenses.add(report);
                            if(subtractedReport.getDebit() < 0) variableExpenses.add(subtractedReport);
                            break;
                        }
                        default: {
                            System.out.println("Unexpected error. Inserted line is not an expected type.");
                            System.out.println("Insert: f = fixed; v = variable; m = mom.");
                            success = false;
                        }
                    }
                }
            } else {
                if (report.getCredit() > 0 && report.getCredit() < REFERENCE_SALARY) momCredit.add(report);
            }
        });
        expensesMap.put(SheetBuilder.FIXED_EXPENSES_KEY, fixedExpenses);
        expensesMap.put(SheetBuilder.VARIABLE_EXPENSES_KEY, variableExpenses);
        expensesMap.put(SheetBuilder.MOM_EXPENSES_KEY, momExpenses);
        expensesMap.put(SheetBuilder.MOM_CREDIT_KEY, momCredit);
        return expensesMap;
    }

    public Workbook buildAndSaveWorkbook(Map<String, List> expensesMap) {
        String outputFilePath = "./" + OUTPUT_SHEET_NAME + outputSheetMonth + OUTPUT_SHEET_EXTENSION;
        Workbook workbook = sheetBuilder.buildExpensesSheet(expensesMap);
        compiledData.buildCompiledDataSpreadsheet(workbook);
        try {
            FileOutputStream fileOut = new FileOutputStream(outputFilePath);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            System.out.printf("An error occurred while building the output spreadsheet. Maybe a file with the same name is already open.");
        }
        return workbook;
    }

    public boolean isSalaryRow(HSSFRow row) {
        boolean isSalaryRow = false;
        List<Double> rowContent = Utils.rowContentToDoubleList(row);
        for (int i = 0; i < rowContent.size(); i++) {
            Double cellValue = rowContent.get(i);
            if (cellValue > REFERENCE_SALARY) {
                isSalaryRow = true;
            }
        }
        return isSalaryRow;
    }

    public boolean isStopReferenceRow(HSSFRow row) {
        boolean isStopReferenceRow = false;
        List<String> rowContent = Utils.rowContentToStringList(row);
        if (rowContent.contains(STOP_REFERENCE)) {
            isStopReferenceRow = true;
        }
        return isStopReferenceRow;
    }
}

