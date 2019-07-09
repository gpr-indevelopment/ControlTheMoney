package com.MoneyControl;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;

@Component
public class CompiledData {

    @Autowired
    private SheetBuilder sheetBuilder;

    private final String COMPILED_DATA_SPREADSHEET = "Compiled data";
    private final String FIXED_COSTS = "Fixed costs:";
    private final String VARIABLE_COSTS_BUDGET = "Variable costs budget:";
    private final String MOM_EXPENSES = "Mom expenses:";
    private final String MOM_CREDIT = "Mom credit:";
    private final String MOM_IMPACT = "Mom impact:";
    private final String VARIABLE_COSTS = "Variable costs:";
    private final String MONEY_RECEIVED = "Salary:";

    private final String[] indicatorsArray = {MONEY_RECEIVED, FIXED_COSTS, VARIABLE_COSTS_BUDGET, MOM_EXPENSES, MOM_CREDIT, MOM_IMPACT, VARIABLE_COSTS};

    private final String[] HEADERS = {"INDICATOR", "VALUE"};

    @Getter @Setter
    private String moneyReceived;

    @Getter @Setter
    private String fixedCosts;

    @Getter @Setter
    private String variableCostsBudget;

    @Getter @Setter
    private String momExpenses;

    @Getter @Setter
    private String momCredit;

    @Getter @Setter
    private String momImpact;

    @Getter @Setter
    private String variableCosts;

    @Getter @Setter
    private String outputSheetPath;

    private Map<String, String> indicatorMap = new HashMap<>();

    public void buildCompiledDataSpreadsheet(Workbook workbook){
        Sheet sheet = workbook.createSheet(COMPILED_DATA_SPREADSHEET);
        sheetBuilder.buildAndStyleHeaders(sheet, HEADERS);
        compileVariableCosts(workbook);
        compileFixedCosts(workbook);
        compileMomCredit(workbook);
        compileMomExpenses(workbook);
        calculateVariableCostsBudget();
        calculateMomImpact();
        buildCompiledDataSheet(sheet);
        sheetBuilder.applyStylesToSheet(sheet);
        System.out.println("teste");
    }

    public void buildCompiledDataSheet(Sheet sheet){
        Integer currentRowNum = 1;
        for(int i=0;i<indicatorMap.size();i++){
            Row currentRow = sheet.createRow(currentRowNum);
            Cell indicatorCell = currentRow.createCell(0);
            indicatorCell.setCellValue(indicatorsArray[i]);
            Cell valueCell = currentRow.createCell(1);
            valueCell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
            valueCell.setCellFormula(indicatorMap.get(indicatorsArray[i]));
            currentRowNum++;
        }
    }

    public void calculateVariableCostsBudget(){
        this.variableCostsBudget = this.moneyReceived + "+" + this.fixedCosts;
        indicatorMap.put(MONEY_RECEIVED, this.moneyReceived);
        indicatorMap.put(VARIABLE_COSTS_BUDGET, variableCostsBudget);
    }

    public void calculateMomImpact(){
        this.momImpact = this.momExpenses + "+" + this.momCredit;
        indicatorMap.put(MOM_IMPACT, momImpact);
    }

    private void compileVariableCosts(Workbook workbook){
        Sheet sheet = workbook.getSheet(SheetBuilder.VARIABLE_EXPENSES_KEY);
        Integer firstRowNum = 1;
        Integer lastRowNum = sheet.getLastRowNum() + 1;
        Integer lastColumnNum = SheetBuilder.HEADERS.length-1;
        String lastColumnLetter = CellReference.convertNumToColString(lastColumnNum);
        String sheetReference = Utils.convertSheetNameToFormulaReference(SheetBuilder.VARIABLE_EXPENSES_KEY);
        String variableCostsformula = "SUM(" + sheetReference + lastColumnLetter + firstRowNum + ":" + sheetReference + lastColumnLetter + lastRowNum + ")";
        setVariableCosts(variableCostsformula);
        indicatorMap.put(VARIABLE_COSTS, variableCosts);
    }

    private void compileFixedCosts(Workbook workbook){
        Sheet sheet = workbook.getSheet(SheetBuilder.FIXED_EXPENSES_KEY);
        Integer firstRowNum = 1;
        Integer lastRowNum = sheet.getLastRowNum() + 1;
        Integer lastColumnNum = SheetBuilder.HEADERS.length-1;
        String lastColumnLetter = CellReference.convertNumToColString(lastColumnNum);
        String sheetReference = Utils.convertSheetNameToFormulaReference(SheetBuilder.FIXED_EXPENSES_KEY);
        String fixedCostsformula = "SUM(" + sheetReference + lastColumnLetter + firstRowNum + ":" + sheetReference + lastColumnLetter + lastRowNum + ")";
        setFixedCosts(fixedCostsformula);
        indicatorMap.put(FIXED_COSTS, fixedCosts);
    }

    private void compileMomExpenses(Workbook workbook){
        Sheet sheet = workbook.getSheet(SheetBuilder.MOM_EXPENSES_KEY);
        Integer firstRowNum = 1;
        Integer lastRowNum = sheet.getLastRowNum() + 1;
        Integer lastColumnNum = SheetBuilder.HEADERS.length-1;
        String lastColumnLetter = CellReference.convertNumToColString(lastColumnNum);
        String sheetReference = Utils.convertSheetNameToFormulaReference(SheetBuilder.MOM_EXPENSES_KEY);
        String momExpensesformula = "SUM(" + sheetReference + lastColumnLetter + firstRowNum + ":" + sheetReference + lastColumnLetter + lastRowNum + ")";
        setMomExpenses(momExpensesformula);
        indicatorMap.put(MOM_EXPENSES, momExpenses);
    }

    private void compileMomCredit(Workbook workbook){
        Sheet sheet = workbook.getSheet(SheetBuilder.MOM_CREDIT_KEY);
        Integer firstRowNum = 1;
        Integer lastRowNum = sheet.getLastRowNum() + 1;
        Integer lastColumnNum = SheetBuilder.HEADERS.length-1;
        String lastColumnLetter = CellReference.convertNumToColString(lastColumnNum);
        String sheetReference = Utils.convertSheetNameToFormulaReference(SheetBuilder.MOM_CREDIT_KEY);
        String momCreditformula = "SUM(" + sheetReference + lastColumnLetter + firstRowNum + ":" + sheetReference + lastColumnLetter + lastRowNum + ")";
        setMomCredit(momCreditformula);
        indicatorMap.put(MOM_CREDIT, momCredit);
    }
}
