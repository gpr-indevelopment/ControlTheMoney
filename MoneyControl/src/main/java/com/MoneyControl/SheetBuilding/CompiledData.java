package com.MoneyControl.SheetBuilding;

import com.MoneyControl.Utils.SheetUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.util.HSSFColor;
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
    private final String SALARY = "Salary:";

    private final String[] indicatorsArray = {SALARY, FIXED_COSTS, VARIABLE_COSTS_BUDGET, VARIABLE_COSTS, MOM_EXPENSES, MOM_CREDIT, MOM_IMPACT};

    private final String[] HEADERS = {"INDICATOR", "VALUE"};

    private final short BORDER = CellStyle.BORDER_THIN;
    private final short HEADER_ALIGN = CellStyle.ALIGN_CENTER;
    private final short HEADER_FONT_COLOR = HSSFColor.ROYAL_BLUE.index;

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
        this.indicatorMap = retreiveDataFromBaseSheet(workbook);
        Sheet sheet = workbook.createSheet(COMPILED_DATA_SPREADSHEET);
        SheetUtils.buildAndStyleHeaders(sheet, HEADERS, BORDER, HEADER_ALIGN, HEADER_FONT_COLOR);
        buildCompiledDataSheet(sheet);
        SheetUtils.applyStylesToSheet(sheet, BORDER);
    }

    private Map<String, String> retreiveDataFromBaseSheet(Workbook workbook){
        Map<String, String> indicatorMap = new HashMap<>();
        for(int i=0;i<workbook.getNumberOfSheets();i++){
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            switch (sheetName){
                case SheetBuilder.FIXED_EXPENSES_KEY:{
                    String key = FIXED_COSTS;
                    String value = SheetUtils.buildFormulaOfLastColumn(sheet, SheetUtils.EXCEL_SUM);
                    indicatorMap.put(key, value);
                    String key2 = VARIABLE_COSTS_BUDGET;
                    String value2 = sheetBuilder.getSalary() + "+" + value;
                    indicatorMap.put(key2, value2);
                    break;
                }
                case SheetBuilder.VARIABLE_EXPENSES_KEY:{
                    String key = VARIABLE_COSTS;
                    String value = SheetUtils.buildFormulaOfLastColumn(sheet, SheetUtils.EXCEL_SUM);
                    indicatorMap.put(key, value);
                    break;
                }
                case SheetBuilder.MOM_EXPENSES_KEY:{
                    String key = MOM_EXPENSES;
                    String value = SheetUtils.buildFormulaOfLastColumn(sheet, SheetUtils.EXCEL_SUM);
                    indicatorMap.put(key, value);
                    break;
                }
                case SheetBuilder.MOM_CREDIT_KEY:{
                    String key = MOM_CREDIT;
                    String value = SheetUtils.buildFormulaOfLastColumn(sheet, SheetUtils.EXCEL_SUM);
                    indicatorMap.put(key, value);
                    break;
                }
            }
        }
        String key = SALARY;
        String value = sheetBuilder.getSalary();
        indicatorMap.put(key, value);
        String key2 = MOM_IMPACT;
        String value2 = indicatorMap.get(MOM_EXPENSES) + "+" + indicatorMap.get(MOM_CREDIT);
        indicatorMap.put(key2, value2);
        return indicatorMap;
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
}
