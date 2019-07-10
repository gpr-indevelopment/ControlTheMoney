package com.MoneyControl.Utils;

import com.MoneyControl.Model.BankReport;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Utils {

    public static final String ANSI_RESET = "\u001B[0m";
    public static final String ANSI_BLUE = "\u001B[34m";
    public static DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");


    public static String formatBankReportDisplay(BankReport bankReport){
        String date = bankReport.getDate().toString();
        String description = bankReport.getDescription();
        String debit = String.valueOf(bankReport.getDebit());
        return ANSI_BLUE + "DATE: " + ANSI_RESET + date + " - " + ANSI_BLUE + "DESCRIPTION: " + ANSI_RESET + description + " - " + ANSI_BLUE + "DEBIT: " + ANSI_RESET + debit;
    }

    public static String trimDescription(String description){
        String[] wordsToRemove = {"COMPRA", "TARIFA", "PAGAMENTO", "PGTO", "OUTRO"};
        for(int i=0;i<wordsToRemove.length;i++){
            String word = wordsToRemove[i];
            if(description.contains(word)){
                description.replace(word, "");
            }
        }
        while(description.contains("  ")){
            description = description.replace("  ", " ");
        }
        return description;
    }

    public static List<String> rowContentToStringList(HSSFRow row){
        List<String> rowContent = new ArrayList<>();
        Iterator cellIterator = row.cellIterator();
        while(cellIterator.hasNext()){
            HSSFCell currentCell = (HSSFCell) cellIterator.next();
            if(currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC){
                rowContent.add(String.valueOf(currentCell.getNumericCellValue()));
            }
            else if(currentCell.getCellType() == Cell.CELL_TYPE_STRING){
                rowContent.add(currentCell.getStringCellValue());
            }
        }
        rowContent.remove(rowContent.size()-1);
        return rowContent;
    }

    public static List<Double> rowContentToDoubleList(HSSFRow row){
        List<Double> rowContent = new ArrayList<>();
        Iterator cellIterator = row.cellIterator();
        while(cellIterator.hasNext()){
            HSSFCell currentCell = (HSSFCell) cellIterator.next();
            if(currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC){
                rowContent.add(currentCell.getNumericCellValue());
            }
        }
        if(rowContent.size()>0){
            rowContent.remove(rowContent.size()-1);
        }
        return rowContent;
    }

    public static String convertDateStringToMonth(String dateString){
        LocalDate date = LocalDate.parse(dateString.trim(), formatter);
        String monthString = String.valueOf(date.getMonthValue());
        if(monthString.length() == 1) monthString = "0" + monthString;
        return monthString;
    }
}
