package com.MoneyControl;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Service;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

@Service
public class MainService {

    private final String SEARCH_REFERENCE = "Data";

    public void init(){

        HSSFSheet sheet = Utils.getSheet();
        HSSFCell referenceCell = findReferenceCell(sheet);
    }

    public HSSFCell findReferenceCell(HSSFSheet sheet){
        Iterator rowIterator = sheet.rowIterator();
        HSSFRow row = null;
        HSSFCell cell = null;
        dateFieldFindingLoop:
        while(rowIterator.hasNext())
        {
            row = (HSSFRow) rowIterator.next();
            Iterator cellIterator = row.cellIterator();
            while(cellIterator.hasNext())
            {
                cell = (HSSFCell) cellIterator.next();
                if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().trim().equals(SEARCH_REFERENCE))
                    break dateFieldFindingLoop;
            }
        }
        return cell;
    }
}

