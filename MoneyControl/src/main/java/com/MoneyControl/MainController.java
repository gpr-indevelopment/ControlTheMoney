package com.MoneyControl;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;

@RestController
public class MainController {

    @Autowired
    private MainService mainService;

    @RequestMapping(name = "/init")
    public void init() {
        mainService.init();
    }
}
