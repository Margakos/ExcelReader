package com.margas;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class ExcelAppender {

    FileInputStream fip = new FileInputStream(new File("C:\\Users\\margas\\Desktop\\sampling.xlsx"));
    XSSFWorkbook wb = new XSSFWorkbook(fip);
    XSSFSheet sheet = wb.getSheetAt(0);
    Cell cell = null;
    cell = worksheet.getRow(1).getCell(1);
    cell.setCellValue

}
