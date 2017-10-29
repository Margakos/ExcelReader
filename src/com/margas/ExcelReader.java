package com.margas;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalTime;
import java.util.Arrays;
import java.util.Date;
import java.util.List;


public class ExcelReader {

    public static void main(String args[]) throws Exception {

        File src = new File("C:\\Users\\margas\\Desktop\\sampling.xlsx");
        FileInputStream fis = new FileInputStream(src);
        FileOutputStream fout = new FileOutputStream(src);

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet1 = wb.getSheetAt(0);
        XSSFWorkbook wbtw= new XSSFWorkbook(OPCPackage.create(fout));
        XSSFSheet sheet1towrite = wbtw.getSheetAt(0);

        for (int i = 0; i < sheet1.getLastRowNum(); i++) {

            String cellContent = sheet1.getRow(i).getCell(16).getStringCellValue();
            int semiColonPosition = cellContent.indexOf(":");
            sheet1towrite.getRow(i).getCell(22).setCellValue(cellContent);
            //System.out.println("---> " +semiColonPosition);


            if (semiColonPosition < 0) {
                System.out.println(cellContent);
                //sheet1.getRow(i).getCell(22).setCellValue("eimai xazos");
//                Cell cell = null;
//                cell.setCellValue(cellContent);
////                sheet1towrite.getRow(i).getCell(22).setCellValue(cellContent);
//                fis.close();
//                FileOutputStream asdf =new FileOutputStream(new File("C:\\Users\\margas\\Desktop\\sampling.xlsx"));
                continue;
            } else if (semiColonPosition < 5) {
                //System.out.println(cellContent);
                String comingHour = cellContent.substring(semiColonPosition - 2, semiColonPosition);
                String comingMinutes = cellContent.substring(semiColonPosition + 1, semiColonPosition + 3);
                //System.out.println("semiColonPosition=" + semiColonPosition);
                // System.out.println("comingHour=" + comingHour);
                // System.out.println("comingMinutes=" + comingMinutes);

                int secondsemiColonPosition = cellContent.lastIndexOf(":");
                String leavingHour = cellContent.substring(secondsemiColonPosition - 2, secondsemiColonPosition);
                String leavingMinutes = cellContent.substring(secondsemiColonPosition + 1, secondsemiColonPosition + 3);

                //System.out.println("leavingHour=" + leavingHour);
                //System.out.println("leavingMinutes=" + leavingMinutes);


                LocalTime comingTime = LocalTime.parse(comingHour + ":" + comingMinutes + ":00");
                LocalTime leavingTime = LocalTime.parse(leavingHour + ":" + leavingMinutes + ":00");

                long elapsedMinutes = Duration.between(comingTime, leavingTime).toMinutes();

                if ((elapsedMinutes) > 480) {
                    System.out.println("20min διότι δουλεψε πανω απο 8ωρο ");



//
//                    //Testing to write on specific column
//                    Row r = sheet1.getRow(9); // 10-1
//                    if (r == null) {
//                        // First cell in the row, create
//                        r = sheet1.createRow(9);
//                    }
//
//                    Cell c = r.getCell(3); // 4-1
//                    if (c == null) {
//                        // New cell
//                        c = r.createCell(3, c.setCellType(XSSFCell.CELL_TYPE_NUMERIC));
//                    }
//                    c.setCellValue(100);

                    continue;
                } else if ((elapsedMinutes) < 540)
                    System.out.println("Ο υπάλληλος δούλεψε 8ωρο ");
                continue;


            } else if (semiColonPosition > 10) {
                System.out.println(cellContent);

            } else
                continue;

            String comingHour = cellContent.substring(semiColonPosition - 2, semiColonPosition);
            String comingMinutes = cellContent.substring(semiColonPosition + 1, semiColonPosition + 3);
            //System.out.println("semiColonPosition=" + semiColonPosition);
            // System.out.println("comingHour=" + comingHour);
            // System.out.println("comingMinutes=" + comingMinutes);

            int secondsemiColonPosition = cellContent.lastIndexOf(":");
            String leavingHour = cellContent.substring(secondsemiColonPosition - 2, secondsemiColonPosition);
            String leavingMinutes = cellContent.substring(secondsemiColonPosition + 1, secondsemiColonPosition + 3);

            //System.out.println("leavingHour=" + leavingHour);
            //System.out.println("leavingMinutes=" + leavingMinutes);


            LocalTime comingTime = LocalTime.parse(comingHour + ":" + comingMinutes + ":00");
            LocalTime leavingTime = LocalTime.parse(leavingHour + ":" + leavingMinutes + ":00");

            long elapsedMinutes = Duration.between(comingTime, leavingTime).toMinutes();


            System.out.println("Ο υπάλληλος δούλεψε  " + elapsedMinutes / 60 + "ώρες πανω σε ->" + (cellContent));


        }

        wb.close();
    }


}








