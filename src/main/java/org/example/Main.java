package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) throws IOException {

        System.out.println("Hello world!");
        String excelFilePath = ".\\src\\main\\resources\\testRecords.xlsx";
        String sheetName = "loginInfo";
       // File testFile = new File(excelFilePath);
        FileInputStream testStreamFile = new FileInputStream(excelFilePath);
        XSSFWorkbook workB = new XSSFWorkbook(testStreamFile) ;
        XSSFSheet sheetTest =  workB.getSheet(sheetName);
        System.out.println(sheetTest.getHeader());
        int lastrowNum = sheetTest.getLastRowNum();
        System.out.println(lastrowNum);
        int firstrowNum = sheetTest.getFirstRowNum();
        System.out.println("hello11"+firstrowNum);
        int recordsNum = lastrowNum-firstrowNum;

        for (int i = 1;i<=recordsNum;i++){
            /*
            XSSFRow row = sheetTest.getRow(i);
            int colNum = row.getLastCellNum();
             */

            Iterator<Cell> rowCellValues = sheetTest.getRow(i).iterator();
            while(rowCellValues.hasNext()){
                System.out.println(rowCellValues.next());
            }
        }


    }
}