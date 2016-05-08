package com.daya.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

/**
 * Created by keshav on 8/5/16.
 */
public class SimpleExcelReaderExample {
  public static void main(String[] args) {
    FileInputStream inputStream = null;
    try {
      inputStream = new FileInputStream(new File("/home/keshav/Downloads/sample.xlsx"));
      Workbook workbook = new XSSFWorkbook(inputStream);
      Sheet firstSheet = workbook.getSheetAt(0);
      Iterator<Row> iterator = firstSheet.iterator();
      while (iterator.hasNext()) {
        Row nextRow = iterator.next();
        Iterator<Cell> cellIterator = nextRow.cellIterator();
        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
              System.out.print(cell.getStringCellValue());
              break;
            case Cell.CELL_TYPE_BOOLEAN:
              System.out.print(cell.getBooleanCellValue());
              break;
            case Cell.CELL_TYPE_NUMERIC:
              System.out.print(cell.getNumericCellValue());
              break;
          }
          System.out.print(" - ");
        }
        System.out.println();
      }
      workbook.close();
      inputStream.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
