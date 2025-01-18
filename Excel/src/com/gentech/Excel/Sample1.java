package com.gentech.Excel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import  java.io.FileInputStream;

public class Sample1 {
    public static void main(String[] args) {
        readContent();
    }

    private static void readContent() {
        FileInputStream fin = null;
        Workbook Wb = null;
        Sheet sh = null;
        Row row = null;
        Cell cell = null;
        try {
            fin = new FileInputStream("C:\\Demo\\Test\\Book.xlsx");
            Wb = new XSSFWorkbook(fin);
            sh = Wb.getSheet("Sheet1");
            int rc = sh.getPhysicalNumberOfRows();
            for (int r = 0; r < rc; r++) {
                row = sh.getRow(r);
                int cc = row.getPhysicalNumberOfCells();
                for (int c = 0; c < cc; c++) {
                    cell = row.getCell(c);
                    String data = cell.getStringCellValue();
                    System.out.printf("%-12s", data);
                }
                System.out.printf("\n");
            }

        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                fin.close();
                Wb.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}


