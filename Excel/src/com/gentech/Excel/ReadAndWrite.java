package com.gentech.Excel;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ReadAndWrite {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Demo\\Test\\fruits.xlsx";  // Input Excel file
        String outputFilePath = "C:\\Demo\\Test\\output_fruits.xlsx"; // Output Excel file

        try (FileInputStream fileIn = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fileIn)) {

            // Read from the first sheet
            Sheet sheet1 = workbook.getSheetAt(0); // First sheet (0-based index)
            String[] fruits = new String[20];

            for (int i = 0; i < 20; i++) {
                Row row = sheet1.getRow(i); // Get the i-th row
                if (row != null) {
                    Cell cell = row.getCell(0); // Get the first column (index 0)
                    if (cell != null) {
                        fruits[i] = cell.getStringCellValue(); // Read the fruit name
                    }
                }
            }

            // Write to the second sheet
            Sheet sheet2 = workbook.createSheet("Sheet2");
            Row row5 = sheet2.createRow(4); // 5th row (index 4)

            for (int i = 0; i < fruits.length; i++) {
                Cell cell = row5.createCell(i); // Write in columns of the 5th row
                cell.setCellValue(fruits[i]); // Set the cell value to the fruit name
            }

            // Write the updated workbook to a new file
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
                System.out.println("Data written to the 5th row of Sheet2 successfully!");
            }

        } catch (IOException e) {
            System.out.println("An error occurred: " + e.getMessage());
        }
    }
}

