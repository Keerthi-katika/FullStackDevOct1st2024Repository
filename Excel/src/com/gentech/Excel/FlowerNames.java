package com.gentech.Excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class FlowerNames {
    public static void main(String[] args) {
        // Create a workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Write "flower1" to "flower20" in the first column
        for (int i = 1; i <= 20; i++) {
            Row row = sheet.createRow(i - 1); // Create row (0-based index)
            Cell cell = row.createCell(0);   // Create cell in the first column
            cell.setCellValue("flower" + i); // Set cell value
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("C:\\Demo\\Test\\flowers.xlsx"))
        {
            workbook.write(fileOut);
            System.out.println("Excel file written successfully!");
        } catch (IOException e) {
            System.out.println("An error occurred while writing the file: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing the workbook: " + e.getMessage());
            }
        }
    }
}

