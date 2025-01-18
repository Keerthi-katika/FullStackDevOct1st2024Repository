package com.gentech.Excel;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class FruitNames {
    public static void main(String[] args) {
        // Array of 20 fruit names
        String[] fruits = {
                "Apple", "Banana", "Cherry", "Date", "Elderberry", "Fig", "Grape", "Honeydew",
                "Indian Fig", "Jackfruit", "Kiwi", "Lemon", "Mango", "Nectarine", "Orange",
                "Papaya", "Quince", "Raspberry", "Strawberry", "Tangerine"
        };

        // Create a workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create the first row
        Row row = sheet.createRow(0); // Row 0 is the first row

        // Write fruit names to the first row
        for (int i = 0; i < fruits.length; i++) {
            Cell cell = row.createCell(i);  // Create cells in the first row
            cell.setCellValue(fruits[i]);  // Set cell value to the fruit name
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("C:\\Demo\\Test\\fruits.xlsx"))
        {
            workbook.write(fileOut);
            System.out.println("Excel file with fruits written successfully!");
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

