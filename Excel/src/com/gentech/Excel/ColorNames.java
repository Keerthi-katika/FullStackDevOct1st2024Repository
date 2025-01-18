package com.gentech.Excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ColorNames {
    public static void main(String[] args) {
        // Array of 20 color names
        String[] colors = {
                "Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown",
                "Black", "White", "Gray", "Cyan", "Magenta", "Gold", "Silver", "Violet",
                "Indigo", "Maroon", "Teal", "Lavender"
        };

        // Create a workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("FirstSheet");

        // Write color names in the 5th column (index 4 because it's 0-based)
        for (int i = 0; i < colors.length; i++) {
            Row row = sheet.createRow(i);     // Create rows starting from index 0
            Cell cell = row.createCell(4);   // Create cell in the 5th column (column index 4)
            cell.setCellValue(colors[i]);    // Set the cell value to the color name
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("C:\\Demo\\Test\\jkcolors.xlsx"))
        {
            workbook.write(fileOut);
            System.out.println("Excel file with color names written successfully!");
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

