package com.gentech.Excel;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CityNames {
    public static void main(String[] args) {
        // Array of 20 city names
        String[] cities = {
                "New York", "Los Angeles", "Chicago", "Houston", "Phoenix", "Philadelphia",
                "San Antonio", "San Diego", "Dallas", "San Jose", "Austin", "Jacksonville",
                "Fort Worth", "Columbus", "San Francisco", "Charlotte", "Indianapolis",
                "Seattle", "Denver", "Washington"
        };

        // Create a workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("FirstSheet");

        // Create the 10th row (row index is 9 because it's 0-based)
        Row row = sheet.createRow(9); // Row index 9 is the 10th row

        // Write city names into the 10th row
        for (int i = 0; i < cities.length; i++) {
            Cell cell = row.createCell(i); // Create cells in the 10th row
            cell.setCellValue(cities[i]);  // Set cell value to the city name
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("K:\\Demo\\Test\\cities.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Excel file with city names written successfully!");
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

