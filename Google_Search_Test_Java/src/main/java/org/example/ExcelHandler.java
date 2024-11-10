package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelHandler {

    // Method to read keywords for the current day from the Excel sheet
    public static List<String> getKeywordsForDay(String currentDayName) {
        List<String> keywords = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream("src/main/resources/files/reformatted_sample.xlsx");
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Get the sheet corresponding to the current day
            Sheet sheet = workbook.getSheet(currentDayName);
            if (sheet == null) {
                System.out.println("Sheet for " + currentDayName + " not found.");
                return keywords;
            }

            // Iterate over all rows and collect keywords from the second column (column index 1, "Keyword")
            boolean isHeaderRow = true;
            for (Row row : sheet) {
                // Skip the first row (header row)
                if (isHeaderRow) {
                    isHeaderRow = false;
                    continue;
                }

                // Get the keyword from the second column (index 1)
                Cell cell = row.getCell(1);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    keywords.add(cell.getStringCellValue());
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return keywords;
    }

    // Method to write the results into the same Excel file (reformatted_sample.xlsx)
    public static void writeResultsToExcel(String currentDayName, String keyword, String longestSuggestion, String smallestSuggestion) {
        File file = new File("src/main/resources/files/reformatted_sample.xlsx");
        Workbook workbook;

        try {
            // Open the existing workbook (no need to create a new file)
            try (FileInputStream fis = new FileInputStream(file)) {
                workbook = new XSSFWorkbook(fis); // Open the existing workbook
            }

            // Get the sheet corresponding to the current day
            Sheet sheet = workbook.getSheet(currentDayName);
            if (sheet == null) {
                sheet = workbook.createSheet(currentDayName); // Create a new sheet with the current day's name if not found
            }

            // Iterate through the rows to find the row with the matching keyword
            boolean keywordFound = false;
            for (Row row : sheet) {
                Cell keywordCell = row.getCell(1); // Keywords are in the second column (index 1)
                if (keywordCell != null && keywordCell.getCellType() == CellType.STRING) {
                    // If the keyword matches, write the suggestions in the corresponding row
                    if (keywordCell.getStringCellValue().equals(keyword)) {
                        // Write the longest and smallest options in columns 2 and 3 respectively
                        row.createCell(2).setCellValue(longestSuggestion); // Column 2 for longest suggestion
                        row.createCell(3).setCellValue(smallestSuggestion); // Column 3 for smallest suggestion
                        keywordFound = true;
                        break;
                    }
                }
            }

            // If the keyword was not found, you can optionally add a new row (or handle this case as needed)
            if (!keywordFound) {
                System.out.println("Keyword not found in the sheet: " + keyword);
            } else {
                // Save the workbook to the same file (no separate output file)
                try (FileOutputStream fileOut = new FileOutputStream(file)) {
                    workbook.write(fileOut);
                }

                // Show a success message
                System.out.println("Data saved successfully!");
            }

            // Close the workbook after writing
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
