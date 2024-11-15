package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Scanner;

public class KeywordTracker {
    public static void main(String[] args) {
        // Variables to track the longest and shortest keywords
        String longestKeyword = "";
        String shortestKeyword = "";

        // Scanner for user input
        Scanner scanner = new Scanner(System.in);

        System.out.println("Enter keywords for Google search (type 'exit' to finish):");

        while (true) {
            System.out.print("Enter keyword: ");
            String input = scanner.nextLine();

            if (input.equalsIgnoreCase("exit")) {
                break;
            }

            // Update longest and shortest keywords
            if (input.length() > longestKeyword.length()) {
                longestKeyword = input;
            }
            if (shortestKeyword.isEmpty() || input.length() < shortestKeyword.length()) {
                shortestKeyword = input;
            }

            System.out.println("Current Longest: " + longestKeyword);
            System.out.println("Current Shortest: " + shortestKeyword);
        }

        // Write results to Excel
        writeKeywordsToExcel(longestKeyword, shortestKeyword);
        System.out.println("Longest and shortest keywords saved to Excel file.");
    }

    private static void writeKeywordsToExcel(String longestKeyword, String shortestKeyword) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Keywords");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Date");
        headerRow.createCell(1).setCellValue("Longest Keyword");
        headerRow.createCell(2).setCellValue("Shortest Keyword");

        // Create data row
        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue(LocalDate.now().toString());
        dataRow.createCell(1).setCellValue(longestKeyword);
        dataRow.createCell(2).setCellValue(shortestKeyword);

        // Write to Excel file
        try (FileOutputStream fileOut = new FileOutputStream("GoogleKeywords.xlsx")) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
