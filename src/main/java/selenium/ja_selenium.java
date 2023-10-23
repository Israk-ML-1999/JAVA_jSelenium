package selenium;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class GoogleSearchSuggestions {
    public static void main(String[] args) {
        // Get the current day of the week (e.g., "Thursday")
        String currentDayOfWeek = getCurrentDayOfWeek();
        
        // Initialize WebDriver (Firefox)
        WebDriver webDriver = initializeFirefoxDriver();
        
        // Path to the Excel file
        String excelFilePath = "/Users/asus/Assignment_java/seleniumproject/Excel.xlsx";

        try (FileInputStream excelFileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook excelWorkbook = new XSSFWorkbook(excelFileInputStream)) {

            // Check if the current day's sheet exists in the workbook
            int currentDaySheetIndex = excelWorkbook.getSheetIndex(currentDayOfWeek);

            if (currentDaySheetIndex != -1) {
                Sheet currentDaySheet = excelWorkbook.getSheet(currentDayOfWeek);
                processWorksheetRows(webDriver, currentDaySheet);
                saveExcelFile(excelFilePath, excelWorkbook);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the browser
            closeWebDriver(webDriver);
        }
    }

    private static String getCurrentDayOfWeek() {
        // Get the current day of the week as a string (e.g., "Thursday")
        SimpleDateFormat dayOfWeekFormat = new SimpleDateFormat("EEEE");
        return dayOfWeekFormat.format(new Date());
    }

    private static WebDriver initializeFirefoxDriver() {
        // Initialize and return a Firefox WebDriver
        return new FirefoxDriver();
    }

    private static void processWorksheetRows(WebDriver webDriver, Sheet worksheet) {
        for (int rowIndex = 1; rowIndex < worksheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = worksheet.getRow(rowIndex);
            Cell keywordCell = row.getCell(1);

            if (keywordCell != null) {
                String keyword = keywordCell.getStringCellValue();

                if (keyword != null && !keyword.isEmpty()) {
                    searchAndProcessKeyword(webDriver, row, keyword);
                }
            }
        }
    }

    private static void searchAndProcessKeyword(WebDriver webDriver, Row row, String keyword) {
        // ... Existing code to search and process keyword ...

        // Update the corresponding columns in the Excel file
        row.createCell(2).setCellValue(minSuggestion);
        row.createCell(3).setCellValue(maxSuggestion);
    }

    private static void saveExcelFile(String excelFilePath, Workbook excelWorkbook) throws IOException {
        // Save the modified Excel workbook back to the file
        try (FileOutputStream excelFileOutputStream = new FileOutputStream(excelFilePath)) {
            excelWorkbook.write(excelFileOutputStream);
        }
    }

    private static void closeWebDriver(WebDriver webDriver) {
        // Close and quit the WebDriver
        webDriver.quit();
    }
}
