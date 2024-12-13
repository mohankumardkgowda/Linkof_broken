package harishsample;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Harish1 {
    WebDriver driver;

    @BeforeTest
    public void beforeTest() {
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        driver.get("https://movedocs.com/");
        driver.manage().window().maximize();
    }

    @Test
    public void test() throws IOException {
        List<WebElement> links = driver.findElements(By.tagName("a"));

        // Lists to store broken and non-broken links
        List<String[]> nonBrokenLinks = new ArrayList<>();
        List<String[]> brokenLinks = new ArrayList<>();

        // Iterate through each link and classify based on status code
        for (WebElement link : links) {
            String url = link.getAttribute("href");
            if (url != null && !url.isEmpty()) {
                int statusCode = getStatusCode(url);

                if (statusCode >= 200 && statusCode < 400) {
                    nonBrokenLinks.add(new String[]{url, String.valueOf(statusCode)});
                } else {
                    brokenLinks.add(new String[]{url, String.valueOf(statusCode)});
                }
            }
        }

        // Write results to an Excel file
        writeToExcel(nonBrokenLinks, brokenLinks);

        // Print totals to the console
        System.out.println("Total Links: " + links.size());
        System.out.println("Non-Broken Links: " + nonBrokenLinks.size());
        System.out.println("Broken Links: " + brokenLinks.size());
    }

    public static int getStatusCode(String url) {
        try {
            @SuppressWarnings("deprecation")
			HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
            connection.setRequestMethod("GET");
            connection.connect();
            return connection.getResponseCode();
        } catch (Exception e) {
            System.err.println("Error fetching status for URL: " + url + " - " + e.getMessage());
            return -1; // Return -1 if an error occurs
        }
    }

    public static void writeToExcel(List<String[]> nonBrokenLinks, List<String[]> brokenLinks) throws IOException {
        @SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook();

        // Create sheets for non-broken and broken links
        Sheet nonBrokenSheet = workbook.createSheet("Non-Broken Links");
        Sheet brokenSheet = workbook.createSheet("Broken Links");
        // Write headers
        createHeader(nonBrokenSheet);
        createHeader(brokenSheet);
        // Write non-broken links
        writeLinksToSheet(nonBrokenSheet, nonBrokenLinks);

        // Write broken links
        writeLinksToSheet(brokenSheet, brokenLinks);

        // Save to file
        try (FileOutputStream fileOut = new FileOutputStream("LinksStatusClassification.xlsx")) {
            workbook.write(fileOut);
        }

        System.out.println("Excel file 'LinksStatusClassification.xlsx' created successfully!");
    }

    public static void createHeader(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("URL");
        headerRow.createCell(1).setCellValue("Status Code");
    }

    public static void writeLinksToSheet(Sheet sheet, List<String[]> links) {
        int rowNum = 1;
        for (String[] link : links) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(link[0]); // URL
            row.createCell(1).setCellValue(link[1]); // Status Code
        }
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
    }

    @AfterTest
    public void afterTest() {
        driver.quit();
    }
}
