package abcd_datadriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class abcd_datadriven {
    WebDriver driver;

    @SuppressWarnings("deprecation")
	@BeforeTest
    public void setUp() {
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    }

    @Test
    public void performLoginLogout() throws IOException {
        String filePath = "LinksStatusClassification.xlsx"; // Update with your Excel file path

        // Read data from Excel
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0); // First sheet

        int rowCount = sheet.getPhysicalNumberOfRows();
        System.out.println("Number of Users in Excel: " + (rowCount - 1)); // Exclude header row

        for (int i = 1; i < rowCount; i++) { // Start from row 1 (skip header)
            Row row = sheet.getRow(i);
            String username = row.getCell(0).getStringCellValue();
            String password = row.getCell(1).getStringCellValue();

            System.out.println("Logging in with Username: " + username);

            // Perform login action
            driver.get("https://practice.expandtesting.com/login"); // Replace with the login URL
            WebElement emailField = driver.findElement(By.id("username")); // Replace with correct locator
            WebElement passwordField = driver.findElement(By.id("password")); // Replace with correct locator
            WebElement loginButton = driver.findElement(By.xpath("//*[@id=\"login\"]/button")); // Replace with correct locator

            emailField.clear();
            emailField.sendKeys(username);

            passwordField.clear();
            passwordField.sendKeys(password);

            loginButton.click();

            // Logout operation
            performLogout();
        }

        workbook.close();
        fis.close();
    }

    public void performLogout() {
        try {
            // Replace these locators based on your application
            WebElement profileIcon = driver.findElement(By.id("profileIcon")); // Example
            profileIcon.click();

            WebElement logoutButton = driver.findElement(By.linkText("Logout"));
            logoutButton.click();

            System.out.println("Logout successful\n");
        } catch (Exception e) {
            System.out.println("Error during logout: " + e.getMessage());
        }
    }

    @AfterTest
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }
    }
}
