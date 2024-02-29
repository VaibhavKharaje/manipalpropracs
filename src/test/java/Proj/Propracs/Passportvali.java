package Proj.Propracs;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class Passportvali {

    WebDriver driver;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFCell cell;

    @SuppressWarnings("deprecation")
    @Test
    public void loginAndPassportValidation() throws IOException {

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\vaibh\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login"); // Replace with your login URL
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(100, TimeUnit.MILLISECONDS);

        // Import excel sheet
        File src = new File("C:\\\\Users\\\\vaibh\\\\eclipse-workspace\\\\Propracs\\\\pass.xlsx");
        // load the file
        FileInputStream fis = new FileInputStream(src);
        // load the work book
        workbook = new XSSFWorkbook(fis);
        // access the sheet from the work book
        sheet = workbook.getSheetAt(0);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            // import the data for username
            cell = sheet.getRow(i).getCell(0);
            String username = cell.getStringCellValue();

            // import the data for password
            cell = sheet.getRow(i).getCell(1);
            String password = cell.getStringCellValue();

            // Log in
            driver.findElement(By.xpath("//input[@name='username']")).sendKeys(username);
            driver.findElement(By.xpath("//input[@name='password']")).sendKeys(password);
            driver.findElement(By.xpath("//button[@type='submit']")).click();

            // Wait for login to complete, adjust the wait time as needed
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

            // Check if login was successful
            if (driver.getCurrentUrl().contains("dashboard")) {
                // Login successful
                cell = sheet.getRow(i).createCell(2);
                cell.setCellValue("Login Successful");
                
                String title = driver.getTitle();
                System.out.println(title);
                cell = sheet.getRow(i).createCell(3); // Assuming you want to write the title in the fourth cell
                cell.setCellValue(title);

                // Navigate to MyInfo
                driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[1]/aside[1]/nav[1]/div[2]/ul[1]/li[6]/a[1]")).click();

                // Navigate to Immigration
                driver.findElement(By.xpath("//a[contains(text(),'Immigration')]")).click();

                // Click on the Add button
                driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/button[1]")).click();

                // Wait for the Add page to load
                driver.manage().timeouts().implicitlyWait(500, TimeUnit.SECONDS);

                // Import the data for passport number
                cell = sheet.getRow(i).getCell(3);
                String passportNumber = cell.getStringCellValue();
                driver.manage().timeouts().implicitlyWait(500, TimeUnit.SECONDS);

                // Enter passport number
                driver.findElement(By.xpath("//body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/div[1]/div[1]/div[1]/div[2]/input[1]")).sendKeys(passportNumber);

                // Perform further actions...

                // You can add validation or other actions as needed

                
                // After performing actions, you can write the result back to Excel if required
                driver.manage().timeouts().implicitlyWait(500, TimeUnit.SECONDS);
                cell = sheet.getRow(i).createCell(4);
                // Assuming you want to write the result in the fifth cell
                cell.setCellValue("Validation Passed");

                // Click on the Save button
                driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/button[2]")).click();

                // Save the changes to the Excel file
                FileOutputStream fos = new FileOutputStream(src);
                workbook.write(fos);
                fos.close();
            } else {
                // Login failed
                cell = sheet.getRow(i).createCell(2);
                cell.setCellValue("Login Failed");

                // Save the changes to the Excel file
                FileOutputStream fos = new FileOutputStream(src);
                workbook.write(fos);
                fos.close();
            }
        }

        
    }
}
