package Proj.Propracs;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Delmul {

    WebDriver driver;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFCell cell;

    @SuppressWarnings("deprecation")
    @Test
    public void deleteImmigrationDetails() throws IOException, InterruptedException {

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\vaibh\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login"); // Replace with your login URL
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        // Import excel sheet
        File src = new File("C:\\Users\\vaibh\\eclipse-workspace\\Propracs\\muldel.xlsx");
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

            // Capture and write the title after login to Excel
            String loginTitle = driver.getTitle();
            System.out.println(loginTitle);
            cell = sheet.getRow(i).createCell(2); // Assuming you want to write the title in the third cell
            cell.setCellValue(loginTitle);

            // Click on MyInfo in the static side menu
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[1]/aside[1]/nav[1]/div[2]/ul[1]/li[6]/a[1]")).click();

            // Click on Immigration
            driver.findElement(By.xpath("//a[contains(text(),'Immigration')]")).click();

            // Get all checkboxes and select the first three
            List<WebElement> allCheckboxes = driver.findElements(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[5]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/label[1]/span[1]"));

            for (int j = 0; j < 3 && j < allCheckboxes.size(); j++) {
                allCheckboxes.get(j).click();
            }

            // Click on the Delete button
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[5]/div[2]/div[1]/div[1]/button[1]")).click();
            
            driver.findElement(By.xpath("//body/div[@id='app']/div[3]/div[1]/div[1]/div[1]/div[3]/button[2]")).click();

            // Wait for the page to reload after deletion, adjust the wait time as needed
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

            // Save the changes to the Excel file
            FileOutputStream fos = new FileOutputStream(src);
            workbook.write(fos);
            fos.close();
        }

        
    }
}
