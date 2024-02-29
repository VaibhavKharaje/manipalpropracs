package Proj.Propracs;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class upload {

    WebDriver driver;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFCell cell;

    @SuppressWarnings("deprecation")
    @Test
    public void fillImmigrationDetails() throws IOException, InterruptedException {

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\vaibh\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login"); // Replace with your login URL
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        // Import excel sheet
        File src = new File("C:\\Users\\vaibh\\eclipse-workspace\\Propracs\\up.xlsx");
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

            // Click on the Add button
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[5]/div[1]/div[1]/button[1]")).click();

            // Click on the browser button to select and upload a file
            WebElement fileInput = driver.findElement(By.xpath("//div[contains(text(),'No file selected')]"));
            // Replace the below string with the absolute path of the file you want to upload
            fileInput.sendKeys("C:\\\\Users\\\\vaibh\\\\OneDrive\\\\Pictures\\\\Screenshots\\\\Screenshot (87).png");
		   

            // Enter comments from Excel
            cell = sheet.getRow(i).getCell(15);
            String commentText = cell.getStringCellValue();
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[5]/div[1]/form[1]/div[2]/div[1]/div[1]/div[1]/div[2]/textarea[1]")).sendKeys(commentText);

            // Capture and write the title after entering Comment to Excel
            String commentTitle = driver.getTitle();
            System.out.println(commentTitle);
            cell = sheet.getRow(i).createCell(16);
            cell.setCellValue(commentTitle);

            // Click on the Save button
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[5]/div[1]/form[1]/div[3]/button[2]")).click();

            // Wait for the page to reload after saving, adjust the wait time as needed
            //WebDriverWait wait = new WebDriverWait(driver, 10);
            //wait.until(ExpectedConditions.titleContains("Saved"));

            // Save the changes to the Excel file
            FileOutputStream fos = new FileOutputStream(src);
            workbook.write(fos);
            fos.close();
        }

      
    }
}
