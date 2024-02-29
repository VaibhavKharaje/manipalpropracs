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
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class Delete {

    WebDriver driver;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFCell cell;

    @Test
	@SuppressWarnings("deprecation")
    public void loginAndDeleteDocument() throws IOException {

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\vaibh\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login"); // Replace with your login URL
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(20, TimeUnit.MILLISECONDS);

        // Import excel sheet
        File src = new File("C:\\Users\\vaibh\\eclipse-workspace\\Propracs\\del.xlsx");
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
            driver.manage().timeouts().implicitlyWait(1000, TimeUnit.SECONDS);

            // Capture and write the title after login to Excel
            String loginTitle = driver.getTitle();
            System.out.println(loginTitle);
            cell = sheet.getRow(i).createCell(2); // Assuming you want to write the title in the third cell
            cell.setCellValue(loginTitle);

            // Navigate to MyInfo
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[1]/aside[1]/nav[1]/div[2]/ul[1]/li[6]/a[1]")).click();

            // Navigate to Immigration
            driver.findElement(By.xpath("//a[contains(text(),'Immigration')]")).click();

            // Click on the checkbox to delete the first document of passport
            WebElement deleteCheckbox = driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[3]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/label[1]/span[1]/i[1][1]"));
            deleteCheckbox.click();

            // Click on the delete button
            driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[3]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/button[1]/i[1]")).click();

            

            // Click on the Yes, delete button in the pop-up
            driver.findElement(By.xpath("//body/div[@id='app']/div[3]/div[1]/div[1]/div[1]/div[3]/button[2]")).click();

            // Wait for the page to reload after deletion, adjust the wait time as needed
            driver.manage().timeouts().implicitlyWait(1000, TimeUnit.SECONDS);

            // Capture and write the title after deletion to Excel
            String deletionTitle = driver.getTitle();
            System.out.println(deletionTitle);
            cell = sheet.getRow(i).createCell(3); // Assuming you want to write the title in the fourth cell
            cell.setCellValue(deletionTitle);

            // Save the changes to the Excel file
            FileOutputStream fos = new FileOutputStream(src);
            workbook.write(fos);
            fos.close();
        }

      
    }
}
