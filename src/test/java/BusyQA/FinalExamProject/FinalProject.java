package BusyQA.FinalExamProject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class FinalProject {
    
    WebDriver driver;
    ExtentReports extent;
    ExtentTest test;
    static Logger logger = Logger.getLogger(FinalProject.class);


    @Test
    public void Test1() throws InterruptedException {
        String excelFilePath = "C:\\Users\\User\\ExcelABC\\FinalExam.xlsx";
        
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            int count = 0;

            while (count < 5) {
                List<WebElement> tableRows = driver.findElements(By.xpath("//table[contains(@class, 't-Report-report')]/tbody/tr"));
                for (WebElement row : tableRows) {
                    List<WebElement> nameLinks = row.findElements(By.xpath(".//td[1]/a"));
                    if (!nameLinks.isEmpty()) {
                        WebElement nameLink = nameLinks.get(0);
                        if (nameLink.isDisplayed()) {
                            String tabName = nameLink.getText();
                            
                            // Take screenshot before clicking the link
                            String beforeScreenshotPath = takeScreenshot("before_" + tabName);
                            test = extent.createTest("Test for " + tabName);
                            test.addScreenCaptureFromPath(beforeScreenshotPath, "Before clicking link");
                            logger.info("Logger File Recorded Successfully");
                            nameLink.click();
                            Thread.sleep(1000);
                            
                            
                         // Take screenshot after clicking the link
                            String afterScreenshotPath = takeScreenshot("after_" + tabName);
                            test.addScreenCaptureFromPath(afterScreenshotPath, "After clicking link");
                            
                            
                            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                            wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.tagName("iframe")));

                            List<WebElement> iframeTable = driver.findElements(By.xpath("//*[@id=\"R1740668184739222315\"]/div[2]/div[2]/table[2]/tbody/tr"));
                            XSSFSheet sheet = workbook.createSheet(tabName);

                            int rowIndex = 0;
                            for (WebElement iframeRow : iframeTable) {
                                XSSFRow excelRow = sheet.createRow(rowIndex++);
                                List<WebElement> cells = iframeRow.findElements(By.tagName("td"));
                                for (int i = 0; i < cells.size(); i++) {
                                    XSSFCell cell = excelRow.createCell(i);
                                    cell.setCellValue(cells.get(i).getText());
                                }
                            }
                            driver.switchTo().defaultContent();
                            driver.findElement(By.xpath("//*[@id=\"t_PageBody\"]/div[2]/div[1]/button")).click();
                            count++;
                            if (count >= 5) break;
                        }
                    }
                }
            }

            // Write to the Excel file 
            try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    // Method to take screenshot
    public String takeScreenshot(String filename) {
        try {
            TakesScreenshot ts = (TakesScreenshot) driver;
            File source = ts.getScreenshotAs(OutputType.FILE);
            
            // Construct the destination path using user.dir
            String destination = System.getProperty("user.dir") + "\\Reports\\Screenshots\\" + filename + ".png"; 
            Path destinationPath = Paths.get(destination);
            
            
            Files.createDirectories(destinationPath.getParent());
            Files.copy(source.toPath(), destinationPath);
            return destination;
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
    }



           
    @BeforeTest
    public void beforeTest() throws InterruptedException {
        // Initialize the WebDriver
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
        PropertyConfigurator.configure("src\\test\\resources\\log4j.properties");

        
        ExtentSparkReporter sparkReporter = new ExtentSparkReporter(System.getProperty("user.dir") + "\\Reports\\extentSparkReport.html");
        sparkReporter.config().setDocumentTitle("Test Report");
        sparkReporter.config().setReportName("Functional Test Results");
        sparkReporter.config().setTheme(Theme.DARK);

        extent = new ExtentReports();
        extent.attachReporter(sparkReporter);
        
    }
    
    @AfterTest
    public void close() {
       
        driver.close();
        driver.quit();
    }
}
