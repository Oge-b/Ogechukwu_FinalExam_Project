package BusyQA.FinalExamProject;


import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.List;

	public class FinalExamPractice {

		 WebDriver driver;
		    Workbook workbook;
		    Sheet sheet;

		    @BeforeTest
		    public void beforeTest() throws InterruptedException {
		        driver = new ChromeDriver();
		        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		        driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
		        workbook = new XSSFWorkbook();
		    }
		    
		    @DataProvider(name = "linkDataProvider")
		    public Object[][] linkDataProvider() {
		        return new Object[][] {
		            { 2, 1 }, // Test 1
		            { 2, 2 }, // Test 2
		            { 2, 3 }, // Test 3
		            { 3, 1 }, // Test 4
		            { 3, 2 }  // Test 5
		        };
		    }

		    @Test(dataProvider = "linkDataProvider")
		    public void Test1(int tbody, int row) throws InterruptedException {
		        // Identify the table
		    	 WebElement linkTitle = driver.findElement(By.xpath("//*[@id='report_BILLETS']/div/div[1]/table/tbody[" + tbody + "]/tr[" + row + "]/td[1]/a"));
		    	 String tabName = linkTitle.getText();
			        linkTitle.click();
			        System.out.println("Title of link is: " + tabName);
		        Thread.sleep(2000);
		        // Create a new workbook and a sheet with the title "Tabname"
		        
		        
		        
		        sheet = workbook.createSheet(tabName);

		        // Locate the table rows
		        List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"R1740668184739222315\"]/div[2]/div[2]/table[2]"));
		        Thread.sleep(2000);
		        // Iterate through rows and columns to fill the sheet
		        for (int i = 0; i < rows.size(); i++) {
		            Row tablerow = sheet.createRow(i);
		            List<WebElement> cells = rows.get(i).findElements(By.tagName("strong"));
		            for (int j = 0; j < cells.size(); j++) {
		                tablerow.createCell(j).setCellValue(cells.get(j).getText());
		            }
		        }

		        // Save the workbook to a file
		        try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\User\\ExcelABC\\FinalExamProject.xlsx")) {
		            workbook.write(fileOut);
		        } catch (IOException e) {
		            e.printStackTrace();
		        }

		        // Close the modal or page
		        WebElement closePageButton = driver.findElement(By.xpath("//button[@title='Fermer']"));
		        closePageButton.click();
		        Thread.sleep(2000);
		    }
		    
		    @AfterTest
		    public void close() {
		        // Close the workbook
		        if (workbook != null) {
		            try {
		                workbook.close();
		            } catch (IOException e) {
		                e.printStackTrace();
		            }
		        }
		        driver.close();
		    }
	}



