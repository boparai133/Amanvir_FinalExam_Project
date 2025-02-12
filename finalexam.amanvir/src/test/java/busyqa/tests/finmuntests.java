package busyqa.tests;

import org.testng.annotations.Test;

import busyqa.pages.finmunwebpage;

import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;

import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class finmuntests {
	WebDriver driver;
	  @Test(enabled = true)
	  public void Test1() throws InterruptedException {
		  
		  System.out.println("This is Test1()");
		  finmunwebpage finmunPage = new finmunwebpage(driver);
		  
		  ///Click on Bond Results Tab
		  finmunPage.clickTabBonds();
		  Thread.sleep(2000);
		  int noOfTables = 5;
		  
		  for(int t=1;t<=noOfTables;t++)
		  {
			  //Read the table
			  WebElement dt = finmunPage.getTable(t);			  
			  System.out.println("\r Table : "+ t );
			  // Find all rows in the table
		      List<WebElement> rows = dt.findElements(By.tagName("tr"));
	
		      // Print the number of rows in the table
		      System.out.println("Number of rows in the table= " + rows.size());
			  Thread.sleep(3000);
			  
				// Iterate through all the rows in table and print data.
			  	for (int i = 0; i < rows.size()-2; i++) {
			          WebElement row = rows.get(i);
			          // Find all columns in the row
			          //List<WebElement> columns = row.findElements(By.tagName("td"));
			          List<WebElement> columns = row.findElements(By.xpath("td/a"));
			          System.out.print("Loop through first column with <a> tag: \t");
			          // Iterate through columns and print data
			          try {
				          for (int j = 0; j < 1; j++) {
				        	  //System.out.print(columns.get(j).findElement(By.tagName("a")).getDomAttribute("href") + "\t");
				        	  WebElement aLink = columns.get(j);
				        	  System.out.print(aLink.getText() + "\t");
				        	  String name = aLink.getText();
				        	  aLink.click();
				        	  Thread.sleep(3000);
				        	  
				        	  finmunPage.switchToIframe();
				        	  Thread.sleep(3000);
				        	  
				        	  WebElement popDT = finmunPage.getPopupTable();		        	  		        	  	        	
				        	  Thread.sleep(3000);  
				        	  
				        	  //Read all the rows of the web table
				  			  List<WebElement> rows1 = popDT.findElements(By.tagName("tr"));
				  			
				  		      // Print the number of rows in the table
				  		      System.out.println("Number of rows in the table= " + rows1.size());
				  		      Thread.sleep(3000); 
				  		      
				  		      //WriteToExcelFile(popDT,name);
				        	  //writeIntoExcelFile(popDT,name);
				  		      CreateExcelSheet(popDT, name);
				        	  Thread.sleep(2000);
				  		      
				        	  finmunPage.switchToMainContent();
				        	  Thread.sleep(2000);
				        	  
				        	 
				        	  
				        	  finmunPage.closeAlert();				        	         	  
				        	  Thread.sleep(5000);
				          }		
			          }
			          catch (Exception e) {
							// TODO: handle exception
							System.out.print("Exception: " + e.getMessage());
						}
			          
			      System.out.println(); // Move to the next row
			  	}
			  	SaveExcelToFileSystem();
		 }	 
		  
  	}	  
  
  @BeforeMethod
  public void beforeMethod() {
  }

  @AfterMethod
  public void afterMethod() {
  }

  @BeforeClass
  public void beforeClass() throws InterruptedException {
	  System.out.println("This is @beforeClass");
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3003::RESLT:");
	  driver.manage().window().maximize();
	  Thread.sleep(3000);	   
  }

  @AfterClass
  public void afterClass() {
	  System.out.println("This is @afterClass");
	  driver.quit();
  }

  @BeforeTest
  public void beforeTest() {
	  System.out.println("This is @beforeTest");
  }

  @AfterTest
  public void afterTest() {
	  System.out.println("This is @afterTest");
  }
  
  
  XSSFWorkbook workbook = new XSSFWorkbook();
  public void CreateExcelSheet(WebElement dt1, String name)  {
	   
			XSSFSheet sheet = workbook.createSheet(name);

				
			//Read all the rows of the web table
			List<WebElement> rows = dt1.findElements(By.tagName("tr"));
			
		      // Print the number of rows in the table
		      System.out.println("Number of rows in the table= " + rows.size());
			  			  
				// Iterate through all the rows in table and print data.
			  	for (int i = 0; i < rows.size(); i++) {
			          WebElement row = rows.get(i);
			          XSSFRow xRow = sheet.createRow(i);
			          // Find all columns in the row
			          List<WebElement> columns = row.findElements(By.tagName("td"));
			          //List<WebElement> columns = row.findElements(By.xpath("td"));
			          System.out.print("Loop through first column with <a> tag: \t");
			          // Iterate through columns and print data
			          try {
			          for (int j = 0; j < columns.size(); j++) {		        	 
			        	  xRow.createCell(j).setCellValue(columns.get(j).getText());  			
			          }		
			          }
			          catch (Exception e) {
							// TODO: handle exception
							System.out.print("Exception::WriteToExcelFile(WebElement dt1, String name):: " + e.getMessage());
						}
			          
			      System.out.println(); // Move to the next row
			  	}
			//Write the workbook in file system

  }
  
  public void SaveExcelToFileSystem()
  {
	  try {
		  FileOutputStream out = new FileOutputStream("c:\\WinningBidder.xlsx");
		  workbook.write(out);			  
		  out.close();
		  System.out.println("WinningBidder.xlsx written successfully on disk.");
		} 
		catch (Exception e) {
			 System.out.println("Exception 1");
		  e.printStackTrace();
		}
  }
  
  public void WriteToExcelFile(WebElement dt1, String name) {
		//Blank workbook

	  System.out.println("WriteToExcelFile() ");
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet = workbook.createSheet(name);

				
			//Read all the rows of the web table
			List<WebElement> rows = dt1.findElements(By.tagName("tr"));
			
		      // Print the number of rows in the table
		      System.out.println("Number of rows in the table= " + rows.size());
			  			  
				// Iterate through all the rows in table and print data.
			  	for (int i = 0; i < rows.size(); i++) {
			          WebElement row = rows.get(i);
			          XSSFRow xRow = sheet.createRow(i);
			          // Find all columns in the row
			          List<WebElement> columns = row.findElements(By.tagName("td"));
			          //List<WebElement> columns = row.findElements(By.xpath("td"));
			          System.out.print("Loop through first column with <a> tag: \t");
			          // Iterate through columns and print data
			          try {
			          for (int j = 0; j < columns.size(); j++) {		        	 
			        	  xRow.createCell(j).setCellValue(columns.get(j).getText());  			
			          }		
			          }
			          catch (Exception e) {
							// TODO: handle exception
							System.out.print("Exception::WriteToExcelFile(WebElement dt1, String name):: " + e.getMessage());
						}
			          
			      System.out.println(); // Move to the next row
			  	}
			//Write the workbook in file system

			try {
			  FileOutputStream out = new FileOutputStream("c:\\WinningBidder.xlsx");
			  workbook.write(out);			  
			  out.close();
			  System.out.println("WinningBidder.xlsx written successfully on disk.");
			} 
			catch (Exception e) {
				 System.out.println("Exception 1");
			  e.printStackTrace();
			}
		} catch (IOException e) {
			System.out.println("Exception 2");
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 

		//Create a blank sheet

	}
}
