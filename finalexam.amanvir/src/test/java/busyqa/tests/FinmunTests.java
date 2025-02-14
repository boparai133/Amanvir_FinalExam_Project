package busyqa.tests;

import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

import busyqa.pages.FinmunWebpage;
import junit.framework.Assert;


import java.util.List;



import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

public class FinmunTests {
	
	private static final Logger logger = LogManager.getLogger(FinmunTests.class);
	
	public static ExtentSparkReporter sparkReporter;
	public static ExtentReports extent;
	public static ExtentTest test;
	WebDriver driver;
	int noOfTables = 5; //Number of tables to read
	XSSFWorkbook workbook = new XSSFWorkbook();
	
	public void initializer() {	
		logger.debug("Begin initializer()");
		//sparkReporter = new ExtentSparkReporter("./test-output/ExtentReport.html");
		sparkReporter =  new ExtentSparkReporter(System.getProperty("user.dir")+"/Reports/extentSparkReport.html");
		sparkReporter.config().setDocumentTitle("Automation Report");
		sparkReporter.config().setReportName("Test Execution Report");
		sparkReporter.config().setTheme(Theme.STANDARD);
		sparkReporter.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
		extent = new ExtentReports();
		extent.attachReporter(sparkReporter);			
		logger.debug("End initializer()");
	}
	public String CaptureScreenShot(WebDriver driver, String name, String fName) throws IOException {
		  String FileSeparator = System.getProperty("file.separator"); // "/" or "\"
		  String Extent_report_path = "."+FileSeparator+"Reports"; // . means parent directory	  
		  String Screenshotname = "screenshot_"+fName+".png";
		  File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);	  
		  File Dst = new File(Extent_report_path+FileSeparator+"Screenshots"+FileSeparator+name+FileSeparator+Screenshotname);	 
		  FileUtils.copyFile(scrFile, Dst);
		  String fPath = Dst.getAbsolutePath();
		  return fPath;
		  }
	  
  
	  
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
			          System.out.print("Find all the td tags in row: \t");
			          
			          // Iterate through columns and print data
			          try {
			          for (int j = 0; j < columns.size(); j++) {		        	 
			        	  	//xRow.createCell(j).setCellValue(columns.get(j).getText()); //working line of code
			        	  	XSSFCell cell = xRow.createCell(j);			        	  	
			        	  	cell.setCellValue(columns.get(j).getText());			        	  	
			        	  	
			        	  	CellStyle cellStyle = workbook.createCellStyle();			        	  		
		        	  		cellStyle.setAlignment(HorizontalAlignment.CENTER);		        	  		
		        	  		cell.setCellStyle(cellStyle);
		        	  		
		        	  		
			        	  	String colSpan = columns.get(j).getDomAttribute("colspan");
			        	  	System.out.print("td colspan value: "+ colSpan);			        	  	
			        	  	if(colSpan != null)
			        	  	{
			        	  		int colsp = Integer.parseInt(colSpan);
			        	  		setMerge(sheet,i,i,j,colsp,true);
			        	  		
			        	  	}
			        	  				        	  				        	  	
			        	  	setBordersToAllCells(sheet,i,i,j,j);			        	  	
			        	 }		
			          }
			          catch (Exception e) {
		        	  	test.log(Status.FAIL,"Failed at WriteToExcelFile(WebElement dt1, String name)");
		        	  	Assert.fail();
		        	  	System.out.println("Exception at: WriteToExcelFile(WebElement dt1, String name):"+ e.getMessage());	
			          }
			          
			      System.out.println(); // Move to the next row
		  	}			
	  }
	
	protected void setMerge(Sheet sheet, int numRow, int untilRow, int numCol, int untilCol, boolean border) {
	    CellRangeAddress cellMerge = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
	    sheet.addMergedRegion(cellMerge);
	    if (border) {
	        setBordersToMergedCells(sheet, cellMerge);
	    }

	}  

	protected void setBordersToMergedCells(Sheet sheet, CellRangeAddress rangeAddress) {
	    RegionUtil.setBorderTop(BorderStyle.DOUBLE, rangeAddress, sheet);
	    RegionUtil.setBorderLeft(BorderStyle.DOUBLE, rangeAddress, sheet);
	    RegionUtil.setBorderRight(BorderStyle.DOUBLE, rangeAddress, sheet);
	    RegionUtil.setBorderBottom(BorderStyle.DOUBLE, rangeAddress, sheet);	    
	}
	protected void setBordersToAllCells(Sheet sheet, int numRow, int untilRow, int numCol, int untilCol) {
		CellRangeAddress rangeAddress = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
	    RegionUtil.setBorderTop(BorderStyle.DOUBLE, rangeAddress, sheet);
	    RegionUtil.setBorderLeft(BorderStyle.DOUBLE, rangeAddress, sheet);
	    RegionUtil.setBorderRight(BorderStyle.DOUBLE, rangeAddress, sheet);
	    RegionUtil.setBorderBottom(BorderStyle.DOUBLE, rangeAddress, sheet);
	}

	public void SaveExcelToFileSystem() throws IOException
	  {
		  String FileSeparator = System.getProperty("file.separator"); // "/" or "\"
		  String Excel_File_path = "."+FileSeparator+"ExcelFiles"; // . means parent directory
		  String fname = "WinningTender.xlsx";	  
		  File Dst = new File(Excel_File_path + FileSeparator + fname);
		  String fPath = Dst.getAbsolutePath();	  
		  FileUtils.createParentDirectories(Dst);
		  System.out.println("Full file save path for excel file: "+ fPath);
		  try {
			  //FileOutputStream out = new FileOutputStream("c:\\WinningBidder.xlsx");
			  FileOutputStream out = new FileOutputStream(fPath);
			  
			  workbook.write(out);			  
			  out.close();
			  System.out.println("WinningTender.xlsx written successfully on disk.");		  
			} 
			catch (Exception e) {			 
				 test.log(Status.FAIL,"Failed at SaveExcelToFileSystem()");
				 Assert.fail();
				 System.out.println("Exception::"+ e.getMessage());			 		  
			  	}
	  }
	  
	 
	
	  @Test(enabled = true)
	  public void TestCopyWebtablesToExcelFile() throws IOException, InterruptedException{
		  String methodName = new Exception().getStackTrace()[0].getMethodName();
		  System.out.println("methodName::"+ methodName);
		  logger.debug("Test method::"+ methodName);
		  
		  String className = new Exception().getStackTrace()[0].getClassName();
		  test = extent.createTest(methodName,"Create Excel file from webtables");
		  test.log(Status.INFO, "Starting test TestCopyWebtablesToExcelFile()");
		  test.assignCategory("Regression Testing");
		  
		  logger.debug("Starting test TestCopyWebtablesToExcelFile()");
		  
		  System.out.println("This is TestCopyWebtablesToExcelFile()");		  
		  FinmunWebpage finmunPage = new FinmunWebpage(driver);
		  test.log(Status.INFO, "TestCopyWebtablesToExcelFile(): Webdriver Initialized");
		  logger.debug("TestCopyWebtablesToExcelFile(): Webdriver Initialized");
		  
		  ///Click on Bond Results Tab
		  finmunPage.clickTabBonds();
		  test.log(Status.INFO, "Clicked on Tab :: Results for the last 90 days - Bonds");
		  logger.debug("Clicked on Tab :: Results for the last 90 days - Bonds");
		  Thread.sleep(2000);
		  
		  
		  
		  for(int t=1;t<=noOfTables;t++)
		  {
			  //Read the table
			  WebElement dt = finmunPage.getTable(t);
			  test.log(Status.INFO, String.format("Started reading of webtable [%d]", t));
			  logger.debug(String.format("Started reading of webtable [%d]", t));
			  System.out.println("\r Table : "+ t );
			  // Find all rows in the table
		      List<WebElement> rows = dt.findElements(By.tagName("tr"));
	
		      // Print the number of rows in the table
		      System.out.println("Number of rows in the table= " + rows.size());	
		      logger.debug("Number of rows in the table= " + rows.size());
			  Thread.sleep(3000);
			  
				// Iterate through all the rows in table and print data.
			  	for (int i = 0; i < rows.size()-2; i++) {
			          WebElement row = rows.get(i);
			          
			          //Find column with <a> tag in the row
			          
			          List<WebElement> columns = row.findElements(By.xpath("td/a"));
			          System.out.print("Loop through first column with <a> tag: \t");
			          logger.debug("Loop through first column with <a> tag: \t");
			          
			          // Iterate through columns and print data
			          try {
				          for (int j = 0; j < 1; j++) {
				        	  
				        	  //System.out.print(columns.get(j).findElement(By.tagName("a")).getDomAttribute("href") + "\t");
				        	  WebElement aLink = columns.get(j);
				        	  System.out.print(aLink.getText() + "\t");				        	  
				        	  String name = aLink.getText();			
				        	  
				        	  logger.debug("Capturing screen shot of main page");			        	  
				        	  test.addScreenCaptureFromPath(CaptureScreenShot(driver,name,"MainPage")); //Step a
				        	  aLink.click();
				        	  Thread.sleep(3000);
				        	  
				        	  
				        	  logger.debug("Capturing screen shot of detail page of "+ name);	
				        	  test.addScreenCaptureFromPath(CaptureScreenShot(driver,name,"PopupWindow")); //Step c
				        	  
				        	  finmunPage.switchToIframe();
				        	  Thread.sleep(3000);
				        	  
				        	  
				        	  WebElement popDT = finmunPage.getPopupTable();
				        	  
				        	  Thread.sleep(3000);  
				        	  //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
				        	  
				        	  //Read all the rows of the web table
				  			  List<WebElement> rows1 = popDT.findElements(By.tagName("tr"));
				  			
				  		      // Print the number of rows in the table
				  		      System.out.println("Number of rows in the table= " + rows1.size());
				  		      Thread.sleep(3000);
				  		      
				  		      test.log(Status.INFO,"CreatingExcelSheet of "+ name);
				  		      CreateExcelSheet(popDT, name);
				        	  Thread.sleep(2000);
				        	  				  		      
				        	  finmunPage.switchToMainContent();
				        	  Thread.sleep(2000);
				        	  
				        	  test.log(Status.INFO,"Detailed table of "+ name +" closed");
				        	  finmunPage.closeAlert();				        	         	  
				        	  Thread.sleep(2000);
				        	  
				          }		
			          }
			          catch (Exception e) {
		        	  		test.log(Status.FAIL,"Exception block: Failed at TestCopyWebtablesToExcelFile()");
		        	  		logger.error("Exception block: Failed at TestCopyWebtablesToExcelFile()");
			        	  	Assert.fail();
			        	  	System.out.println("Exception block: TestCopyWebtablesToExcelFile():"+ e.getMessage());	
						}
			          
			      System.out.println(); // Move to the next row
			  	}
			   
			  		SaveExcelToFileSystem();
			  		test.log(Status.INFO,"Excel file with all the sheets saved");
			  		logger.debug("Excel file with all the sheets saved");
		 }	 
		  
  	}	  
 
	  @BeforeTest
	  public void driverSetup() {
		  System.out.println("This is @BeforeTest method driverSetup()");	  
		  initializer();	   
		  driver = new ChromeDriver();
		  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3003::RESLT:");	  	  	   
		  driver.manage().window().maximize();
		  driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(2));
	  }
	  @AfterTest
	  public void closeMethod() {
		  System.out.println("This is @afterTest method closeMethod()");  
		  extent.flush();
		  driver.quit();
	  }
 }
