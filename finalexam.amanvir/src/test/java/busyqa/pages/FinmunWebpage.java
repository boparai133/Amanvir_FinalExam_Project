package busyqa.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import java.lang.*;
import java.util.List;

public class FinmunWebpage {

	WebDriver driver;
	public FinmunWebpage(WebDriver driver) {
		// TODO Auto-generated constructor stub
		this.driver = driver;
	}
	
	//Locator for Table2 field
	By tabBonds = By.xpath("//*[@id=\"OBLIGATIONS_tab\"]/a");

	
	//Alert Close button
	By alertCloseButton = By.cssSelector("button[title=\"Fermer\"]");
	
	//Pop up Table locator	
	By popupTable = By.xpath("//*[@id=\"R1469412186955323305\"]/div[2]/div[2]/table[2]");
	
	//find iframe
	By ifrm = By.tagName("iframe");
		
	//Method to return tab bonds
	public WebElement  getTabBonds() {
		return driver.findElement(tabBonds);
	}
	
	//Method to click on Bonds Tab
	public void clickTabBonds() {
		driver.findElement(tabBonds).click();		
	}	
	
	//Method to read web page table
	public WebElement getTable(int tableNumber) {
		By webelement=null;
		try
		{
			 webelement = By.xpath(String.format("//*[@id=\"report_OBLIGATIONS\"]/div/div[1]/table/tbody[%d]",tableNumber+1));
		}
		catch(Exception e)
		{
			throw e;
		}
		return driver.findElement(webelement);
	}	
	 
	//Method to close the pop up window
	public void closeAlert()
	{
		driver.findElement(alertCloseButton).click();
	}
	
	//Method to click on Bonds Tab
	public WebElement getPopupTable1() {		
		By webelement=null;
		try
		{
			 webelement = By.xpath("//*[@id=\"R1469412186955323305\"]/div[2]/div[2]/table[2]");
		}
		catch(Exception e)
		{
			throw e;
		}
		return driver.findElement(webelement);
	}
	
	//Method to click on Bonds Tab
	public WebElement getPopupTable() {		 
		return driver.findElement(popupTable);
	}
	
	//Method to return tab bonds
	public WebElement  getChildiFrame() {
		return driver.findElement(ifrm);
	}
	
	public void switchToIframe() {
		try
		{
		driver.switchTo().frame(getChildiFrame());
		System.out.println("switchToIframe()");
		}
		catch(Exception e)
		{
			throw e;
			//System.out.println("Exception::switchToIframe()::" + e.getMessage());
		}
		
	}
	
	public void switchToMainContent() {
		try
		{
		driver.switchTo().defaultContent();
		System.out.println("switchToMainContent()");
		}
		catch(Exception e)
		{
			throw e;
			//System.out.println("Exception::switchToIframe()::" + e.getMessage());
		}
	}
	 
}
