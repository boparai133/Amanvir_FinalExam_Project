package busyqa.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class FinmunWebpage {

	WebDriver driver;
	public FinmunWebpage(WebDriver driver) {		
		this.driver = driver;
		PageFactory.initElements(driver,this);
	}
	
	@FindBy(xpath = "//*[@id='OBLIGATIONS_tab']/a") WebElement tabBonds;
	@FindBy(css = "button[title='Fermer']") WebElement alertCloseButton;
	@FindBy(xpath = "//*[@id='R1469412186955323305']/div[2]/div[2]/table[2]") WebElement popupTable;
	@FindBy(tagName = "iframe") WebElement ifrm;
		
	//Method to click on Bonds Tab
	public void clickTabBonds() {		
		tabBonds.click();
	}	
	
	//Method to read main page table 
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
		alertCloseButton.click();
	}
		
	//Method to click on Bonds Tab
	public WebElement getPopupTable() {		 
		return popupTable;
	}
			
	//Method to switch to child frame
	public void switchToIframe() {
		try
		{
		driver.switchTo().frame(ifrm);	
		System.out.println("switchToChildIframe()");
		}
		catch(Exception e)
		{
			throw e;			
		}
		
	}
	
	//Method to switch to Main content
	public void switchToMainContent() {
		try
		{
		driver.switchTo().defaultContent();
		System.out.println("switchToMainContent()");
		}
		catch(Exception e)
		{
			throw e;			
		}
	}
	 
}
