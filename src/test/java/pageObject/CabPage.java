package pageObject;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.*;
import org.openqa.selenium.support.FindBy;
//import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class CabPage extends BasePage{

	WebDriverWait wait;
	public CabPage(WebDriver driver) {
		super(driver);
		this.wait = super.myWait;
	}
	
	//Car filter Element
	@FindBy(xpath="//div[@class='appliedFilterButton']/p[text()='SUV']")
	WebElement cabFilter;
	
	//Round trip element 
	@FindBy(xpath="//li[@data-cy='outstationRoundTrip']")
	WebElement roundtripButton;
	
	JavascriptExecutor j=(JavascriptExecutor)driver;
	
	// sort button Element
	@FindBy(xpath="//span[@class='cursorPointer dodgerBlueColor']")
	WebElement sort_button;
	
	// low-high button Element
	@FindBy(xpath="//span[text()= ' (Lowest to Highest)']")
	WebElement low_rate_button;
	
	//SUV button element
	@FindBy(xpath="//div[@class='appendBottom30']/div/span/label[text()='SUV']")
	WebElement suv_button;
	
	// Lowest Value Element
	@FindBy(xpath="//*[@id='List']/div[1]/div/div[3]/div/div[2]/div/p[1]")
	WebElement low_price;
	
	// More Button Element
	@FindBy(xpath=("//span[@class='darkGreyText']"))
	WebElement more_button;
	
	// Gift Card Button
	@FindBy(xpath=("//a[text()='Giftcards']"))
	WebElement gift_button;
	
	//Price list Elements
	@FindBy(xpath="//*[@id=\"List\"]/div/div/div[3]/div/div[2]/div/p[1]")
	List<WebElement> price_list;
	
	// covered distance text
	@FindBy(xpath="//*[@id=\"root\"]//span[contains(text(),'Distance')]")
	WebElement cab_distance;
	
	// car model checkBox
	@FindBy(xpath="//div/p[text()='Cab Model']/following::div[@data-testid='filter_selection']/span")
	List<WebElement> cab_Model;
	
	//Check the is car displayed or not method
	public boolean isCabDistDisplayed() throws Exception {
		//Thread.sleep(4000);
		boolean distance=cab_distance.isDisplayed();
		return distance;
		
	}
	
	//Method for price list
	public List<String> priceList(){
		
		List<String> listprice=new ArrayList<>();
		for(WebElement price:price_list) {
			listprice.add(price.getText());
		}
		
		return listprice;
	}

	//method for sort button 
	public void sortButton() {
		sort_button.click();
	}
	
	// methods for filter low rate to high rate
	public void lowRatebutton() {
		
	
		j.executeScript("arguments[0].click()", low_rate_button);
	}
	
	//Is low rate button method
	public boolean isLowRatButnSelected() {
		
		
		boolean val=low_rate_button.isDisplayed();
		return val;
	}
	
	//click the Car SUV button method
	public void carSuv() {
		
		j.executeScript("arguments[0].click()", suv_button);
	}
	
	//Click the car filter method
	public boolean isCabFilterApplied() throws Exception {
		
		Thread.sleep(3000);
		boolean val=cabFilter.isDisplayed();
		return val;
	}
	
	//lowest price method
	public String lowestPrice() {
		
		String lowprice=low_price.getText();
		return lowprice;
	}
	
	//Click the more button method
	public void moreButton() {
		//WebElement more = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[@class='darkGreyText']")));
		more_button.click();
		//more.click();
	}
	
	// Gift button method
	public void giftButton() {
		gift_button.click();
	}
	
	// Round trip button method
	public void roundtripRButton() {
		roundtripButton.click();
	}
	
	//method to check the round trip button enabled or not
	public boolean isroundtripRButtonEnabled() {
		boolean val=roundtripButton.isEnabled();
		return val;
	}
	
	//Check the car model is enabled or not method
	public boolean isCabModelEnabled()
	{
		int i=0;
		
		for(WebElement cab:cab_Model)
		{
			if(cab.isEnabled()) {
				i++;
				break;
			}
		}
		if(i>0) 
		{
			return true;
		}
		
		return false;
	}
	
}