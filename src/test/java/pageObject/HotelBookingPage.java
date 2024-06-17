package pageObject;

import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.*;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class HotelBookingPage extends BasePage{

	public HotelBookingPage(WebDriver driver) {
		super(driver);
	}

	JavascriptExecutor j= (JavascriptExecutor)driver;
	
	// Hotel Button in home Page
	@FindBy(xpath=("(//span[text()='Hotels'])[1]"))
	WebElement hotel_button;

	
	// Location Button in hotel Page
	@FindBy(xpath=("//*[@id=\"top-banner\"]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/label"))
	WebElement location_button;
	
	// Input in Location 
	@FindBy(xpath=("//input[@placeholder='Where do you want to stay?']"))
	WebElement location_input;


	// suggestion Location
	@FindBy(xpath=("//*[@id='react-autowhatever-1-section-0-item-0']"))
	WebElement suggestion_location;
	
	//Check in Button
	@FindBy(xpath=("//input[@id='checkin']"))
	WebElement checkin_button;
	
	//Check out Button
	@FindBy(xpath=("//input[@id='checkout']"))
	WebElement checkout_button;
	
	//Adult Button
	@FindBy(xpath=("(//div[@class='gstSlctCont'])[2]"))
	WebElement adult_button;
	
	@FindBy(xpath="//*[@id=\"guest\"]")
	WebElement room_guest;
	
	@FindBy(xpath="//*[@id=\"POPULAR\"]/ul/li[1]/span/label/div/span")
	WebElement select_filter;
	
	@FindBy(xpath="//*[@id=\"POPULAR\"]/ul/li[3]/span")
	WebElement check_filter;
	
	@FindBy(xpath="//button[normalize-space()='Apply']")
	WebElement apply_button;
	
	@FindBy(xpath="//*[@id=\"hsw_search_button\"]")
	WebElement search_button;
	
	//List of Adult in DropDown
	@FindBy(xpath=("//ul[@class='gstSlct__list']/li"))
	List<WebElement> adult_list;

	@FindBy(xpath="//*[@id=\"top-banner\"]/div[2]/div/div[1]/div[2]/div/div[1]/div[4]/div[1]/div[1]/div[1]/div[2]/div")
	WebElement room_button;
	
	@FindBy(xpath="//*[@id=\"top-banner\"]/div[2]/div/div[1]/div[2]/div/div[1]/div[4]/div[1]/div[1]/div[1]/div[2]/ul/li/text()[2]")
	List<WebElement> room_members;
	
	@FindBy(xpath="//*[@id=\"top-banner\"]/div[2]/div/div[1]/div[2]/div/div[1]/div[5]/label/span")
	WebElement price_per_night_button;
	
	@FindBy(xpath="//*[@id=\"top-banner\"]/div[2]/div/div[1]/div[2]/div/div[1]/div[5]/div[1]/ul/li[1]")
	WebElement rate_first_low;
	
	@FindBy(xpath="//*[@id=\"hotelListingContainer\"]/div/div/p[3]")
	WebElement result_text;
	
	//search button selecting method
	public void searchButton() {
		
		j.executeScript("arguments[0].click()",search_button);
	}
	//applying the values method
	public void applyButton() {
		apply_button.click();
	}
	//method to check the filterApplied 
	public boolean isFilterApplied() {
		
		j.executeScript("arguments[0].scrollIntoView(true);", check_filter);
		boolean val=check_filter.isDisplayed();
		return val;
	}
	//method to select Filter 
	public void selectFilter() {
		select_filter.click();
	}
	//method to click room guest
	public void roomGuest() {
		room_guest.click();
	}
	
	
	//method to click on the adult value
	public void adultList(int i){
		
		WebElement adultno=driver.findElement(By.xpath("//ul[@class='gstSlct__list']/li["+i+"]"));
		adultno.click();
	}
	
	
	//method to click on the hotel option
	public void hotelButton() {
		hotel_button.click();
	}
	
	//method to select location
	public void locationButton() {
		location_button.click();
	}
	
	//method to send to location value
	public void locationName(String area) {
		
		location_input.sendKeys(area);
	}
	
	//method to select the suggestion value
	public void suggestionLocation() throws InterruptedException {
		Thread.sleep(2000);
		suggestion_location.click();
	}
	
	//method to click on check In button
	public void checkInButton() {
		checkin_button.click();
	}
	//method to select the date
	public void checkInDate(String checkInDate) {
		
	
		WebElement indate=driver.findElement(By.xpath("//*[contains(@class,'DayPicker-Day')][@tabindex=-1][@aria-disabled='false'][contains(@aria-label,'"+checkInDate+"')]"));
		
		try {
			Thread.sleep(3000);
		} catch (InterruptedException e) {
			
			e.printStackTrace();
		}
		
		indate.click();

	}
	//method to click on check out button
	public void checkOutButton() {
		checkout_button.click();
	}
	
	//method to select the date

	public void checkOutDate(String checkOutDate) {
		WebElement outdate=driver.findElement(By.xpath("//*[contains(@class,'DayPicker-Day')][@tabindex=-1][@aria-disabled='false'][contains(@aria-label,'"+checkOutDate+"')]"));
		j.executeScript("arguments[0].click()", outdate);

	}
	
	//method to click on the adult button
	public void adultButton() {
		adult_button.click();
	}
	//mehtod to get the list of adult
	public List<String> adultList(){
		
		List<String> adultList=new ArrayList<String>();
		 for(WebElement adultlist:adult_list) {
			 adultList.add(adultlist.getText());
		    }
		 
		 return adultList;
		
	}

	//method to click the room button 
	public void roomButton() {
		room_button.click();
	}
	

	
	//method to select room member
	public void roomMember(int i){
		try {
			Thread.sleep(2000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		WebElement web=driver.findElement(By.xpath("//*[@id=\"top-banner\"]/div[2]/div/div[1]/div[2]/div/div[1]/div[4]/div[1]/div[1]/div[1]/div[2]/ul/li["+i+"]"));
		WebDriverWait wait=new WebDriverWait(driver,Duration.ofSeconds(10));
		wait.until(ExpectedConditions.visibilityOf(web));
		web.click();
		
	}
	
	//method to get the price
	public void pricePerNightButton() {
		
		j.executeScript("arguments[0].click()", price_per_night_button);
	}
	
	
	//method for rate selection
	public void rate_selection() {
		rate_first_low.click();
		
	}
	
	//method to get the result
	public String result() {
		try {
			Thread.sleep(3000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return result_text.getText();
	}
	
	
}
