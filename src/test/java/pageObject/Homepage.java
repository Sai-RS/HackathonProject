package pageObject;

import org.openqa.selenium.*;
import org.openqa.selenium.support.FindBy;
//import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Homepage extends BasePage{

	
	WebDriverWait wait;
	public Homepage(WebDriver driver) {
		super(driver);
		wait = super.myWait;
		
	}
	JavascriptExecutor j=(JavascriptExecutor)driver;
	
	@FindBy(xpath="//li[@class='menu_Flights']")
	WebElement flight_button;
	
	@FindBy(xpath="//*[@class='close']")
	WebElement ads_close; 
	
	@FindBy(xpath="//*[@id='webklipper-publisher-widget-container-notification-frame']")
	WebElement switch_frame;

	// clicking cab
	@FindBy(xpath="(//span[text()='Cabs'])[1]")
	WebElement cab_button;
	
	//Clicking one Way
	@FindBy(xpath="//*[@id=\"top-banner\"]/div[2]/div/div/div[1]/ul/li[1]/span")
	WebElement one_way;

	//clicking from city
	@FindBy(xpath="//*[@id='fromCity']")
	WebElement from_button;
	
	//From input
	@FindBy(xpath="//input[@placeholder='From']")
	WebElement from_input;
	
	//clicking TO city
	@FindBy(xpath="//*[@id='toCity']")
	WebElement to_button;
	
	//To input
	@FindBy(xpath="//input[@placeholder='To']")
	WebElement to_input;
	
	//First Suggestions city
	@FindBy(xpath="//*[@id='react-autowhatever-1-section-0-item-0']")
	WebElement suggestion_city;
	
	//Apply Button of Date and Time
	@FindBy(xpath="//span[@class='applyBtnText']")
	WebElement apply_button;
	
	// Search Button
	@FindBy(xpath="//a[text()='Search']")
	WebElement search_button;
	
	public void searchButton() {
		//WebElement search = wait.until(ExpectedConditions.elementToBeClickable(search_button));
		search_button.click();
		//search.click();
	}
	
	public void adsClose() throws InterruptedException {
		Thread.sleep(2000);
		driver.switchTo().frame(switch_frame);
		ads_close.click();
		
		//Thread.sleep(2000);
		//driver.findElement(By.xpath("//*[@id=\"SW\"]/div[1]/div[2]/div[2]/div/section/span")).click();
	}
	
	public void loginClose() throws InterruptedException {
		//Thread.sleep(2000);
		driver.findElement(By.xpath("//*[@id='SW']/div[1]/div[2]/div[2]/div/section/span")).click();
		
	}
	
	public WebElement switchFrame() {
		return switch_frame;
	}
	
	public void oneWay() {
		
		
		j.executeScript("arguments[0].click()",one_way );
	}
	
	// Cab Button Method
	public void cabButton()  {

		cab_button.click();
	}
	
	// From Button Method
	public void fromButton()  {
			
		from_button.click();
	}
	
	public void toButton()  {
		
		j.executeScript("arguments[0].click()", to_button);
	}
	
	public String searchFromCity(String fromcity) throws InterruptedException {
		
		try {
			from_input.sendKeys(fromcity);
			Thread.sleep(3000);
			String output=suggestion_city.getText();
			return output;
		}
		
		catch(Exception e){
			return "No Data Found";
		}
	}
	
    public String searchToCity(String tocity) throws InterruptedException {
    	try {
    		to_input.sendKeys(tocity);
    		Thread.sleep(3000);
    		String output=suggestion_city.getText();
    		return output;
    	}
    	catch(Exception e){
    		return "No Data Found";
    	}
	}
	
	
	public void fromCity(String fromcity) throws InterruptedException  {
		
		from_input.sendKeys(fromcity);
		Thread.sleep(5000);
		suggestion_city.click();
	}
	
	public void toCity(String tocity) throws InterruptedException  {
		
		to_input.sendKeys(tocity);
		Thread.sleep(5000);
		suggestion_city.click();
	}

	public void date(String depDate) {
		
		driver.findElement(By.xpath("//*[@id=\"top-banner\"]/div[2]/div/div/div[2]/div[1]/div[3]")).click();
		WebElement date=driver.findElement(By.xpath("//*[contains(@class,'DayPicker-Day')][@tabindex=-1][@aria-disabled='false'][contains(@aria-label,'"+depDate+"')]"));
		
		date.click();
	}
	
	public void time(String depTime) {
		
		driver.findElement(By.xpath("//*[@id=\"top-banner\"]/div[2]/div/div/div[2]/div[1]/div[5]")).click();
		WebElement time=driver.findElement(By.xpath("//li[@class='hrSlotItemParent']/span[contains(text(),'"+depTime+"')]"));
		try {
			Thread.sleep(2000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		time.click();
	}
	
	public void min(String depMin) {
		
		WebElement min=driver.findElement(By.xpath("//li[@class='minSlotItemParent']/span[contains(text(),'"+depMin+"')]"));
		j.executeScript("arguments[0].click()", min);

	}
	
	public void depMeridiem(String depMeridiem) {
	
		WebElement meridiem=driver.findElement(By.xpath("//li[contains(@class,'meridianSlotItemParent')]/span[contains(text(),'"+depMeridiem+"')]"));
		
		j.executeScript("arguments[0].click()", meridiem);
	}
	
	public void applyButton() {
	
		j.executeScript("arguments[0].click()", apply_button);
	}
	
	public void flightButton()  {

		flight_button.click();
	}
}
