package pageObject;

import org.openqa.selenium.*;
import org.openqa.selenium.support.FindBy;

public class BirthdayGiftCardPage extends BasePage{

	public BirthdayGiftCardPage(WebDriver driver) {
		super(driver);
		
	}

	JavascriptExecutor j= (JavascriptExecutor)driver;
	
	//Email Button
	@FindBy(xpath="//div[@class='deliver__wrap']/ul/li[2]")
	WebElement email_button;
	
	//Recipients Name Element
	@FindBy(xpath="//input[@name='name']")
	WebElement recipients_name;
	
	//Recipients Mobile Number Element
	@FindBy(xpath="//input[@name='mobileNo']")
	WebElement recipients_mobile_number;

	//Recipients mail Element
	@FindBy(xpath="//input[@name='emailId']")
	WebElement recipients_mail;
	
	//Sender Name Element
	@FindBy(xpath="//input[@name='senderName']")
	WebElement sender_name;

	//Recipients Mobile Number Element
	@FindBy(xpath="//input[@name='senderMobileNo']")
	WebElement sender_mobile_number;
	
	//Recipients mail Element
	@FindBy(xpath="//input[@name='senderEmailId']")
	WebElement sender_mail;
	
	//Buy Now Button Element
	@FindBy(xpath="//button[contains(text(),'BUY NOW')]")
	WebElement buy_now;

	//Error Mail Message Element
	@FindBy(xpath="//div[@class='sender__details__wrap']//p[contains(@class,'red-text')][normalize-space()='Please enter a valid Email id.']")
	WebElement mail_error;
	
	//Home Page Navigation
	@FindBy(xpath="//*[@id=\"top-banner\"]/div[2]/div/div/div/div[1]/div[1]/ul/li[1]/a")
	WebElement home_page;
	
	//INPUT FOR COMPOSE EMAIL
	@FindBy(xpath="//*[@id=\"deliveredSection\"]/div/div/div/div[1]/div/div[4]/textarea")
	WebElement compose_message;
	
	//Recipients Name id
	@FindBy(xpath="//div/label[text()=\"RECIPIENT'S NAME\"]/ancestor::div[1]/following-sibling::p")
	WebElement recip_Name_valid;
	
	//Recipients Email 
	@FindBy(xpath="//div/label[text()=\"RECIPIENT'S E-MAIL ID\"]/ancestor::div[1]/following-sibling::p")
	WebElement email_valid;
	
	//Sender Mobile element
	@FindBy(xpath="//div/label[text()=\"SENDER'S MOBILE NUMBER\"]/ancestor::div[1]/following-sibling::p")
	WebElement send_mob_valid;
	
	// Final payment option element
	@FindBy(xpath="//span[text()='Payment options']")
	WebElement final_paymentPage;
	
	
	//Check for invalid Recipients name methods
	public boolean checkForInvalidRepName() 
	{	
		try 
		{
		boolean val=recip_Name_valid.isDisplayed();
		return val;
		}
		catch(Exception e) 
		{
		return false;
		}
		
	}
	
	//check for final page method
	public boolean checkForFinalPage() 
	{	
		try 
		{
		boolean val=final_paymentPage.isDisplayed();
		return val;
		}
		catch(Exception e) 
		{
		return false;
		}
		
		
	}
	
	//check for invalid mobile number method
	public boolean checkForInvalidMobile() 
	{
		try 
		{
			boolean val=send_mob_valid.isDisplayed();
			return val;
		}
		catch(Exception e) 
		{
			return false;
		}
		
	}
	
	//check for invalid email method
	
	public boolean checkForInvalidEmail() 
	{
		try 
		{
			boolean val=email_valid.isDisplayed();
			return val;
		}
		catch(Exception e) 
		{
			return false;
		}
		
	}
	
	
	//Method for compose email
	public void composeMessage(String message) {
		compose_message.sendKeys(message);
	}
	
	//Method for email button
	public void emailButton() {
		j.executeScript("arguments[0].scrollIntoView(true);", email_button);
		j.executeScript("arguments[0].click();", email_button);
	}
	
	//method for Recipients name
	public void recipientsName(String receipientName) {
		recipients_name.sendKeys(receipientName);
	}
	
	//Methods for Recipients number
	public void recipientsNumber(String receipientNumber) {
		recipients_mobile_number.sendKeys(receipientNumber);
	}
	
	//Methods for Recipients Mail
	public void recipientsMail(String receipientMail) {
		recipients_mail.sendKeys(receipientMail);
	}
	
	//Methods for sender name
	public void senderName(String senderName) {
		sender_name.sendKeys(senderName);
	}
	
	// Methods for sender number
	public void senderNumber(String senderNumber) {
		sender_mobile_number.sendKeys(senderNumber);
	}

	//methods for sender mail
	public void senderMail(String senderMail) {
		sender_mail.sendKeys(senderMail);
	}
	
	//methods for buy now button
	public void buyNow() throws InterruptedException {
		buy_now.click();
		Thread.sleep(14000);
	}
	
	//mail error method
	public String mailError() {
		String errorMessage=mail_error.getText();
		return errorMessage;
		
	}
	
	//Home page navigation method
	public void Homepagenavigation() {
		j.executeScript("window.scrollTo(0,document.body.scrollHeight()");
		home_page.click();
	}
	
	
}
