package pageObject;

import org.openqa.selenium.*;
import org.openqa.selenium.support.FindBy;

public class GiftPage extends BasePage{

	public GiftPage(WebDriver driver) {
		super(driver);
	}

	JavascriptExecutor j= (JavascriptExecutor)driver;
	
	// Wedding Gift Card Element
	@FindBy(xpath=("//div[contains(@class,'card__details')]/h3[text()='Birthday Card']"))
	WebElement birthday_gift_card;
		
	//Method for Gift button 
	public void birthdayGift() {
		j.executeScript("arguments[0].scrollIntoView();",birthday_gift_card);
		j.executeScript("arguments[0].click()", birthday_gift_card);
	}
}
