package pageObject;

import java.time.Duration;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.WebDriverWait;

public class BasePage {
	public WebDriver driver;
	public WebDriverWait myWait;

	// initialize the Driver setup
		public BasePage(WebDriver driver)
		{
			this.driver=driver;
			myWait = new WebDriverWait(driver, Duration.ofMinutes(5));
			
			//Initialize the PageFactory Class
			PageFactory.initElements(driver,this);
		}

}
