package stepDefinitions;

import java.io.IOException;
import java.time.Duration;
import java.util.Properties;

import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import baseClass.BaseClass;
import io.cucumber.java.After;
import io.cucumber.java.Before;
import io.cucumber.java.Scenario;


public class Hooks {

	 WebDriver driver;
	 Properties p;
	 TakesScreenshot ts;
	 
     
	@Before
    public void setup() throws IOException
    {
    	driver=BaseClass.initilizeBrowser();
    	driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
    	p=BaseClass.getProperties();
    	driver.get(p.getProperty("appURL"));
    	ts=(TakesScreenshot) driver;
    	
    	driver.manage().window().maximize();
	}
		
    
   
    

    @After
    public void addScreenshot(Scenario scenario) {
        
    	// this is for cucumber junit report
        if(scenario.isFailed()) {
        	
        	byte[] screenshot=ts.getScreenshotAs(OutputType.BYTES);
        	scenario.attach(screenshot, "image/png",scenario.getName());
        	            
        }
        else {
        	byte[] screenshot=ts.getScreenshotAs(OutputType.BYTES);
        	scenario.attach(screenshot, "image/png",scenario.getName());
        }
      
    }
    
    @After(order=2)
    public void tearDown(Scenario scenario) {
        		
       driver.quit();
       
    }
   
}
