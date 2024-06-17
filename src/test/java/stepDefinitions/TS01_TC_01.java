package stepDefinitions;

import org.openqa.selenium.WebDriver;
import baseClass.BaseClass;
import io.cucumber.java.en.*;
import pageObject.Homepage;

public class TS01_TC_01 {
	
	
	WebDriver driver;
	Homepage page1;

	@Given("the user navigates to page")
	public void the_user_navigates_to_page() {
		
		BaseClass.getLogger().info("User Navigation to the page............");
		driver=BaseClass.getDriver();
	    page1=new Homepage(driver);
	}
	
	@Then("user close the Ads")
	public void user_close_the_ads() {
		try {
			BaseClass.getLogger().info("User close the Ads......");
			page1.adsClose();
		}
		catch(Exception e) {
			BaseClass.getLogger().info("No Ads in window so Skipped..");
			//System.out.println(e);
		}
	    
	}
	
	@Then("user close the login\\/signup")
	public void user_close_the_login_signup() {
	    try {
	    	BaseClass.getLogger().info("User close the Login/SignUp......");
	    	page1.loginClose();
	    }
	    catch(Exception e) {
	    	BaseClass.getLogger().info("No Login/SignUp pop up in window so Skipped..");
	    }
	}


	@Then("user click the Cab icon")
	public void user_click_the_cab_icon() {
		
		BaseClass.getLogger().info("User click the Cab Button On Home Page");
		page1.cabButton();
	    
	}

	@Then("user click the One-Way cab")
	public void user_click_the_one_way_cab() {
		
		BaseClass.getLogger().info("User Click on th One Way Cab Button");
	    page1.oneWay();
		
	}

	@When("user enter from {string}")
	public void user_enter_from(String from) throws InterruptedException {
		
		BaseClass.getLogger().info("Cab Page ---> From Button");
	    page1.fromButton();
	    
	    BaseClass.getLogger().info("Cab Page ---> From input");
	    page1.fromCity(from);
	    
	    Thread.sleep(2000);
	}
	
	@When("user enter to {string}")
	public void user_enter_to(String to) throws InterruptedException {
		
		BaseClass.getLogger().info("Cab Page ---> To Button");
	    page1.toButton();
	    
	    BaseClass.getLogger().info("Cab Page ---> To input");
	    page1.toCity(to);
	}

	@When("user enter Departure {string}")
	public void user_enter_departure(String Date) {
		
		BaseClass.getLogger().info("Cab Page ---> User input the Date");
		page1.date(Date);
	    
	}

	@When("user enter Pickup Time-Hours {string}")
	public void user_enter_pickup_time_hours(String hours) {
		
		BaseClass.getLogger().info("Cab Page ---> User input the Hours");
		page1.time(hours);
	    
	}

	@When("user enter Pickup Time-Min {string}")
	public void user_enter_pickup_time_min(String Min) {
		
		BaseClass.getLogger().info("Cab Page ---> User input the Minutes");
		page1.min(Min);
	    
	}

	@When("user enter Pickup Time-Meridiam {string}")
	public void user_enter_pickup_time_meridiam(String meridiem) {
		
		BaseClass.getLogger().info("Cab Page ---> User input the Meridian");
	    page1.depMeridiem(meridiem);
	}
	
	
	@Then("Click on Apply button")
	public void click_on_apply_button() {
		
		BaseClass.getLogger().info("Cab Page ---> User click the apply button");
		page1.applyButton();
	    
	}

	@Then("Click on Search button")
	public void click_on_search_button() {
		
		BaseClass.getLogger().info("Cab Page ---> User click the search Button");
		page1.searchButton();
	    try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
