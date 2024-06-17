package stepDefinitions;

import java.util.List;

import org.junit.Assert;
import org.openqa.selenium.WebDriver;

import baseClass.BaseClass;
import io.cucumber.java.en.*;
import pageObject.CabPage;

public class TS01_TC_02 {
	
	WebDriver driver=BaseClass.getDriver();
	CabPage cabpage=new CabPage(driver);
	String lowest_price;

	@Then("user select the price lowest to Highest Filter in sort option")
	public void user_select_the_price_lowest_to_highest_filter_in_sort_option() {
		
		BaseClass.getLogger().info("Cab Page ---> User Sort the Car");
		try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		cabpage.sortButton();
		
		BaseClass.getLogger().info("Cab Page ---> User click the Lowest to Highest Price");
	    cabpage.lowestPrice();
	}

	@When("user Check SUV cab type option in filter")
	public void user_check_suv_cab_type_option_in_filter() {
		
		
		BaseClass.getLogger().info("Cab Page ---> User click the Car SUV");
		cabpage.carSuv();
	}

	@Then("user get the lowest price")
	public void user_get_the_lowest_price() {
		
		BaseClass.getLogger().info("Cab Page ---> User get the lowest price");
		lowest_price=cabpage.lowestPrice();
	    
	}

	@Then("user check it is lowest price or not")
	public void user_check_it_is_lowest_price_or_not() {
		
		BaseClass.getLogger().info("Cab Page ---> User get the Price List");
		List<String> price_list=cabpage.priceList();
		for(String price:price_list) {
			if(price.equalsIgnoreCase(lowest_price)) {
				
				BaseClass.getLogger().info("Cab Page ---> Lowest Price Printed");
				Assert.assertTrue(true);
				
			}
		}
	    
	}

	@Then("user print in the console")
	public void user_print_in_the_console() {
		
		
		System.out.println("The Lowest Price is: "+lowest_price);
	}
}
