package stepDefinitions;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.WebDriver;

import baseClass.BaseClass;
import io.cucumber.java.en.*;
import pageObject.HotelBookingPage;

public class TS03_TC_01 {

	WebDriver driver=BaseClass.getDriver();
	HotelBookingPage hotel=new HotelBookingPage(driver);
	Logger logger=LogManager.getLogger();
	
	@Then("user click on hotels option.")
	public void user_click_on_hotels_option() {
		
		BaseClass.getLogger().info("User Click the Hotel Option in Home Page");

		hotel.hotelButton();
		
		
	}

	@When("user enter City, Property Name Or Location\\(value={string}).")
	public void user_enter_city_property_name_or_location_value(String hotelname) throws InterruptedException {
		
		BaseClass.getLogger().info("User click the location button");

		hotel.locationButton();
		
		BaseClass.getLogger().info("User enter the location name");

		hotel.locationName(hotelname);
		
		BaseClass.getLogger().info("User click the Suggestion location");

		hotel.suggestionLocation();
	}

	@When("user Select a valid date from CHECK-IN\\(value={string})")
	public void user_select_a_valid_date_from_check_in_value(String check_in_date) {
		
		BaseClass.getLogger().info("User click the checkIn button");

		hotel.checkInButton();
		
		BaseClass.getLogger().info("User enter the check In Date");

		hotel.checkInDate(check_in_date);
	}

	@When("CHECK-OUT\\(value={string}) in calendar.")
	public void check_out_value_in_calendar(String check_out_date) {
		
		BaseClass.getLogger().info("User click the checkout button");

		hotel.checkOutButton();
		
		BaseClass.getLogger().info("User enter the check out Date");

		hotel.checkOutDate(check_out_date);
	}

	@Then("user Check for the values in ADULTS \\(Value={string}) list box.")
	public void user_check_for_the_values_in_adults_value_list_box(String string) {
		
		BaseClass.getLogger().info("User select the room button");

		hotel.roomGuest();
	}

	@Then("user print the list of Adults Number List")
	public void user_print_the_list_of_adults_number_list() {
		
		BaseClass.getLogger().info("User click the adult button");

		hotel.adultButton();
		List<String> adult_list=hotel.adultList();
		
		BaseClass.getLogger().info("User Print the List Of Adults Number");

		System.out.println("Adult List Numbers: ");
		for(String number:adult_list) {
			System.out.println(number);
		}
		
		
	}

}
