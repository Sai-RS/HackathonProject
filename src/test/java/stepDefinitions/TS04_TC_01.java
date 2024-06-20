package stepDefinitions;

import java.util.HashMap;
import java.util.List;

import org.junit.Assert;
import org.openqa.selenium.WebDriver;

import baseClass.BaseClass;
import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import pageObject.BirthdayGiftCardPage;
import utilities.DataReader;

public class TS04_TC_01 {
	
	WebDriver driver=BaseClass.getDriver();
	List<HashMap<String,String>> datamap;
	BirthdayGiftCardPage birthdayPage = new BirthdayGiftCardPage(driver);
	
	// Scenario Outline 1
	
	@Given("enter the required details1 {string}")
	public void enter_the_required_details1(String rows) throws InterruptedException {
	    datamap = DataReader.data(System.getProperty("user.dir") + "\\TestData\\ExcelData.xlsx", "Recipient_Email"); 
	    
	    int index=Integer.parseInt(rows)-1;
       String rname= datamap.get(index).get("RECIPIENT_NAME");
       String rmobile= datamap.get(index).get("RECIPIENT_MOBILE");
       String rmail= datamap.get(index).get("RECIPIENT_MAIL");
       String sname= datamap.get(index).get("SENDER_NAME");
       String smobile= datamap.get(index).get("SENDER_MOBILE");
       String smail= datamap.get(index).get("SENDER_MAIL");
       
       birthdayPage.recipientsName(rname);
       birthdayPage.senderName(sname);
       birthdayPage.senderNumber(smobile);
       birthdayPage.recipientsNumber(rmobile);
       birthdayPage.senderMail(smail);
       birthdayPage.recipientsMail(rmail);
       birthdayPage.buyNow();	
	}

	@Then("check whether the emailID is valid for recipient {string}")
	public void check_whether_the_email_id_is_valid_for_recipient(String rows) {
	    datamap = DataReader.data(System.getProperty("user.dir") + "\\TestData\\ExcelData.xlsx", "Recipient_Email");
	    int index=Integer.parseInt(rows)-1;
	    
	    String exp_result = datamap.get(index).get("EXPECTED_RESULT");
	    
	    boolean value = birthdayPage.checkForInvalidEmail();
	    
	    if(value) {
	    	if(exp_result.equalsIgnoreCase("Invalid")) {
	    		Assert.assertTrue(true);
	    	}
	    	else {
	    		Assert.fail();
	    	}
	    }
	    else
		   {
			   if(exp_result.equalsIgnoreCase("Valid")) 
			   {
				   
				   Assert.assertTrue(true);
			   }
			   else 
			   {
				   Assert.fail();
			   }
			   
		   }
	}
	
	// Scenario Outline 2
	
	@Given("enter the required details2 {string}")
	public void enter_the_required_details2(String rows) throws InterruptedException {
		datamap = DataReader.data(System.getProperty("user.dir") + "\\TestData\\ExcelData.xlsx", "Mobile_Number"); 
	    
	    int index=Integer.parseInt(rows)-1;
        String rname= datamap.get(index).get("RECIPIENT_NAME");
        String rmobile= datamap.get(index).get("RECIPIENT_MOBILE");
        String rmail= datamap.get(index).get("RECIPIENT_MAIL");
        String sname= datamap.get(index).get("SENDER_NAME");
        String smobile= datamap.get(index).get("SENDER_MOBILE");
        String smail= datamap.get(index).get("SENDER_MAIL");
       
        birthdayPage.recipientsName(rname);
        birthdayPage.senderName(sname);
        birthdayPage.senderNumber(smobile);
        birthdayPage.recipientsNumber(rmobile);
        birthdayPage.senderMail(smail);
        birthdayPage.recipientsMail(rmail);
        birthdayPage.buyNow();	
	}

	@Then("check the sender mobile number entered is valid or not {string}")
	public void check_the_sender_mobile_number_entered_is_valid_or_not(String rows) {
		datamap = DataReader.data(System.getProperty("user.dir") + "\\TestData\\ExcelData.xlsx", "Mobile_Number");
	    int index=Integer.parseInt(rows)-1;
	    
	    String exp_result = datamap.get(index).get("EXPECTED_RESULT");
	    
	    boolean value = birthdayPage.checkForInvalidMobile();
	    
	    if(value) {
	    	if(exp_result.equalsIgnoreCase("Invalid")) {
	    		Assert.assertTrue(true);
	    	}
	    	else {
	    		Assert.fail();
	    	}
	    }
	    else
		   {
			   if(exp_result.equalsIgnoreCase("Valid")) 
			   {
				   
				   Assert.assertTrue(true);
			   }
			   else 
			   {
				   Assert.fail();
			   }
			   
		   }
	}




}
