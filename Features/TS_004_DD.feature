Feature: Data Driven Test

@regression
Scenario Outline: To do data driven test in gift card with valid and invalid emailID.

		Given the user navigates to page
    Then user close the Ads
    Then user close the login/signup
    And user click the Cab icon
    And Click on Search button
    When the user clicks on more options
    And user click the gift Cards
    Given the user is on the gift card page and selects Birthday gift
    And the user select the Email button
    And enter the required details1 "<row_val>"
		Then check whether the emailID is valid for recipient "<row_val>"
		
Examples:

		| row_val |
		| 1				|
		| 2				|
		

@regression
Scenario Outline: To do the data driven test on gift cards with valid and invalid mobile number
		Given the user navigates to page
    Then user close the Ads
    Then user close the login/signup
    And user click the Cab icon
    And Click on Search button
    When the user clicks on more options
    And user click the gift Cards
    Given the user is on the gift card page and selects Birthday gift
    And the user select the Email button
    And enter the required details2 "<row_val>"
    Then check the sender mobile number entered is valid or not "<row_val>"
    
Examples:

		| row_val |
		| 1				|
		| 2				| 