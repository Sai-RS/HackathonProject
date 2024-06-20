Feature: Verify the Search Functionality of Cab

  @smoke
  Scenario: Successful of Search
    Given the user navigates to page
    Then user close the Ads
    Then user close the login/signup
    And user click the Cab icon
    And user click the One-Way cab
    When user enter from "Delhi"
    And user enter to "Manali, Himachal Pradesh"
    And user enter Departure "Jul 21 2024"
    And user enter Pickup Time-Hours "6"
    And user enter Pickup Time-Min "30"
    And user enter Pickup Time-Meridiam "AM"
    Then Click on Apply button 
    And Click on Search button
    Then user select the price lowest to Highest Filter in sort option
    When user Check SUV cab type option in filter
    Then user get the lowest price
    And user check it is lowest price or not
    And user print in the console

    
