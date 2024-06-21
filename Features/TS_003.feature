Feature: Verify the Search Functionality of Cab

  @smoke
  Scenario: Successful of Search
    Given the user navigates to page
    #Then user close the Ads
    #Then user close the login/signup
    And user click on hotels option.
    When user enter City, Property Name Or Location(value="Manali").
    And user Select a valid date from CHECK-IN(value="Jul 13 2024")
    And CHECK-OUT(value="Jul 15 2024") in calendar.
    Then user Check for the values in ADULTS (Value="5") list box.
    And user print the list of Adults Number List
    
