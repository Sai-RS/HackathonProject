Feature: gift cards

  @smoke
  Scenario: verification of invalid inputs
    Given the user navigates to page
    #Then user close the Ads
    #Then user close the login/signup
    And user click the Cab icon
    And Click on Search button
    When the user clicks on more options
    And user click the gift Cards
    Given the user is on the gift card page and selects Birthday gift
    And the user select the Email button
    And the user enters the valid values (Reciepient name: "Sai",Senders name: "Gopu")
    And the user enters the valid values (Reciepient number : "9876543210", sender number: "6382907857")
    And the user entersb the invalid mail Ids(sender mail id: "sai1452@gmail/com" ,reciever mail id: "gopu2421@gmail.com")
    And the user compose a message ( message: "Wishing you a very Happy birthday")
    Then the user clicks on the Buy now button
		