Feature: Demonstrate a Chrome test using Selenium

Scenario: Google search
Given the browser is open 
When I navigate to "https://www.google.co.uk"
And I search for "Eiffel Tower"
Then the browser title contains "Eiffel Tower"

Scenario: Google search
Given the browser is open 
When I navigate to "https://www.google.co.uk"
And I search for "Ghent"
Then the browser title contains "Ghent"
But it does not contain "qwerty"
