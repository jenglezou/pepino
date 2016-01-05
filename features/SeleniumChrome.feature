Feature: Demonstrate a Chrome test using Selenium

Scenario: Google search
Given the browser is open 
When I navigate to "https://www.google.co.uk"
And I search for "Eiffel Tower"
Then the browser title contains "Eiffel Tower"

