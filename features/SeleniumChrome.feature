Feature: Demonstrate a Chrome test using Selenium

Scenario Layout: Google search
	Given the browser is open 
	When I navigate to "https://www.google.co.uk"
	And I search for <SearchItem>
	Then the browser title contains <SearchItem>
Examples:
	|SearchItem|
	|Eiffel Tower|
	|London|

Scenario: Google search
	Given the browser is open 
	When I navigate to "https://www.google.co.uk"
	And I search for "Ghent"
	Then the browser title contains "Ghent"
	But it does not contain "brussels"
