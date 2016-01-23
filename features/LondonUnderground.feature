Feature: London Underground

Scenario Layout: Correct fare is charged for travel from zone 5 to london bridge
	Given I have touched in with my oystercard at <Entry> station
	And the time is <EntryTime>
	And the barrier displays <OpeningBalance> 
	When I touch out at <Exit> station 
	Then I am charged <Fare> 
	And the barrier displays <ClosingBalance>
	Examples:
		|Entry  |Exit   |Fare  |EntryTime |OpeningBalance |ClosingBalance|
		|zone 5 | Zone 1| 4.70 |07:00     |20.00          | 15.30        |
		|zone 1 | Zone 5| 2.50 |10:00     |15.30          | 11.80        |
		