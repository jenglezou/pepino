Feature: Sample feature to test cucumber4vbs

Scenario: This is the first sample scenario being run
Given a client that has a name of "mr" and "client"
When I run this first scenario
And I look at the "contents" of the output
Then I see that <cucumber4vbs> is working

Scenario: This is the second sample scenario being run
Given I have provide this "value" to the test
When I try to pass values "one" and "two" as parameters
And I also pass the following:
|name|address|phone|
|abc| def | 123456|
|xyzxyz|lmnop|654321|
And I try to check if it really works "well"
Then I know the parser is working
And I know the "values" are read
But it should also have "one" and:
|value|
|xxxxx|

Scenario Outline: This is the third sample scenario being run
Given I have provide this "value" to the test
When I try to pass values <name> and <address> and <phone> as parameters
Then I know the parser is working
Examples:
|name|address|phone|
|abc| def | 123456|
|xyzxyz|lmnop|654321|
