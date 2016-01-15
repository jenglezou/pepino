Option Explicit

Dim oDriver

'***********************************************************************
'Generated on 14/01/2016 12:35:39
'Step function not found. Template function generated below ...
'***********************************************************************
Function X_Scenario_Google_search()
	'MsgBox "Edited step function:" & vbNewLine & "Step: X_Scenario: Google search" & vbnewline & "Args: none", ,"X_Scenario_Google_search()"
	'**** Add code for this step ****
	X_Scenario_Google_search = True 'Set to False for steps when generated.
End Function

'***********************************************************************
'Generated on 14/01/2016 12:35:39
'Step function not found. Template function generated below ...
'***********************************************************************
Function Given_the_browser_is_open()
	'**** Add code for this step ****
	Set oDriver = CreateObject("Selenium.ChromeDriver")

	'MsgBox "Edited step function:" & vbNewLine & "Step: Given the browser is open " & vbnewline & "Args: none", ,"Given_the_browser_is_open()"
	Given_the_browser_is_open = True 'Set to False for steps when generated.
End Function

'***********************************************************************
'Generated on 14/01/2016 12:35:39
'Step function not found. Template function generated below ...
'***********************************************************************
Function When_I_navigate_to_Arg1(Arg1)
	'**** Add code for this step ****
    oDriver.Get Arg1	
	
	'MsgBox "Edited step function:" & vbNewLine & "Step: When I navigate to ""https://www.google.co.uk"" " & vbnewline & "Args: Arg1=" & Arg1 & "", ,"When_I_navigate_to_Arg1(Arg1)"
	When_I_navigate_to_Arg1 = True 'Set to False for steps when generated.
End Function

'***********************************************************************
'Generated on 14/01/2016 12:35:39
'Step function not found. Template function generated below ...
'***********************************************************************
Function When_And_I_search_for_Arg1(Arg1)
	'**** Add code for this step ****
	oDriver.FindElementByName("q").SendKeys Arg1 & vbLf
	
	'MsgBox "Edited step function:" & vbNewLine & "Step: When And I search for ""Eiffel Tower"" " & vbnewline & "Args: Arg1=" & Arg1 & "", ,"When_And_I_search_for_Arg1(Arg1)"
	When_And_I_search_for_Arg1 = True 'Set to False for steps when generated.
End Function

'***********************************************************************
'Generated on 14/01/2016 12:35:39
'Step function not found. Template function generated below ...
'***********************************************************************
Function Then_the_browser_title_contains_Arg1(Arg1)
	'**** Add code for this step ****
	'MsgBox "Edited step function:" & vbNewLine & "Step: Then the browser title contains ""Eiffel Tower"" " & vbnewline & "Args: Arg1=" & Arg1 & "", ,"Then_the_browser_title_contains_Arg1(Arg1)"
    msgbox "Title=" & oDriver.Title & vbLF & "Click OK to terminate"
    msgbox "Title=" & oDriver.Title & vbLF & "Click OK to terminate"
	
	Then_the_browser_title_contains_Arg1 = True 'Set to False for steps when generated.
End Function

