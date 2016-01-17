Option Explicit

Dim oCucumber : Set oCucumber = New clsCucumber

oCucumber.FeaturesPath = ".\features"
'oCucumber.FeaturesList = "SeleniumChrome"

oCucumber.StepsPath = ".\steps"
oCucumber.ResultsPath = ".\results"
'oCucumber.RegenerateSpecs = True
'oCucumber.ShowDebug = True

oCucumber.Run()

Class clsCucumber
'**********************************************************************************
	Private gsFeaturesPath
	Private gsFeaturesList
	Private gsStepsPath
	Private gsResultsPath
	Private gbRegenerateSpecs			'True = don't run - just regenerate
	
	Public gsGeneratedSteps			'Text of generated step functions (concatenated)
	Public garrStepFunctionSpecs()
	Public garrStepFunctionCalls()
	Public garrStepText()
	Public garrResults()
	Public garrXUnitResults()
	
	Private garrCachedFeatureFile()
	Private gdicStepFunctionsCreated
	Public giBackgroundBegin
	Public giBackgroundEnd
	
	Private gbShowDebug

	'******************************************************************************
	Public Property Let ShowDebug(bTrueFalse)
		gbShowDebug = bTrueFalse
	End Property
	
	Public Property Get ShowDebug()
		ShowDebug = gbShowDebug
	End Property

	'******************************************************************************
	Public Property Let RegenerateSpecs(bTrueFalse)
		gbRegenerateSpecs = bTrueFalse
	End Property
	
	Public Property Get RegenerateSpecs()
		RegenerateSpecs = gbRegenerateSpecs
	End Property
		
	'******************************************************************************
	Public Property Let FeaturesPath(sPath)
		gsFeaturesPath = sPath
	End Property
	
	Public Property Get FeaturesPath()
		FeaturesPath = gsFeaturesPath
	End Property

	'******************************************************************************
	Public Property Let FeaturesList(sList)
		gsFeaturesList = sList
	End Property
	
	Public Property Get FeaturesList()
		FeaturesList = gsFeaturesList
	End Property

	'******************************************************************************
	Public Property Let StepsPath(sPath)
		gsStepsPath = sPath
	End Property

	Public Property Get StepsPath()
		StepsPath = gsStepsPath
	End Property

	'******************************************************************************
	Public Property Let ResultsPath(sPath)
		gsResultsPath = sPath
	End Property

	Public Property Get ResultsPath()
		ResultsPath = gsResultsPath
	End Property

	'******************************************************************************
	Private Sub Class_Initialize()		'On Set to New instance
		Dim oFS, sPath
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		sPath = oFS.GetAbsolutePathName(".")
		Set oFS = Nothing

		gsFeaturesPath = sPath & "/features"
		gsFeaturesList = ""
		gsStepsPath = sPath & "/steps"
		gsStepsPath = sPath & "/results"
		gbRegenerateSpecs = False
		gsGeneratedSteps = ""
		giBackgroundBegin = -1
		giBackgroundEnd = -1
		
		Set gdicStepFunctionsCreated = CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()		'When Set instance to Nothing
	End Sub
	
	'******************************************************************************
	Public Sub Run()
		LoadExistingSteps()
		ExecuteFeatures()
	End Sub

	'******************************************************************************
	Private Sub LoadExistingSteps()
		Dim oFS, oFile, oFolder
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		'Check if the folder exists
		If oFS.FolderExists(gsStepsPath) Then
			Set oFolder = oFS.GetFolder(gsStepsPath)
			For Each oFile in oFolder.Files
				If UCase(Right(oFile.Name, 4)) = ".VBS" Then FileExecuteGlobal(gsStepsPath & "/" & oFile.Name)
			Next
			Set oFile = Nothing
			Set oFolder = Nothing
		Else
			oFS.CreateFolder(gsStepsPath)
		End If
		
		Set oFS = Nothing
		
	End Sub

	'******************************************************************************
	Private Sub ExecuteFeatures()
		Dim oFS, oFolder, oFile, sFilename, sFeature
		Dim iStartLine, iEndLine
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		Set oFolder = oFS.GetFolder(gsFeaturesPath)
			
		For Each oFile In oFolder.Files
			ReDim garrResults(0)
			ReDim garrXUnitResults(0)
			If UCase(Right(UCase(oFile.Name), Len(".FEATURE"))) = ".FEATURE"  Then
				If (gsFeaturesList = "") Or (Left(UCase(oFile.Name), Len(oFile.Name)-Len(".FEATURE")) = UCase(Trim(gsFeaturesList))) Then 
					giBackgroundBegin = -1
					giBackgroundEnd = -1
	
					sFilename = "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_GeneratedSteps.vbs"
	
					Call LoadFeatureFile(oFile.Path, garrCachedFeatureFile)	'Read and cache the features in the file
					sFeature = GetScenariosAndSteps(garrCachedFeatureFile, garrStepFunctionSpecs, garrStepFunctionCalls, garrStepText, garrResults, garrXUnitResults) 'Process the features in the file
WriteFile gsStepsPath & "\" & "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_FunctionSpecs.txt" , ArrayText(garrStepFunctionSpecs), 2
WriteFile gsStepsPath & "\" & "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_FunctionCalls.txt" , ArrayText(garrStepFunctionCalls), 2
WriteFile gsStepsPath & "\" & "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_StepText.txt" , ArrayText(garrStepText), 2
	 				iStartLine = 0
					If giBackgroundBegin > -1 Then iStartLine = giBackgroundEnd + 1
					iEndLine = UBound(garrStepText)
					gsGeneratedSteps = ""
					MsgBox "Executed " & oFile.Name & vbnewline & "Result = " & _
						ExecuteScenarios(iStartLine, iEndLine, sFeature, garrStepFunctionSpecs, garrStepFunctionCalls, garrStepText, gbRegenerateSpecs, gsGeneratedSteps, garrResults, garrXUnitResults)', iResult)				

WriteFile gsResultsPath & "\" & "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_Result.txt" , ArrayText(garrResults), 2
WriteFile gsResultsPath & "\" & "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_Result.xml" , ArrayText(garrXUnitResults) & "</testsuite>" & vbNewLine & "</testsuites>", 2
	
					If gsGeneratedSteps <> "" Then 
						WriteFile gsStepsPath & "\" & sFilename, gsGeneratedSteps, 8
					End If
				
				End if		
			End If
		Next	
		
		Set oFolder = Nothing
		Set oFS = Nothing
	End Sub
	
	'**********************************************************************************
	Private Sub LoadFeatureFile(sFileName, arrCache)
		Dim oFS, oFile, sFile, sLine
		Dim sStepType
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		Set oFile = oFS.OpenTextFile(sFileName)
		
		Dim iLineNum : iLineNum = 0
		While Not oFile.AtEndOfStream
			sLine = Replace(Trim(oFile.ReadLine), vbTab, "")
			If sLine <> "" Then
				ReDim Preserve arrCache(iLineNum)
				arrCache(iLineNum) = sLine
				iLineNum = iLineNum + 1
			End If
		Wend

		Call ShowDebugMsg(ArrayText(arrCache))
		
		oFile.Close				
		Set oFile = Nothing
		Set oFS = Nothing
	End Sub

	'**********************************************************************************
	Private Function GetScenariosAndSteps(arrFeatureLines, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText, arrResults, arrXUnitResults)
		Dim sLine, sFeature, sKeyword
		Dim sStepType
		Dim iLine, iNumLines, iLastStep
		Dim iStep, sNextLine, sTable
		Dim iScenarioStartLine			'Line where scenario starts 
		Dim iScenarioNumber				'First scenario is one etc
		Dim iScenarioIterations			'Number of iterations for scenario
		
		iScenarioNumber = 0
		iScenarioStartLine = 0
		iScenarioIterations = 1
		
		iStep = 0
		iLine = 0
		iNumLines = UBound(arrFeatureLines)
		ReDim arrStepFunctionSpecs(0)
		ReDim arrStepText(0)
				
		While iLine <= iNumLines
			sLine = Trim(Replace(arrFeatureLines(iLine), vbTab, " "))
			sKeyword = Split(sLine, " ")(0)
			Select Case UCase(sKeyword)
			Case "FEATURE:"
				sStepType = "Feature"
				sFeature = Replace(Trim(Right(sLine, Len(sLine) - Len("FEATURE:"))), " ", "_")
				arrResults(0) = Now() & " " & "FEAT" & " " & sLine 
				arrXUnitResults(0) = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine & _
								 "<testsuites>" & vbNewLine & _
								 "<testsuite name=""" & "Cucumber4VBSTestSuite" & """ tests=""[TESTS]"" errors=""[ERRORS]"" failures=""[FAILURES]"" skip=""[SKIP]"">"
			Case "SCENARIO:", "SCENARIO", "BACKGROUND:"
				If UCase(sKeyword) = "BACKGROUND:" Then giBackgroundBegin = iLine
				sKeyword = Split(sLine, ":")(0)
				sStepType = Replace(sKeyword, " ", "")
				If iScenarioNumber > 0 Then 
					Call AddToSteps("X_ENDSCENARIO_" & Right("00" & iScenarioNumber, 2) , arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
					If giBackgroundBegin > -1 And giBackgroundEnd = -1 Then giBackgroundEnd = iLine
				End If
				iScenarioNumber = iScenarioNumber + 1
				iScenarioStartLine = AddToSteps("X_BEGIN" & UCase(Left(sKeyword, 8)) &  Right("00" & iScenarioNumber, 2) & "~~""1""~~""|default|~|none|""", arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
				Call AddToSteps("X_" & sLine, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
				If UCase(sKeyword) = "BACKGROUND" Then
					giBackgroundBegin = iScenarioStartLine
				End If 
			Case "GIVEN", "WHEN", "THEN"
				sStepType = Capitalise(sKeyword)
				
				'Is there a table of data?
				If Right(sLine, 1) = ":" Then						'Expect table data to follow
					'sLine = Left(sLine, Len(sLine) - 1) 			'Get rid of the ":"
					sTable = """"
					sNextLine = Trim(Replace(arrFeatureLines(iLine + 1), vbTab, " ")) 'Next line
					Do While Left(sNextLine, 1) = "|"
						sTable = sTable & sNextLine & "~"
						iLine = iLine + 1
						If iLine = iNumLines Then Exit Do 			'In case it's the last row
						sNextLine = Trim(Replace(arrFeatureLines(iLine + 1), vbTab, " "))
					Loop
					iLine = iLine - 1								'Finished the table so jump back a line
					sTable = Left(sTable, Len(sTable) - 1) & """"	'Remove the end comma and add a double quote
				End If

				Call AddToSteps(sLine & " " & sTable, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
				sTable = ""
			Case "AND", "BUT"
				arrFeatureLines(iLine) = sStepType & " " & sLine	'Change it to a Given, When or Then
				iLine = iLine - 1 									'Jump back a line to reprocess as a Given, When or Then
			'Case "BACKGROUND:"
			'	sStepType = "Background"
			Case "EXAMPLES:"
				sStepType = "Examples"
				'Is there a table of data?
				sTable = """"
				iScenarioIterations = 0
				sNextLine = Trim(Replace(arrFeatureLines(iLine + 1), vbTab, " ")) 'Next line
				Do While Left(sNextLine, 1) = "|"
					sTable = sTable & sNextLine & "~"
					iLine = iLine + 1
					If iLine = iNumLines Then Exit Do 			'In case it's the last row
					iScenarioIterations = iScenarioIterations + 1
					sNextLine = Trim(Replace(arrFeatureLines(iLine + 1), vbTab, " "))
				Loop
				iLine = iLine - 1								'Finished the table so jump back a line
				sTable = Left(sTable, Len(sTable) - 1) & """"	'Remove the end comma and add a double quote

				arrStepFunctionSpecs(iScenarioStartLine) = GenerateStepFunctionCallOrSpec("Spec", "X_BEGINSCENARIO" &  Right("00" & iScenarioNumber, 2) & " """ & iScenarioIterations & """ " & sTable)
				arrStepFunctionCalls(iScenarioStartLine) = GenerateStepFunctionCallOrSpec("Call", "X_BEGINSCENARIO" &  Right("00" & iScenarioNumber, 2) & " """ & iScenarioIterations & """ " & sTable)
				arrStepText(iScenarioStartLine) = "X_BEGINSCENARIO" &  Right("00" & iScenarioNumber, 2) & "~~""" & iScenarioIterations & """~~" & sTable

				Call AddToSteps("X_" & sLine & " " & sTable, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
				sTable = ""
			Case Else

			End Select
			
			iLine = iLine + 1
		Wend

		Call AddToSteps("X_ENDSCENARIO_" & Right("00" & iScenarioNumber, 2) , arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)

		Call ShowDebugMsg(ArrayText(arrStepText))
		Call ShowDebugMsg(ArrayText(arrStepFunctionSpecs))
		Call ShowDebugMsg(ArrayText(arrStepFunctionCalls))
		
		GetScenariosAndSteps = sFeature
	End Function


	'**********************************************************************************
	Private Function AddToSteps(sStepText, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
	
		Dim iUBound : iUBound = UBound(arrStepText)
		If IsEmpty(arrStepText(iUBound)) Then iUBound = iUBound - 1
		ReDim Preserve arrStepFunctionSpecs(iUBound + 1)
		ReDim Preserve arrStepFunctionCalls(iUBound + 1)
		ReDim Preserve arrStepText(iUBound + 1)
		arrStepFunctionSpecs(iUBound+1) = GenerateStepFunctionCallOrSpec("Spec", sStepText)
		arrStepFunctionCalls(iUBound+1) = GenerateStepFunctionCallOrSpec("Call", sStepText)
		arrStepText(iUBound+1) = sStepText
		
		AddToSteps = iUBound+1
	End Function
	
	'**********************************************************************************
	Private Function ExecuteScenarios(ByVal iStartLine, ByVal iEndLine, ByVal sFeature, ByVal arrStepFunctionSpecs, ByVal arrStepFunctionCalls, ByVal arrStepText, bRegenerateSpecs, sGeneratedSteps, ByRef arrResults, arrXUnitResults)', iResult)
		Dim iStep, sStepText
		Dim sGeneratedStep, bRetVal, sGeneratedReturn
		Dim iIter, iIters, iBeginLine 
		Dim sTableData, dicTableData, arrRows
		Dim iCell, sCell, sRow, arrCellNames, arrCellValues
		Dim arrTableData, sStepResult, iResult, sScenario
Dim sTempData
		
		'iResult = 0
		For iStep = iStartLine To iEndLine
		
			Select Case Left(arrStepText(iStep), 7)
			Case "X_BEGIN" 
				sScenario = arrStepText(iStep)
				If giBackgroundBegin > -1 And iStartLine > giBackGroundBegin Then 
					Call ExecuteScenarios(giBackgroundBegin, giBackgroundEnd, sFeature, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText, bRegenerateSpecs, sGeneratedSteps, arrResults, arrXUnitResults)', iResult)
				End If 

				iIters = CInt(Replace(Split(arrStepText(iStep), "~~")(1), """", ""))'Number of iterations
				sTableData = Replace(Split(arrStepText(iStep), "~~")(2), """", "")	'Table data
				arrRows = Split(sTableData, "~")									'Array of rows of |xx|yy|zz|..|
				arrCellNames = Split(arrRows(0), "|")								'Array of cell names from first (zero) element
				
				'Get the rest of the data into an array of dictionary objects
				ReDim arrTableData(0)			'Reset the array of table data
				For iIter = 1 To iIters
					arrCellValues = Split(arrRows(iIter), "|")
					ReDim Preserve arrTableData(iIter-1)
					Set dicTableData = CreateObject("Scripting.Dictionary")
					For iCell = 1 To UBound(arrCellValues) - 1						'Omit the first and last values
						dicTableData.Add Trim(arrCellNames(iCell)), Trim(arrCellValues(iCell)) 
					Next
					Set arrTableData(iIter-1) = dicTableData
				Next

				iIter = 1									'First iteration (there is always one)
				iBeginLine = iStep							'Remember the beginning of the scenario
				
			Case "X_ENDSC"
				iIter = iIter + 1 								'Next iteration
				If iIter <= iIters Then 
					If giBackgroundBegin > -1 And iStartLine > giBackGroundBegin Then 
						Call ExecuteScenarios(giBackgroundBegin, giBackgroundEnd, sFeature, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText, bRegenerateSpecs, sGeneratedSteps, arrResults, arrXUnitResults)
					End If 
					iStep = iBeginLine	'Iterate the scenario
				End if
				
			Case Else 
				sGeneratedStep = ""
				sGeneratedReturn = "False"
				If Left(arrStepFunctionSpecs(iStep), 1) = "X" Then sGeneratedReturn = "True"
				If Not bRegenerateSpecs Then		
					bRetVal = False : sStepResult = "FAIL"
On Error Resume Next
					Execute "bRetVal = " & ReplaceParameters(arrStepFunctionCalls(iStep), arrTableData(iIter-1))
					If Err.Number = 13 Then
						sStepResult = "NDEF"
						If Not gdicStepFunctionsCreated.Exists(arrStepFunctionSpecs(iStep)) Then 
							gdicStepFunctionsCreated.Add arrStepFunctionSpecs(iStep), "Created"
							sGeneratedStep = GenerateStepFunctionCode("Step function not found", arrStepFunctionSpecs(iStep), arrStepText(iStep), sGeneratedReturn)
							ShowDebugMsg sGeneratedStep
						End if
					Else
						If bRetVal = True Then sStepResult = "PASS"
						'ShowDebugMsg bRetVal
					End If
					
					iResult = UBound(arrResults) + 1
					ReDim Preserve arrResults(iResult)
					ReDim Preserve arrXUnitResults(iResult)
					arrResults(iResult) = Now() & " " & sStepResult & " " & arrStepText(iStep) 
					
					arrXUnitResults(iResult) = "<testcase classname=""" & sFeature & """ name=""" & iResult & "-" & ReplaceChars(arrStepText(iStep), Array("""='", "<=[", ">=]")) & """ time=""0"">" & vbNewLine 
					if bRetVal = False Then arrXUnitResults(iResult) = arrXUnitResults(iResult) & "<error type=""exception"" message=""error message"">"  & "FAIL" & "</error>" & vbNewLine
					arrXUnitResults(iResult) = arrXUnitResults(iResult) & "</testcase>"
On Error Goto 0
				Else
					If Not gdicStepFunctionsCreated.Exists(arrStepFunctionSpecs(iStep)) Then 
						gdicStepFunctionsCreated.Add arrStepFunctionSpecs(iStep), "Created"
						sGeneratedStep = GenerateStepFunctionCode("Regenerate function spec", arrStepFunctionSpecs(iStep), arrStepText(iStep), sGeneratedReturn)
						ShowDebugMsg sGeneratedStep
					End if
				End If
				
				sGeneratedSteps = sGeneratedSteps & sGeneratedStep
			End Select
		Next

		ExecuteScenarios = bRetVal
	
	End Function

	'**********************************************************************************
	'Generates the step function call
	Private Function GenerateStepFunctionCallOrSpec(sCallOrSpec, sStepText)
		Dim sStepFunctionText, arrTokens, iToken, sArgs, iArgCount
	
		sStepFunctionText = sStepText
		iArgCount = 0
		sArgs = ""
		
		'First the arguments in quotes
		arrTokens = Split(sStepFunctionText, """")
		For iToken = 1 To UBound(arrTokens) Step 2
			iArgCount = iArgCount+1
			sStepFunctionText = Replace(sStepFunctionText, """" & arrTokens(iToken) & """", "Arg" & iArgCount)
			If UCase(Left(sCallOrSpec, 1)) = "S" Then
				sArgs = sArgs & ", Arg" & iArgCount									'Spec
			Else
				sArgs = sArgs & ", " & """" & arrTokens(iToken) & """"				'Call
			End If 
		Next 
		
		'Then the parameters in <>
		iArgCount = 0
		arrTokens = Split(Replace(sStepFunctionText, ">", "<"), "<")
		For iToken = 1 To UBound(arrTokens) Step 2
			iArgCount = iArgCount+1
			sStepFunctionText = Replace(sStepFunctionText, "<" & arrTokens(iToken) & ">", arrTokens(iToken))
			If UCase(Left(sCallOrSpec, 1)) = "S" Then
				sArgs = sArgs & ", " & "p" & Capitalise(arrTokens(iToken))			'Spec
			Else
				sArgs = sArgs & ", " & """<" & arrTokens(iToken) & ">"""			'Call
			End If 
		Next 
	
		If sArgs <> "" Then
			sArgs = Right(sArgs, Len(sArgs)-2)										'Function arguments
			sStepFunctionText = Replace(Replace(Trim(Replace(sStepFunctionText, ":", "")), " ", "_") & "(" & sArgs & ")", "__", "_")
		Else
			sStepFunctionText = Replace(Trim(Replace(sStepFunctionText, ":", "")), " ", "_") & "()"	'No function arguments
		End If
	
		GenerateStepFunctionCallOrSpec = sStepFunctionText
	
	End Function

	'**********************************************************************************
	'Generate the step function code
	Private Function GenerateStepFunctionCode(sReason, sStepFunctionSpec, sStepText, sReturn)
		Dim arrArgs, sArg, sArgs
		
		arrArgs = Split(Split(Replace(sStepFunctionSpec, ")", " "), "(")(1),",")
		sArgs = """Args: "
		If arrArgs(0) = " " Then
			sArgs = sArgs & "none"""
		Else 
			For Each sArg In arrArgs
				sArgs = sArgs & Trim(sArg) & "="" & " & Trim(sArg) & " & "", " 
			Next
			sArgs = Left(sArgs, Len(sArgs)-2) & """"
		End If
	
		GenerateStepFunctionCode =  "'***********************************************************************" & vbNewLine & _
					"'Generated on " & Now() & vbNewLine & _
					"'" & sReason & ". Template function generated below ..." & vbNewLine & _
					"'***********************************************************************" & vbNewLine & _
					"Function " & sStepFunctionSpec & vbNewLine & _
					vbTab & "MsgBox ""A generated step function (because " & LCase(sReason) & "):"" & vbNewLine & " & _
					"""Step: " & Replace(sStepText, """", """""") & """ & vbnewline & " & sArgs & ", ,""" & sStepFunctionSpec & """" & vbNewLine & _
					vbTab & "'**** Add code for this step ****" & vbNewLine & _ 
					vbTab & Left(sStepFunctionSpec, InStr(1, sStepFunctionSpec, "(")-1) &  " = " & sReturn  & " 'Set to False for steps when generated." & vbNewLine & _
					"End Function" & vbNewLine & vbNewLine 
									
	End Function
		
'******************************************************************************
End Class

'******************************************************************************
' arrReplacements is an array of name=value pairs. eg Array("""='", "<=[", ">=]")
'******************************************************************************
Function ReplaceChars(ByVal sString, ByVal arrReplacements)
	Dim sRep
	Dim sRetVal
	
	sRetVal = sString
	
	For Each sRep In arrReplacements
		sRetVal = Replace(sRetVal, Split(sRep , "=")(0), Split(sRep, "=")(1))
	Next
	
	ReplaceChars = sRetVal
	
End Function
'******************************************************************************
'******************************************************************************
Sub FileExecuteGlobal(sFile)
	Dim oFS, oFile, sText
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	sText = ""
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FileExists(sFile) Then
		Set oFile = oFS.OpenTextFile(sFile, ForReading)
		On Error Resume Next
		sText = oFile.ReadAll
		On Error Goto 0
		oFile.Close
		ExecuteGlobal sText
	End If
	Set oFS = Nothing
End Sub

'**********************************************************************************
Function Capitalise(sText)
	Capitalise = UCase(Left(sText,1)) & LCase(Right(sText,Len(sText)-1))
End Function

'******************************************************************************
Sub WriteFile(sFile, sText, iMode)
	Dim oFS, oFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FolderExists(oFS.GetParentFolderName(sFile)) Then
		Set oFile = oFS.OpenTextFile(sFile, iMode, True)
		oFile.Write sText
		oFile.Close
	End If
	Set oFS = Nothing
End Sub	

'**********************************************************************************
Function ArrayText(arrOfText)
	Dim sLine, sLines		
	For Each sLine In arrOfText
		sLines = sLines & sLine & vbNewLine
	Next
	
	ArrayText = sLines
End Function
	
'**********************************************************************************
Function ReplaceParameters(ByVal sText, ByVal dicParams)
	Dim arrKeys, arrItems
	Dim i
	
	arrKeys = dicParams.Keys
	arrItems = dicParams.Items
	For i = 0 To UBound(arrKeys)
		sText = Replace(sText, "<" & arrKeys(i) & ">", arrItems(i))
	Next
	
	ReplaceParameters = sText
End Function

'**********************************************************************************
Function DicToText(ByVal dicOfText)
	Dim sText
	Dim arrKeys, arrItems
	Dim i
	
	arrKeys = dicOfText.Keys
	arrItems = dicOfText.Items
	For i = 0 To UBound(arrKeys)
		sText = sText & arrKeys(i) & "=" & arrItems(i) & ","
	Next
	
	DicToText = sText
End Function

'**********************************************************************************
Sub ShowDebugMsg(sText)
	If oCucumber.ShowDebug = True Then MsgBox sText
End Sub

'**********************************************************************************

Dim sXUnitFileText, sXUnitTestSuiteName, sXUnitTestClassName, sXUnitStepRow, sXUnitStepName

Function XUnitTestIntro(sXUnitFileText, sXUnitTestSuiteName)

	sXUnitTestSuiteName = "AutomatedTestSuite"
	sXUnitFileText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine & _
					 "<testsuites>" & vbNewLine & _
					 "<testsuite name=""" & sXUnitTestSuiteName & """ tests=""[TESTS]"" errors=""[ERRORS]"" failures=""[FAILURES]"" skip=""[SKIP]"">" & vbNewLine
End Function

Function XUnitTestStep(sXUnitFileText, sXUnitTestSuiteName, sXUnitTestClassName, sXUnitStepName)
	sXUnitTestClassName = sXUnitTestSuiteName & "." & oFso.GetBaseName(sWorkbook)
	'***** For xUnit reporting ********
	sXUnitStepName = CStr(oRow.Cells(1, XL_KEYWORD))
	sXUnitStepRow = right("0" & oRow.Row, 2)
	if sXUnitStepName <> "" then
		sXUnitFileText = sXUnitFileText & "<testcase classname=""" & sXUnitTestClassName & """ name=""Step" & sXUnitStepRow & "-" & sXUnitStepName & """ time=""0"">" & vbNewLine 
		if iRetval = XL_DISPATCH_FAIL or iRetval = XL_DISPATCH_FAILCONTINUE then
			sXUnitFileText = sXUnitFileText & "<error type=""exception"" message=""error message"">" & vbNewLine & sLog & vbNewLine & "</error>" & vbNewLine
		end if 
		sXUnitFileText = sXUnitFileText & "</testcase>" & vbNewLine
	End If
	'***** For xUnit reporting ********
End Function

Function XUnitTestClose(sXUnitFileText)	
	sXUnitFileText = sXUnitFileText & _
					"</testsuite>" & vbNewLine & _
					"</testsuites>" 
End Function

