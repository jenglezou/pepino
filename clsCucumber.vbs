Option Explicit

Dim oCucumber : Set oCucumber = New clsCucumber

oCucumber.FeaturesPath = ".\features"
oCucumber.StepsPath = ".\steps"
oCucumber.StepsOutputFile = "GeneratedSteps.vbs"
'oCucumber.RegenerateSpecs = True
oCucumber.ShowDebug = True

oCucumber.Run()

Class clsCucumber
'**********************************************************************************
	Private gsFeaturesPath
	Private gsFeaturesList
	Private gsStepsPath
	Private gsStepsOutputFile
	Private gbRegenerateSpecs			'True = don't run - just regenerate
	
	Private gsGeneratedSteps			'Text of generated step functions (concatenated)
	Private garrStepFunctionSpecs()
	Private garrStepFunctionCalls()
	Private garrStepText()
	Private garrCachedFeatureFile()
	
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
	Public Property Let StepsOutputFile(sFileName)
		gsStepsOutputFile = sFileName
	End Property
	
	Public Property Get StepsOutputFile()
		StepsOutputFile = gsStepsOutputFile
	End Property

	'******************************************************************************
	Private Sub Class_Initialize()		'On Set to New instance
		Dim oFS, sPath
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		sPath = oFS.GetAbsolutePathName(".")
		Set oFS = Nothing

		gsFeaturesPath = sPath & "/features"
		gsStepsPath = sPath & "/steps"
		gsStepsOutputFile = "StepsTemplateFile.vbs"
		gbRegenerateSpecs = False
		gsGeneratedSteps = ""
	End Sub
	
	Private Sub Class_Terminate()		'When Set instance to Nothing
	End Sub
	
	'******************************************************************************
	Public Sub Run()
		LoadSteps()
		ExecuteFeatures()
		
		'If gsGeneratedSteps <> "" Then WriteFile gsStepsPath & "\" & gsStepsOutputFile, gsGeneratedSteps
	End Sub

	'******************************************************************************
	Private Sub LoadSteps()
		Dim oFS, oFile, oFolder
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		'Check if the folder exists
		If oFS.FolderExists(gsStepsPath) Then
			Set oFolder = oFS.GetFolder(gsStepsPath)
			For Each oFile in oFolder.Files
				FileExecuteGlobal(gsStepsPath & "/" & oFile.Name)
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
		Dim oFS, oFolder, oFile, sFilename
		
		Set oFS = CreateObject("Scripting.FileSystemObject")
		Set oFolder = oFS.GetFolder(gsFeaturesPath)
			
		For Each oFile In oFolder.Files
			If UCase(Right(UCase(oFile.Name), Len(".FEATURE"))) = ".FEATURE" Then
				sFilename = "Feature_" & Capitalise(Left(oFile.Name, Len(oFile.Name) - Len(".Feature"))) & "_GeneratedSteps.vbs"

				Call LoadFeatureFile(oFile.Path, garrCachedFeatureFile)	'Read and cahe the features in the file
				Call GetScenariosAndSteps (garrCachedFeatureFile, garrStepFunctionSpecs, garrStepFunctionCalls, garrStepText) 'Process the features in the file
				MsgBox "Executed " & oFile.Name & vbnewline & "Result = " & ExecuteScenarios(garrStepFunctionSpecs, garrStepFunctionCalls, garrStepText, gbRegenerateSpecs, gsGeneratedSteps)				
				If gsGeneratedSteps <> "" Then WriteFile gsStepsPath & "\" & sFilename, gsGeneratedSteps
			End if		
		Next	
		
		Set oFolder = Nothing
		Set oFS = Nothing
	End Sub
	
	'**********************************************************************************
	Private Sub LoadFeatureFile(sFileName, arrCache)
		Dim oFS, oFile, sFile, sLine, sFeature
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

		'Call ShowArr(arrCache)
		Call ShowDebugMsg(ArrayText(arrCache))
		
		oFile.Close				
		Set oFile = Nothing
		Set oFS = Nothing
	End Sub

	'**********************************************************************************
	Private Sub GetScenariosAndSteps(arrFeatureLines, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
		Dim sLine, sFeature, sKeyword
		Dim sStepType
		Dim iLine, iNumLines, iLastStep
		Dim iStep, sNextLine, sTable
		
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

				Call AddToSteps(sLine & " " & sTable, arrStepFunctionSpecs, arrStepFunctionCalls,arrStepText)
				sTable = ""
			Case "AND", "BUT"
				arrFeatureLines(iLine) = sStepType & " " & sLine	'Change it to a Given, When or Then
				iLine = iLine - 1 									'Jump back a line to reprocess as a Given, When or Then
			Case "SCENARIO:"
				sStepType = "Scenario"
				Call AddToSteps(sLine, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
			Case "BACKGROUND:"
				sStepType = "Background"
			Case "SCENARIO OUTLINE:"
				sStepType = "OutLine"
			Case "EXAMPLES:"
				sStepType = "Examples"
			Case Else
			End Select
			
			iLine = iLine + 1
		Wend

		'Call ShowStepArr(arrStepFunctionSpecs, arrStepText)
		Call ShowDebugMsg(ArrayText(arrStepText))
		Call ShowDebugMsg(ArrayText(arrStepFunctionSpecs))

	End Sub


	'**********************************************************************************
	Private Sub AddToSteps(sStepText, arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText)
	
		Dim iUBound : iUBound = UBound(arrStepText)
		If IsEmpty(arrStepText(iUBound)) Then iUBound = iUBound - 1
		ReDim Preserve arrStepFunctionSpecs(iUBound + 1)
		ReDim Preserve arrStepFunctionCalls(iUBound + 1)
		ReDim Preserve arrStepText(iUBound + 1)
		arrStepFunctionSpecs(iUBound+1) = GenerateStepFunctionCallOrSpec("Spec", Replace(sStepText, ":", ""))
		arrStepFunctionCalls(iUBound+1) = GenerateStepFunctionCallOrSpec("Call", Replace(sStepText, ":", ""))
		arrStepText(iUBound+1) = sStepText
		
	End Sub
	
	'**********************************************************************************
	Private Function ExecuteScenarios(arrStepFunctionSpecs, arrStepFunctionCalls, arrStepText, bRegenerateSpecs, sGeneratedSteps)
		Dim iStep, sStepText
		Dim sGeneratedStep, bRetVal, sGeneratedReturn
		
		bRetVal = False
		
		For iStep = 0 To UBound(arrStepText)
			sGeneratedStep = ""
			sGeneratedReturn = "False"
			If Left(arrStepFunctionSpecs(iStep), 1) = "S" Then sGeneratedReturn = "True"
			
			If Not bRegenerateSpecs Then		
				On Error Resume Next
				Execute "bRetVal = " & arrStepFunctionCalls(iStep)
				If Err.Number = 13 Then
					sGeneratedStep = GenerateStepFunctionCode("Step function not found", arrStepFunctionSpecs(iStep), arrStepText(iStep), sGeneratedReturn)
					ShowDebugMsg sGeneratedStep
				Else 
					'ShowDebugMsg bRetVal
				End If
				On Error Goto 0
			Else
				sGeneratedStep = GenerateStepFunctionCode("Regenerate function spec", arrStepFunctionSpecs(iStep), arrStepText(iStep), sGeneratedReturn)
				ShowDebugMsg sGeneratedStep
			End If
			
			sGeneratedSteps = sGeneratedSteps & sGeneratedStep

		Next

		ExecuteScenarios = bRetVal
	
	End Function

	'**********************************************************************************
	'Generates the step function call
	Private Function GenerateStepFunctionCallOrSpec(sCallOrSpec, sStepText)
		Dim sStepFunctionText, arrStep, sArgs, iStep, iArgCount
	
		sStepFunctionText = sStepText
		iArgCount = 0
		sArgs = ""
		arrStep = Split(sStepFunctionText, """")
		For iStep = 1 To UBound(arrStep) Step 2
			iArgCount = iArgCount+1
			sStepFunctionText = Replace(sStepFunctionText, """" & arrStep(iStep) & """", "Arg" & iArgCount)
			If UCase(Left(sCallOrSpec, 1)) = "S" Then
				sArgs = sArgs & ", Arg" & iArgCount									'Spec
			Else
				sArgs = sArgs & ", " & """" & arrStep(iStep) & """"					'Call
			End If 
		Next 
	
		If sArgs <> "" Then
			sArgs = Right(sArgs, Len(sArgs)-2)										'Function arguments
			sStepFunctionText = Replace(Replace(Trim(sStepFunctionText), " ", "_") & "(" & sArgs & ")", "__", "_")
		Else
			sStepFunctionText = Replace(Trim(sStepFunctionText), " ", "_") & "()"	'No function arguments
		End If
	
		GenerateStepFunctionCallOrSpec = sStepFunctionText
	
	End Function

	'**********************************************************************************
	'Generate the step function code
	Private Function GenerateStepFunctionCode(sReason, sStepFunctionSpec, sStepText, sReturn)

		GenerateStepFunctionCode =  "'***********************************************************************" & vbNewLine & _
					"'Generated on " & Now() & vbNewLine & _
					"'" & sReason & ". Template function generated below ..." & vbNewLine & _
					"'***********************************************************************" & vbNewLine & _
					"Function " & sStepFunctionSpec & vbNewLine & _
					vbTab & "MsgBox ""A generated step function (because " & LCase(sReason) & "):"" & vbNewLine & " & _
					"""- " & Replace(sStepText, """", """""") & """, ,""" & sStepFunctionSpec & """" & vbNewLine & _
					vbTab & "'**** Add code for this step ****" & vbNewLine & _ 
					vbTab & Left(sStepFunctionSpec, InStr(1, sStepFunctionSpec, "(")-1) &  " = " & sReturn  & " 'Set to False for steps when generated." & vbNewLine & _
					"End Function" & vbNewLine & vbNewLine 
									
	End Function
		
'******************************************************************************
End Class


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
Sub WriteFile(sFile, sText)
	Dim oFS, oFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FolderExists(oFS.GetParentFolderName(sFile)) Then
		Set oFile = oFS.OpenTextFile(sFile, ForAppending, True)
		oFile.Write sText
		oFile.Close
	End If
	Set oFS = Nothing
End Sub	

'**********************************************************************************
Function ArrayText(arrOfText)
	Dim i, sTemp1		
	For i = 0 To UBound(arrOfText)
		sTemp1 = sTemp1 & arrOfText(i) & vbNewLine
	Next
	
	ArrayText = sTemp1
End Function
	

'**********************************************************************************
Sub ShowArr(arrOfText)
	Dim i, sTemp1		
	For i = 0 To UBound(arrOfText)
		sTemp1 = sTemp1 & arrOfText(i) & vbNewLine
	Next
	
	ShowDebugMsg sTemp1
End Sub

'**********************************************************************************
Sub ShowStepArr(arrStepFunctionSpecs, arrStepText)
	Dim i, sTemp1
	For i = 0 To UBound(arrStepText)
		sTemp1 = sTemp1 & arrStepText(i) & vbNewLine
		sTemp1 = sTemp1 & arrStepFunctionSpecs(i) & vbNewLine
	Next
	
	ShowDebugMsg sTemp1
End Sub

'**********************************************************************************
Sub ShowDebugMsg(sText)
	If oCucumber.ShowDebug = True Then MsgBox sText
End Sub
