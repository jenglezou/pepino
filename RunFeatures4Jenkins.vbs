Option Explicit

FileExecuteGlobal(".\clsCucumber.vbs")

Dim oCucumber : Set oCucumber = New clsCucumber

oCucumber.FeaturesPath = ".\features"
oCucumber.FeaturesList = "SeleniumChrome4Jenkins"
oCucumber.StepsPath = ".\steps4jenkins"
'oCucumber.RegenerateSpecs = True
'oCucumber.ShowDebug = True

oCucumber.Run()


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

