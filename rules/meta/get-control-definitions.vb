'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Summary: Outputs all control definitions to a text file in the %Temp% folder

Private Sub Main
	Dim outputFolder As String = System.IO.Path.GetTempPath()
	Dim outputFilename As String = "ControlDefinitions-" & ThisApplication.SoftwareVersion.DisplayVersion & ".txt"
	Const delimiter As String = "|"
	Dim IncludeMacros As Boolean = True 'Macros are public VBA subs.
	
	Dim lines As New List(Of String)
	lines.Add(Join({"InternalName","ClientID","DescriptionText","DisplayName"},delimiter))
	
	For Each ctrlDef As ControlDefinition In ThisApplication.CommandManager.ControlDefinitions
		With ctrlDef
			If Not IncludeMacros AndAlso Left(.InternalName, 6) = "macro:" Then Continue For
			lines.Add(Join({.InternalName, .ClientId, CleanStr(.DescriptionText), CleanStr(.DisplayName)}, delimiter))
		End With
	Next
	
	Dim filePath As String = System.IO.Path.Combine(outputFolder,outputFilename)
	System.IO.File.WriteAllLines(filePath,lines)
	Process.Start(filePath) 'Opens the file in its default application
End Sub

Private Function CleanStr(str As String) As String
	'Remove preceding newline
	Select Case Left(str, 1)
	Case vbCr, vbLf, vbCrLf
		str = Mid(str,2)
	End Select
	
	'Remove trailing newline
	Select Case Right(str, 1)
	Case vbCr, vbLf, vbCrLf
		str = Left(str,Len(str)-1)
	End Select
	
	'Replace any remaining newlines with spaces
	str = Replace(str, vbCr, " ")
	str = Replace(str, vbLf, " ")
	str = Replace(str, vbCrLf, " ")
	
	'Trim preceding & trailing spaces
	str = Trim(str)
	
	Return str
End Function
