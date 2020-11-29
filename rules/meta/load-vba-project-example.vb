'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: This example loads a file named TestProject.ivb located in
'the "Default VBA project" folder that is set in Application Options - File

Private Sub Main
	Const ruleName As String = "load-vba-project.vb"
	Const vbaProjName As String = "TestProject"
	
	Dim rulePath As String = FindFileInIlogicRuleDirectories(ruleName)
	If rulePath = "" Then
		Logger.Fatal(ruleName & " could not be found in the iLogic external rule directories.")
		ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
	Else
		Dim map As Inventor.NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
		map.Add("filename", vbaProjName)
		iLogicVb.RunExternalRule(rulePath, map)
	End If
End Sub

Private Function FindFileInIlogicRuleDirectories(fileName As String) As String
'Returns the first matching file found in the iLogic external rule directories
'fileName can contain wilcard characters * and ?
'Returns an empty string if nothing was found
	For Each ruleDir As String In iLogicVb.Automation.FileOptions.ExternalRuleDirectories()
		Dim files As String() = System.IO.Directory.GetFiles(ruleDir, fileName, System.IO.SearchOption.AllDirectories)
		If files.Count > 0 Then Return files(0)
	Next
End Function
