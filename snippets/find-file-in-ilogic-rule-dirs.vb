'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw

Private Function FindFileInIlogicRuleDirectories(fileName As String) As String
'Returns the first matching file found in the iLogic external rule directories
'fileName can contain wilcard characters * and ?
'Returns an empty string if nothing was found
	For Each ruleDir As String In iLogicVb.Automation.FileOptions.ExternalRuleDirectories()
		Dim files As String() = System.IO.Directory.GetFiles(ruleDir, fileName, System.IO.SearchOption.AllDirectories)
		If files.Count > 0 Then Return files(0)
	Next
End Function