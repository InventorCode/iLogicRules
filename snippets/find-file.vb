'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw

Private Function FindFile(fileName As String, rootFolder As String) As String
'Returns the first matching file in a directory.
'fileName can contain wilcard characters * and ?
'Returns an empty string if nothing was found
	Dim files As String() = System.IO.Directory.GetFiles(rootFolder, fileName, System.IO.SearchOption.AllDirectories)
	If files.Count > 0 Then Return files(0)
End Function