'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw

''' <summary>
''' Returns the first matching file in a directory. Filename can contain wilcard characters * and ?
''' </summary>
''' 
''' <param name="Recurse">Search all subfolders</param>
''' 
''' <remarks>
''' Source: <seealso href="https://github.com/InventorCode/iLogicRules"/>
''' </remarks>
''' 
''' <returns>Full file path, or an empty string if nothing was found.</returns>
''' 
Private Function FindFirstFile(Filename As String, RootFolder As String, Optional Recurse As Boolean = True) As String
	Dim files As String() = System.IO.Directory.GetFiles(
		path :=RootFolder, 
		searchPattern :=Filename, 
		searchOption :=If (Recurse, System.IO.SearchOption.AllDirectories, System.IO.SearchOption.TopDirectoryOnly))
	If files.Count > 0 Then Return files(0)
End Function