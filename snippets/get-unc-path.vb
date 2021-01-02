'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw

Private Sub Main
	'This example displays the network path of a folder mapped to Z:
	MessageBox.Show(GetUNCPath("Z:"))
End Sub

''' <summary>
''' Resolves a file or folder path on a mapped network drive to its full UNC path.
''' </summary>
''' 
''' <remarks>
''' This function does not check if strPath is a valid path.
''' <br/><br/>
''' Source: <seealso href="https://github.com/InventorCode/iLogicRules"/>
''' </remarks>
''' 
''' <returns>The original path will be returned if it is already a UNC path.</returns>
''' 
Private Function GetUNCPath(strPath) As String
	Dim fso As Object = CreateObject("Scripting.FileSystemObject")
	Dim strDrive As String = fso.GetDriveName(strPath)

	Try
		Dim strShare As String = fso.Drives(Left(strDrive, Len(strDrive) - 1)).ShareName
		If strShare = "" Then Return strPath
		Return Replace(strPath, strDrive, strShare)
	Catch ex As Exception
		Return strPath
	End Try
End Function