'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Title: Load VBA Project
'Description: Loads a VBA project file
'The RuleArgument "filename" must be passed to this rule from another script.  See the beginning of this script for an example.
'The file path can be absolute, or relative to the "Default VBA project" directory that is set in Application Options - File.
'The filename's .ivb file extension is optional.

Option Explicit On

#Region "Example Calling Rule"
'Copy this code to another iLogic rule, and uncomment it.
'It loads a file named TestProject.ivb located in the "Default VBA project" folder.

'Dim map As Inventor.NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
'map.Add("filename", "TestProject")
'iLogicVb.RunExternalRule("load-vba-project", map)
#End Region

Private Sub Main
	Logger.Info("Rule Started: " & iLogicVb.RuleName)
	
	Dim filename As String = RuleArguments("filename")
	If filename = "" Then
		Logger.Fatal("A filename was not passed to this rule" & vbCrLf)
		LogWindow_Show
		Exit Sub
	End If
	
	Logger.Info("Filename: " & filename)
	
	'Add file extension if it is missing
	If System.IO.Path.GetExtension(filename) <> ".ivb" Then filename = filename & ".ivb"
	
	If System.IO.Path.IsPathRooted(filename) Then
		If Not System.IO.File.Exists(filename) Then
			Logger.Fatal("File not found" & vbCrLf)
			LogWindow_Show
			Exit Sub
		End If
	Else 'Path is relative
		Dim defaultProjFile As String = ThisApplication.FileOptions.DefaultVBAProjectFileFullFilename
		Dim defaultVBAProjFolder As String = System.IO.Path.GetDirectoryName(defaultProjFile)
		filename = System.IO.Path.Combine(defaultVBAProjFolder, filename)
		
		If System.IO.File.Exists(filename) Then
			Logger.Info("File found in " & defaultVBAProjFolder)
		Else
			Logger.Fatal("File not found in " & defaultVBAProjFolder & vbCrLf)
			LogWindow_Show
			Exit Sub
		End If
	End If
	
	'If a project with the same fullFileName is already loaded, do not open it again
	If GetLoadedVBAProjectByFileName(filename) IsNot Nothing Then
		Logger.Info("Project already loaded" & vbCrLf)
		LogWindow_Show
		Exit Sub
	End If
		
	Try
		ThisApplication.VBAProjects.Open(filename)
		Logger.Info("Project loaded successfully" & vbCrLf)
	Catch ex As Exception
		If GetLoadedVBAProjectByFileName(filename) Is Nothing Then
			Logger.Fatal("Project could not be loaded" & vbCrLf)
			LogWindow_Show
			Exit Sub
		End If
		Throw
	End Try
End Sub

Private Function GetLoadedVBAProjectByFileName(fileName As String) As InventorVBAProject
	Dim formatPath = Function(x As String) Replace(LCase(GetUNCPath(x)),"/","\")
	
	For Each proj As InventorVBAProject In ThisApplication.VBAProjects
		If proj.ProjectType <> kUserVBAProject Then Continue For
		If formatPath(proj.VBProject.FileName) = formatPath(fileName) Then Return proj
	Next
End Function

Private Function GetUNCPath(strPath) As String
	Dim fso As Object = CreateObject("Scripting.FileSystemObject")
	Dim strDrive As String = fso.GetDriveName(strPath)
	
	Try
		Dim strShare As String = fso.Drives(Left(strDrive, Len(strDrive) -1)).ShareName
		If strShare = "" Then Return strPath
		Return Replace(strPath, strDrive, strShare)
	Catch ex As Exception
		Return strPath
	End Try
End Function

Private Sub LogWindow_Show()
	ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
End Sub