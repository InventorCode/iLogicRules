'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Title: Unload VBA Project
'Description: Unloads all instances of a VBA project file
'
'The RuleArgument "filename" must be passed to this rule from another script.
'The file path can be absolute, or relative to the "Default VBA project" directory that is set in Application Options - File.
'The filename's .ivb file extension is optional.
'See load-vba-project-example.vb for an example of how to call this rule.

Option Explicit On

Private Sub Main
	ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
	
	Logger.Info("Rule Started: " & iLogicVb.RuleName)
	
	Dim filename As String = RuleArguments("filename")
	If filename = "" Then
		Logger.Fatal("A filename was not passed to this rule" & vbCrLf)
		Exit Sub
	End If
	
	Logger.Info("Filename: " & filename)
	
	'Add file extension if it is missing
	If System.IO.Path.GetExtension(filename) <> ".ivb" Then filename = filename & ".ivb"
	
	If System.IO.Path.IsPathRooted(filename) Then
		If Not System.IO.File.Exists(filename) Then
			Logger.Fatal("File not found" & vbCrLf)
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
			Exit Sub
		End If
	End If
	
	Dim loadedProjInstances As List(Of InventorVBAProject) = GetLoadedVBAProjInstancesByFilename(filename)
	
	Dim numOpenInstances As Integer = loadedProjInstances.Count
	
	Select Case numOpenInstances
	Case 0
		Logger.Info("Project already unloaded" & vbCrLf)
		Exit Sub
	Case 1
		Logger.Info("1 loaded project instance found")
	Case Else
		Logger.Info(numOpenInstances & " loaded project instances found")
	End Select
	
	For Each loadedProjInstance As InventorVBAProject In loadedProjInstances
		Try
			loadedProjInstance.Close
		Catch ex As Exception
		End Try
	Next
	
	numOpenInstances = GetLoadedVBAProjInstancesByFilename(filename).Count
	
	Select Case numOpenInstances
	Case 0
		Logger.Info("Project unloaded successfully" & vbCrLf)
	Case 1
		Logger.Fatal("Project could not be unloaded" & vbCrLf)
	Case Else
		Logger.Fatal(numOpenInstances & " project instances could not be unloaded" & vbCrLf)
	End Select
End Sub

Private Function GetLoadedVBAProjInstancesByFilename(fileName As String) As List(Of InventorVBAProject)
'More than one instance of a VBA project can be open at once
	Dim formatPath = Function(x As String) Replace(LCase(GetUNCPath(x)),"/","\")
	
	Dim projInstances As New List(Of InventorVBAProject)
	
	For Each proj As InventorVBAProject In ThisApplication.VBAProjects
		If proj.ProjectType <> kUserVBAProject Then Continue For
		If formatPath(proj.VBProject.FileName) = formatPath(fileName) Then projInstances.Add(proj)
	Next
	
	Return projInstances
End  Function

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
