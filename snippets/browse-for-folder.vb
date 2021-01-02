'--------------------------------
'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: The BrowseForFolder function prompts the user to select a single folder, and returns that folder path as a string.
'This function uses the msoFileDialogFolderPicker from Microsoft Office, and requires Excel to be installed.
'This dialog is a more modern alternative to System.Windows.Forms.FolderBrowserDialog, which has limited functionality.
'Tested with Inventor 2020.3 and Excel 2016
'--------------------------------
Option Explicit On

AddReference "office.dll"
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core

'Sub Main is provided as an example for calling BrowseForFolder
Private Sub Main
	Dim selectedFolder As String = BrowseForFolder()
	MessageBox.Show(
		caption := "Folder Selected:", 
		text := If (selectedFolder = "", "No folder was selected", selectedFolder)
		)
End Sub

''' <summary>
''' Prompts the user to select a single folder, and returns that folder path as a string.
''' </summary>
''' 
''' <remarks>
''' This function uses the msoFileDialogFolderPicker from Microsoft Office, and requires Excel to be installed. <br/>
''' This dialog is a more modern alternative to System.Windows.forms.FolderBrowserDialog, which has limited functionality.
''' <br/><br/>
''' This function will execute faster if an instance of Excel is already running and <paramref name="ForceNewXLInstance"/> = False <br/>
''' However, it is more reliable to leave <paramref name="ForceNewXLInstance"/> = True, because if this function tries to grab an instance of Excel 
''' that is currently in the process of closing (which can happen if this function is called again within a few seconds of it ending the first time), it can crash.
''' <br/><br/>
''' Source: <seealso href="https://github.com/InventorCode/iLogicRules"/>
''' </remarks>
''' 
''' <returns>Selected folder path</returns>
''' 
Private Function BrowseForFolder(
	Optional InitialFolder As String = "",
	Optional InitialView As MsoFileDialogView = MsoFileDialogView.msoFileDialogViewList,
	Optional Title As String = "Select a Folder",
	Optional ForceNewXLInstance As Boolean = True
	) As String

	Dim xlApp As Excel.Application = Nothing
	Dim DestroyExcelObjectWhenDone As Boolean = False

	Try
		If ForceNewXLInstance OrElse Process.GetProcessesByName("EXCEL").Count = 0 Then
			'A new temporary instance of Excel will be created and later destroyed.
			xlApp = New Excel.ApplicationClass
			DestroyExcelObjectWhenDone = True
		Else
			'Get an existing instance of Excel; it doesn't matter which one.
			xlApp = TryCast(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
		End If

		With xlApp.FileDialog(MsoFileDialogType.msoFileDialogFolderPicker)
			.Title = Title
			.InitialView = InitialView

			If InitialFolder <> "" AndAlso System.IO.Directory.Exists(InitialFolder) Then
				If Right(InitialFolder, 1) <> "\" Then InitialFolder &= "\"
				.InitialFileName = InitialFolder
			End If

			'This is required so that the folder picker window can be forced into focus.
			'If a temporary Excel instance was created, it will flash briefly before the folder picker is shown.
			xlApp.Visible = True

			'This forces the folder picker into focus.
			'This is sometimes needed to prevent Inventor from locking up while waiting for user input to a hidden and inaccessible folder picker.
			AppActivate(xlApp.Application.Caption)

			'If for some reason the deadlock issue above still occurs, then comment out the following line of code.
			'That will keep the temporary Excel instance visible until the folder picker is closed, 
			'allowing the user to manually restore its window focus from the taskbar if necessary.
			If DestroyExcelObjectWhenDone Then xlApp.Visible = False

			.Show

			Return .SelectedItems(0)
		End With
	Catch ex As Exception
		Throw
	Finally
		If DestroyExcelObjectWhenDone Then
			xlApp.Quit()
			xlApp = Nothing
			GC.Collect()
			GC.WaitForPendingFinalizers()
		End If
	End Try
End Function
