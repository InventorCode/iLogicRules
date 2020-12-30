'--------------------------------
'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: Prompts the user to browse for one or more .ipt or .iam files,
'and places them grounded at the origin of the assembly that called this rule.
'
'This script is equivalent to right clicking during the Place command and selecting "Place Grounded at Origin",
'except this script bypasses the component preview that follows the mouse cursor movement, which can be slow for large assemblies.
'--------------------------------
Option Explicit On

Dim transactionName As String = "Place Component"

Dim doc As AssemblyDocument = TryCast(ThisDoc.Document, AssemblyDocument)
If doc Is Nothing Then Exit Sub
	
Dim app As Inventor.Application = doc.Parent

Dim uiMgr As UserInterfaceManager = app.UserInterfaceManager

Dim repMgr As RepresentationsManager = doc.ComponentDefinition.RepresentationsManager

Dim oFileDialog As Inventor.FileDialog
app.CreateFileDialog(oFileDialog)

With oFileDialog
	.DialogTitle = "Place Component(s)"
	.InitialDirectory = app.DesignProjectManager.ActiveDesignProject.WorkspacePath
	.MultiSelectEnabled = True
	.Filter = "Inventor Files (*.iam;*.ipt)|*.iam;*.ipt"
	.CancelError = False
End With

Dim SilentOperation_Backup As Boolean = app.SilentOperation

Dim activePosRep_Backup As PositionalRepresentation = repMgr.ActivePositionalRepresentation

Dim UserInteractionDisabled_Backup As Boolean = uiMgr.UserInteractionDisabled

Dim oTransaction As Transaction = Nothing

Try
	'SilentOperation = False is needed in case Inventor tries to display this prompt:
	'"The location of the selected file is not in the search path of the active project file."
	'Otherwise, the user wouldn't be able to select the file for placement.
	app.SilentOperation = False 

	oFileDialog.ShowOpen

	If oFileDialog.FileName <> "" Then

		Dim arrFileNames() As String = Split(oFileDialog.FileName, "|")
		If arrFileNames.Length > 1 Then transactionName = transactionName & "s"

		oTransaction = app.TransactionManager.StartGlobalTransaction(doc, transactionName)
		
		uiMgr.UserInteractionDisabled = True
		
		'The Master positional representation must be active for components to be placed into the assembly.
		If activePosRep_Backup.Name <> "Master" Then repMgr.PositionalRepresentations("Master").Activate
		
		For Each fileName As String In arrFileNames
			Dim occ As ComponentOccurrence = doc.ComponentDefinition.Occurrences.Add(fileName, app.TransientGeometry.CreateMatrix)
			occ.Grounded = True
		Next fileName
		
		activePosRep_Backup.Activate
	
		uiMgr.UserInteractionDisabled = UserInteractionDisabled_Backup
	End If
	
	app.SilentOperation = SilentOperation_Backup
	
	If oTransaction IsNot Nothing Then oTransaction.End
	
Catch ex As Exception
	activePosRep_Backup.Activate
	
	app.SilentOperation = SilentOperation_Backup
	
	uiMgr.UserInteractionDisabled = UserInteractionDisabled_Backup
	
	If oTransaction IsNot Nothing Then oTransaction.Abort
	
	Throw
End Try
