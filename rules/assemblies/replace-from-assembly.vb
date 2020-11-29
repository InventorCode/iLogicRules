'-------------------------------------------------------------------------------------------------------------------------------------
'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description:  This script prompts the user to select a component from the active assembly to replace one or more selected components.
'Tested in Inventor 2020.3
'-------------------------------------------------------------------------------------------------------------------------------------
Option Explicit On

Imports System.Linq
Imports Inventor.SelectionFilterEnum

'iLogicVb.RuleName can be used in Inventor 2020, but that option is not available in 2018.
Const SCRIPT_NAME As String = "replace-from-assembly" 'This should correlate with the filename of this script.
Const TRANSACTION_NAME As String = "Replace Components" 'This name shows up in the undo stack.

'Make sure the active document is an assembly
Dim doc As Document =  TryCast(ThisDoc.Document,AssemblyDocument)
If doc Is Nothing Then
	MessageBox.Show("An assembly document must be active.", SCRIPT_NAME)
	Exit Sub
End If

Dim app As Inventor.Application = ThisApplication
Dim cmdMgr As CommandManager = app.CommandManager
Dim selSet As SelectSet = doc.SelectSet
Dim repMan As RepresentationsManager = doc.ComponentDefinition.RepresentationsManager
Dim activePosRep As PositionalRepresentation = repMan.ActivePositionalRepresentation 'to be restored at the end of the script

'Get list of all selected objects (to be restored at the end of the script)
Dim selectedObjects As List (Of Object) = (From obj In selSet).ToList

'Get list of selected components
Dim selectedComps As List(Of ComponentOccurrence) = (
	From obj In selectedObjects.OfType(Of ComponentOccurrence)
	).ToList

Dim SuppressError As Boolean = False

Dim oTransaction As Transaction

'Wrap the transaction in a Try/Catch block to ensure the transaction gets ended or aborted.
'If the transaction does not terminate before the end of the script, Inventor can crash.
Try
	'Create transaction
	oTransaction = app.TransactionManager.StartTransaction (doc, TRANSACTION_NAME)
	
	'Prompt user to select a component if none were pre-selected
	If selectedComps.Count = 0 Then
		selectedComps.Add(cmdMgr.Pick(kAssemblyOccurrenceFilter, "Select a component to replace"))
		
		Try
			If selectedComps(0) Is Nothing Then Throw New NullReferenceException
		Catch ex1 As Exception
			'The Pick command was cancelled.  This script will end without displaying any error messages.
			SuppressError = True
			Throw
		End Try
	End If
	
	'De-select all
	selSet.Clear
	
	'Prompt the user to pick a replacement component
	Dim replacementComponent As ComponentOccurrence = cmdMgr.Pick(kAssemblyOccurrenceFilter, "Select replacement component")
	
	If replacementComponent Is Nothing Then
		'The Pick command was cancelled.  This script will end without displaying any error messages.
		SuppressError = True
		Throw New NullReferenceException
	End If
	
	'Make sure the replacement component has been saved.
	Dim replacementDoc As Document = replacementComponent.Definition.Document
	If replacementDoc.FileSaveCounter = 0 Then
		MessageBox.Show(replacementDoc.DisplayName & " must be saved before it can be selected as a replacement.", SCRIPT_NAME)
		SuppressError = True
		Throw New Exception
	End If
	
	'Get filename of replacement component
	Dim replacementFileName As String = replacementComponent.Definition.Document.Fullfilename
	
	'The Master positional representation must be active in order to replace assembly components
	repMan.PositionalRepresentations("Master").Activate
	
	'Replace all pre-selected components with new component
	selectedComps.ForEach(Sub(comp) comp.Replace(replacementFileName, False))
	
	'Return the positional representation to what it was before running this script
	activePosRep.Activate
	
	oTransaction.End
	
Catch ex As Exception
	oTransaction.Abort
	
	'Re-select all components that were selected before running this script
		For Each obj As Object In selectedObjects
			Try
				selSet.Select(obj)
			Catch ex1 As Exception
				'Ignore errors
			End Try
		Next obj
	
	'Display unhandled error message
	If Not SuppressError Then Throw
End Try
