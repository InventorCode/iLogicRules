AddReference "Autodesk.iLogic.Core.dll"
AddReference "Autodesk.iLogic.UiBuilderCore.dll"
Imports iLogicCore = Autodesk.iLogic.Core

'Source: https://github.com/InventorCode/iLogicRules
'Title: Delete iLogic Rules and Forms (Recursive)
'Author: Matthew D. Jordan
'Description: Deletes the iLogic rules contianed in all referenced files in the current document. Output is written to the iLogic log file.

Sub Main

	Dim document As Document = ThisDoc.Document
	Select Case document.DocumentType

		Case kPartDocumentObject
			Call Execute(document)

		Case kDrawingDocumentObject, kAssemblyDocumentObject, kPresentationDocumentObject
			Call Execute(document)
			
			For Each d In document.AllReferencedDocuments
				Call Execute(d)
			Next

		Case Else
			Dim message As String = "This tool is not compatible with this type of file: " & ThisDoc.Document.DocumentType.ToString
			Logger.Error(message)

	End Select
	
	RefreshILogicWindow()
	Logger.Info("Delete Rules is completed")
End Sub

Sub Execute(document As Document) 

	Dim iLogicAuto As Object = iLogicVb.Automation
	Dim fileName = System.IO.Path.GetFileName(document.FullFileName)

	DeleteForms(document)

	Dim rules As List(Of String) = GetRules(document)
	If (rules.Count > 0) Then
		Logger.Info("File: " + fileName)
	End If


	For j As Integer = 0 To 2
		rules = GetRules(document)
		DeleteRules(rules, document)
	Next

End Sub


Function GetRules(document As Document) As List (Of string)

	Dim iLogicAuto As Object = iLogicVb.Automation
	Dim rules As Object = iLogicAuto.rules(document)
	
	Dim result As new List (Of String)

	If (rules is Nothing) Then
		Return result
	End If

	For Each rule As iLogicRule In rules
		result.Add(rule.Name)
	Next

	return result
End Function

Sub DeleteRules(rules As List(Of string), document as Document) 

		Dim iLogicAuto As Object = iLogicVb.Automation

		If (rules.Count < 1) Then
			Exit Sub
		End If
			
		For Each rule As String In rules
			Try
				Logger.Info("  Deleting rule: " + rule)
				iLogicAuto.DeleteRule(document, rule)
			Catch
			End Try
		Next
End Sub

Sub DeleteForms(document As Document)

	Dim fileName = System.IO.Path.GetFileName(document.FullFileName)

	Dim uiAttributes As New iLogicCore.UiBuilderStorage.UiAttributeStorage(document)
	Dim formNames As Object = uiAttributes.FormNames

	If (formNames is Nothing) Then
		Return
	End If
	
	'Actually delete the buggers
	Dim attributeSet As Inventor.AttributeSets
	For Each aSet As Inventor.AttributeSet in document.AttributeSets
		If (aSet.Name Like "iLogicInternalUi*") Then
		Try
			aSet.Delete
			'Logger.Info("Deleted forms in file " + fileName)
		Catch
			'Logger.Info("Error deleting form data from " + fileName)
		End Try
		End If
	Next

End Sub

Sub RefreshILogicWindow()
	'refresh iLogic/form browser window
	For Each dockableWindow As Inventor.DockableWindow In ThisApplication.UserInterfaceManager.DockableWindows
		If dockableWindow.InternalName = "ilogic.treeeditor" Then
			dockableWindow.Visible = False
			dockableWindow.Visible = True
		End If
	Next
End Sub