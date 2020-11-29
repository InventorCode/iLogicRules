Option Explicit On

'Source: https://github.com/InventorCode/iLogicRules
'Title: Copy Selected FullFileNames
'Author: nannerdw
'Description:
'	Copies the fullFileNames of 1 or more selected component occurrences to the clipboard, 
'	sorted alphabetically, each separated by a newline

Const delimiter As String = vbCrLf

Dim activeDoc As AssemblyDocument = TryCast(ThisDoc.Document, AssemblyDocument)
If activeDoc Is Nothing Then
	MessageBox.Show("An assembly document must be active.", iLogicVb.RuleName)
	Exit Sub
End If

Dim selSet As SelectSet = activeDoc.SelectSet

'Get list of documents from selected component occurrences
Dim selectedDocs As List(Of Document) = (
	From x In selSet.OfType(Of ComponentOccurrence)
	Select DirectCast(x.Definition.Document, Document)
	).Distinct.ToList
	
'Add documents from selected patterns or pattern elements
For Each obj As Object In selSet
	Dim patternElem As OccurrencePatternElement = Nothing
	
	Select Case obj.Type
	Case kRectangularOccurrencePatternObject, _
		 kCircularOccurrencePatternObject, _
		 kFeatureBasedOccurrencePatternObject
		 
		patternElem = obj.OccurrencePatternElements(1)
		
	Case kOccurrencePatternElementObject
		patternElem = obj
		
	End Select
	
	If patternElem Is Nothing Then Continue For
		
	For Each occ As ComponentOccurrence In patternElem.Occurrences
		Dim tmpDoc As Document = occ.Definition.Document
		If Not selectedDocs.Contains(tmpDoc) Then selectedDocs.Add(tmpDoc)
	Next
Next

If selectedDocs.Count = 0 Then
	MessageBox.Show("No component occurrences selected", iLogicVb.RuleName)
	Exit Sub
End If

Dim selectedFileNames As New List(Of String)
Dim unsavedDocDisplayNames As New List(Of String)

For Each doc As Document In selectedDocs
	If doc.FileSaveCounter = 0 Then
		unsavedDocDisplayNames.Add(vbtab & doc.DisplayName)
	Else
		selectedFileNames.Add(doc.FullFileName)
	End If
Next

'Copy to clipboard
If selectedFileNames.Count > 0 Then
	selectedFileNames.Sort()
	My.Computer.Clipboard.SetText(Join(selectedFileNames.ToArray, delimiter))
	Logger.Info(selectedFileNames.Count & " filename(s) copied to clipboard." & If(unsavedDocDisplayNames.Count = 0, vbCrLf,""))
End If

'Display error message for unsaved documents
If unsavedDocDisplayNames.Count > 0 Then
	unsavedDocDisplayNames.Sort()
	ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
	Dim msg As String = unsavedDocDisplayNames.Count & " unsaved file(s) not included:"
	Logger.Warn(msg & vbCrLf & Join(unsavedDocDisplayNames.ToArray, vbCrLf) & vbCrLf)
	MessageBox.Show(msg & vbCrLf & "See iLogic log for details.", iLogicVb.RuleName)
End If
