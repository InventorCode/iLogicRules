'Source: https://github.com/InventorCode/iLogicRules
'Title: Open iComponent Factory
'Author: nannerdw
'Description: If the document is an iPart or iAssembly member, its factory will be opened.
'If "doc" is not passed in as a rule argument, the calling document will be used instead.

Imports Inventor.ObjectTypeEnum

Dim app As Inventor.Application = ThisApplication

Dim doc As Document = Nothing

If RuleArguments("doc") Is Nothing Then
	doc = ThisDoc.Document
Else
	doc = RuleArguments("doc")
End If

Select Case doc.DocumentType
Case kPartDocumentObject, kAssemblyDocumentObject
Case Else
	Exit Sub
End Select

Dim compDef As ComponentDefinition = doc.ComponentDefinition

Dim factoryDoc As Document = Nothing

If doc.DocumentType = kAssemblyDocumentObject AndAlso compDef.IsIassemblyMember Then
	factoryDoc = compDef.iAssemblyMember.ParentFactory.Parent.Document
	
ElseIf doc.DocumentType = kPartDocumentObject AndAlso compDef.IsIpartMember Then
	factoryDoc = compDef.iPartMember.ParentFactory.Parent
	
Else
	Dim docName As String
	
	If doc.FileSaveCounter = 0 Then
		docName = doc.DisplayName
	Else
		docName = System.IO.Path.GetFileName(doc.FullFileName)
	End If
	
	MessageBox.Show(docName & " is not an iPart or iAssembly member." & vbCrLf)
	
	Exit Sub
End If

app.Documents.Open(factoryDoc.FullFileName)