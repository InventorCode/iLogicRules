'Source: https://github.com/InventorCode/iLogicRules
'Title: List iLogic Rules (Recursive)
'Author: Matthew D. Jordan
'Description: Lists the iLogic rules contianed in all referenced files in the current document. Output is written to the iLogic log file.

Sub Main	

	Dim document As Document = ThisDoc.Document
	Select Case document.DocumentType

		Case kPartDocumentObject
			Call ListILogicRules(document)

		Case kDrawingDocumentObject, kAssemblyDocumentObject, kPresentationDocumentObject
			Call ListILogicRules(document)

			For Each d In document.AllReferencedDocuments
				Call ListILogicRules(d)
			Next

		Case Else
			Dim message As String = "This tool is not compatible with this type of file: " & ThisDoc.Document.DocumentType.ToString
			Logger.Error(message)

	End Select
End Sub

Sub ListILogicRules(document As Document) 

	Dim iLogicAuto As Object = iLogicVb.Automation
	Dim rules As List(Of String) = GetRules(document)

	If (rules.Count > 0) Then
		Logger.Info("File: " + fileName)
	End If

	ListRules(rules, document)
End Sub

Function GetRules(document As Document) As List (Of string) 
	Dim iLogicAuto As Object = iLogicVb.Automation
	Dim rules As Object = iLogicAuto.rules(document)
	Dim result As new List (Of string)

	If (rules is Nothing) Then
		Return result
	End If

	For Each rule As iLogicRule In rules
		result.Add(rule.Name)
	Next

	return result
End Function

Sub ListRules(rules As List(Of string), document as Document) 

	If (rules.Count > 0)  Then
		Dim fileName = System.IO.Path.GetFileName(document.FullFileName)
		
		For Each rule In rules
			Logger.Info("  Rule: " + rule)
		Next
	End If
End Sub