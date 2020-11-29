'Source: https://github.com/InventorCode/iLogicRules
'Title: Open Active Web Link
'Author: nannerdw
'Description: Opens the active edit document's Web Link iProperty in the default browser

Dim doc As Document = ThisApplication.ActiveEditDocument
Dim webLink As String = doc.PropertySets("Design Tracking Properties").Item("Catalog Web Link").Value
If Not webLink = "" Then
	Process.Start(webLink)
Else
	MessageBox.Show("The active edit document does not contain a web link.", iLogicVb.RuleName)
End If