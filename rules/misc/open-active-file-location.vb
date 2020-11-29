'Source: https://github.com/InventorCode/iLogicRules
'Title: Open Active File Location
'Author: nannerdw
'Description: Opens the active edit document's file location

Dim doc As Document = ThisApplication.ActiveEditDocument

If doc.FileSaveCounter = 0 Then
	MessageBox.Show("The active edit document has not been saved.", iLogicVb.RuleName)
Else
	Process.Start(System.IO.Path.GetDirectoryName(doc.FullFileName))
End If