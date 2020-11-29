'Source: https://github.com/InventorCode/iLogicRules
'Title: Copy Active File Location
'Author: nannerdw
'Description: Copies the active edit document's file location to the clipboard

Dim doc As Document = ThisApplication.ActiveEditDocument

If doc.FileSaveCounter = 0 Then
	MessageBox.Show("The active edit document must be saved first.")
Else
	My.Computer.Clipboard.SetText(System.IO.Path.GetDirectoryName(doc.FullFileName))
End If