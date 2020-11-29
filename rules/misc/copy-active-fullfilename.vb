'Source: https://github.com/InventorCode/iLogicRules
'Title: Copy Active FullFileName
'Author: nannerdw
'Description: Copies the active edit document's fullFileName to the clipboard

Dim doc As Document = ThisApplication.ActiveEditDocument

If doc.FileSaveCounter = 0 Then
	MessageBox.Show("The active edit document must be saved first.")
Else
	My.Computer.Clipboard.SetText(doc.FullFileName)
End If
