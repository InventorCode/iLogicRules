'Author: nannerdw
'Description: Copies the active edit document's fullFileName to the clipboard

Dim doc As Document = ThisApplication.ActiveEditDocument

If doc.FileSaveCounter = 0 Then
	MessageBox.Show("The active document must be saved first.")
Else
	My.Computer.Clipboard.SetText(doc.FullFileName)
End If
