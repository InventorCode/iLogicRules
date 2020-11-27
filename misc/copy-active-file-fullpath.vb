'Author: nannerdw
'Last Modified Date:  27 Nov, 2020
'Description: Copies the active document's fullFileName to the clipboard

Dim doc As Document = ThisApplication.ActiveDocument

If doc.FileSaveCounter = 0 Then
	MessageBox.Show("The active document must be saved first.")
Else
	My.Computer.Clipboard.SetText(doc.FullFileName)
End If