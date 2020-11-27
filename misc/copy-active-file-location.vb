'Author: nannerdw
'Last Modified Date:  27 Nov, 2020
'Description: Copies the active document's file location to the clipboard

Dim doc As Document = ThisApplication.ActiveDocument

If doc.FileSaveCounter = 0 Then
	MessageBox.Show("The active document must be saved first.")
Else
	My.Computer.Clipboard.SetText(System.IO.Path.GetDirectoryName(doc.FullFileName))
End If