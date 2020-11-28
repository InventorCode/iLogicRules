'Source: https://github.com/InventorCode/iLogicRules
'Title: Rotate Revision Table
'Author: Matthew D. Jordan
'Description: New revision tables in drawing sheets are inserted in a
' horizontal orientation.  This rule rotates them a specified amount.

'!!! Change the revisionTableName variable to match the rev table name in your drawing. !!!

Sub Main()

	' Only run if the document is a drawing
	If ThisDoc.Document.DocumentType <> kDrawingDocumentObject Then
		MsgBox("This File Type is not supported.")
		Exit Sub
	End If

    Dim revisionTableName As String = "REVISION HISTORY"
    Dim rotationAmount As Integer = 90

	Dim drawDoc as DrawingDocument = ThisDoc.Document
    Dim activeSheet As Sheet = drawDoc.ActiveSheet
    Dim revisionTables As RevisionTables = activeSheet.RevisionTables

    revisionTables.Item(revisionTableName).Rotation= rotationAmount * Math.Pi/180

End Sub    
