'Source: https://github.com/InventorCode/iLogicRules
'Title: Insert Sheet Format
'Description: Insert an existing drawing sheet format to the current drawing.

'!!! Change the sheetFormatName variable to match the sheet format in your drawing. !!!

Sub Main()

	' Only run if the document is a drawing
	If ThisDoc.Document.DocumentType <> kDrawingDocumentObject Then
		MsgBox("This File Type is not supported.")
		Exit Sub
	End If

	Dim sheetFormatName As String = "Sheet Format Name"
	Dim drawDoc as DrawingDocument = ThisDoc.Document

	Try
		' Get the number of sheets to create from user, convert to integer if able
		Dim numberOfSheets As Integer
			numberOfSheets = CInt(Inputbox("Enter the number of sheets to create...","Sheet creation","1"))

		If numberOfSheets >= 10 Then
			If (MsgBox("Are you sure you want to create " & numberOfSheets & " number of sheets?", vbYesNo)=vbNo) Then Exit Sub
		End If
		
		'Loop i times to create the number of sheets required
		Dim sheets as Sheets = drawDoc.Sheets
		Dim sheet As Sheet
		Dim sheetFormat As SheetFormat = drawDoc.SheetFormats.Item(sheetFormatName)

		For i=1 to numberOfSheets
			sheet=sheets.AddUsingSheetFormat(sheetFormat,,"")
		Next i
	
	Catch
		Messagebox.show("Something went sideways; the sheet could not be created." & vbCrLf & "Sorry for your misfortune.")
	End Try

End Sub

 
