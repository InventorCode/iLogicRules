'Source: https://github.com/InventorCode/iLogicRules
'Title: Sort Parts Lists by Stock Number
'Author: Matthew D. Jordan
'Description: This will sort all parts lists in a drawing by the
' Stock Number (or whatever column you choose) and renumber the item numbers.

Sub Main()
	Try
		ThisApplication.ScreenUpdating = False
		If ThisApplication.ActiveDocument.DocumentType <> kDrawingDocumentObject
			MessageBox.Show("This filetype is not supported, please run from a drawing document.")
		Else
			BomSort()
		ThisApplication.ScreenUpdating = True
		End If
	Catch
		ThisApplication.ScreenUpdating = True
	End Try
End Sub


Sub BomSort

    Dim oDoc As Document = ThisDoc.Document
    Dim oDrawDoc As DrawingDocument = ThisApplication.ActiveDocument
    Dim oSheets As Sheets = oDrawDoc.Sheets
    Dim oSheet As Sheet
    Dim oPartsLists As PartsLists
    Dim oPartsList As PartsList

    For Each oSheet In oSheets

    	If oSheet.PartsLists.Count = 0 Then
    		Continue For
    	End If

    	oSheet.Activate

		For Each oPartsList In oSheet.PartsLists

    	    oPartsList.Sort("STOCK NUMBER")
    	    oPartsList.Renumber
    	    oPartsList.SaveItemOverridesToBOM
            
        Next oPartsList

    Next oSheet
End Sub
