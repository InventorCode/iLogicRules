'Source: https://github.com/InventorCode/iLogicRules

'All document types in DocumentTypeEnum:
'	kAssemblyDocumentObject
'	kDesignElementDocumentObject
'	kDrawingDocumentObject
'	kForeignModelDocumentObject
'	kNoDocument
'	kPartDocumentObject
'	kPresentationDocumentObject
'	kSATFileDocumentObject
'	kUnknownDocumentObject

Dim oDoc As Inventor.Document = ThisDoc.Document

Select Case oDoc.DocumentType

	Case kPartDocumentObject


	Case kAssemblyDocumentObject


	Case kDrawingDocumentObject
		
		
	Case Else
		Dim message As String = "This tool is not compatible with this type of file: " & ThisDoc.Document.DocumentType.ToString
		
		Logger.Fatal(message)
		
		'Show the iLogic log window
		ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
		
		MessageBox.Show(message, iLogicVb.RuleName)
		
		Exit Sub
	
End Select