Dim oDoc As Inventor.Document
oDoc = ThisDoc.Document

Select Case oDoc.DocumentType

	Case kPartDocumentObject


	Case kAssemblyDocumentObject


	Case kDrawingDocumentObject


    Case kDesignElementDocumentObject, kForeignModelDocumentObject, kNoDocument, kPresentationDocumentObject, kSATFileDocumentObject, kUnknownDocumentObject
        If gDebug = True Then Debug.Print ("Strang DocumentObjectType detected.  Aborting.")
        MsgBox ("This tool is not compatible with this type of file.")
        Exit Sub

End Select