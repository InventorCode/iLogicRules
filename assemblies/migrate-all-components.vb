'Source: https://github.com/InventorCode/iLogicRules
'Title: Migrate All Components (to current version)
'Author: Matthew D. Jordan
'Description: This routine will examine all parts in an assembly and determine if they need to be updated to the newest inventor version.
' If so, it will update them all at once.  This is useful when editing the assembly in the Design Assistant, which requires
' all files be up-to-date.

Sub Main()	

	If Not ThisApplication.ActiveDocument.DocumentType = kAssemblyDocumentObject Then
		msgbox("This routine should be run in an assembly file. Exiting...")
     Return
	End If

	Dim assyDoc As AssemblyDocument = ThisApplication.ActiveDocument
		assyDoc.Save

	Dim docFile As Document
	For Each docFile In assyDoc.AllReferencedDocuments

		If docFile.NeedsMigrating = True
		   	ThisApplication.Documents.Open(docFile.FullFileName, False)
			docFile.Save
			docFile.Close
		End If

	Next

	assyDoc.Save

End Sub
