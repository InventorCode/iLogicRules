'Set Document BOM Structure to Purchased

If ThisApplication.ActiveDocument.DocumentType = kPartDocumentObject Then
	Dim oDoc As Document = ThisApplication.ActiveDocument
	oDoc.ComponentDefinition.BOMStructure = 51973
End If

If ThisApplication.ActiveDocument.DocumentType = kAssemblyDocumentObject Then
	Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
	oDoc.ComponentDefinition.BOMStructure = 51973
End If