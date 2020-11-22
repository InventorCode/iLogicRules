'Set Document BOM Structure to Normal

If oDoc.DocumentType = kPartDocumentObject Then
	Dim oDoc As Document = ThisApplication.ActiveDocument
	oDoc.ComponentDefinition.BOMStructure = 51970
End If

If ThisApplication.ActiveDocument.DocumentType = kAssemblyDocumentObject Then
	Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
	oDoc.ComponentDefinition.BOMStructure = 51970
End If