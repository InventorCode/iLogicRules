'Set Document BOM Structure to Reference

Dim oDoc As Document = ThisApplication.ActiveDocument
oDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kReferenceBOMStructure