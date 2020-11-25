'Set Document BOM Structure to Phantom

Dim oDoc As Document = ThisApplication.ActiveDocument
oDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPhantomBOMStructure