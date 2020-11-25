'Set Document BOM Structure to Normal

Dim oDoc As Document = ThisApplication.ActiveDocument
oDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kNormalBOMStructure