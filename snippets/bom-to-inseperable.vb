'Set Document BOM Structure to Inseperable

Dim oDoc As Document = ThisApplication.ActiveDocument
oDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kInseparableBOMStructure