'Set Document BOM Structure to Purchased

Dim oDoc As Document = ThisApplication.ActiveDocument
oDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure