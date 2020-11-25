'Simple Example:

Dim doc As Document = ThisDoc.Document
Dim uom As UnitsOfMeasure = doc.UnitsOfMeasure
Dim len As Double = uom.ConvertUnits(5, "mm", "cm")


'Example using lambda expression:

Dim doc As Document = ThisDoc.Document
Dim uom As UnitsOfMeasure = doc.UnitsOfMeasure

Dim cu = Function(num) uom.ConvertUnits(num, "mm", "cm")

Dim len1 As Double = cu(50)
Dim len2 As Double = cu(30)