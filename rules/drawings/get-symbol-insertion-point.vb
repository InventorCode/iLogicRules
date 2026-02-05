Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument

' Prompt user to select a Sketched Symbol
' kDrawingSketchedSymbolFilter ensures the user can only click symbols
Dim oSelectedEntity As Object = ThisApplication.CommandManager.Pick( _
    SelectionFilterEnum.kDrawingSketchedSymbolFilter, _
    "Select a sketched symbol to find its insertion point")

' Check if the user cancelled the selection (hit Esc)
If (oSelectedEntity Is Nothing) Then
    Return ' Exit rule
End If

' Cast the selected object to a SketchedSymbol
Dim oSymbol As SketchedSymbol = oSelectedEntity

' Get the X and Y coordinates (internal units: cm)
Dim xCoord As Double = oSymbol.Position.X
Dim yCoord As Double = oSymbol.Position.Y

' Display results
MessageBox.Show("Symbol Name: " & oSymbol.Name & vbLf & _
                "Insertion X: " & Round(xCoord, 4) & " cm" & vbLf & _
                "Insertion Y: " & Round(yCoord, 4) & " cm", "Symbol Position")
