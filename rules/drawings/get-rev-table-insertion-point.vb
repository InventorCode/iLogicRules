Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument

' 1. Prompt user to select a Revision Table
' Filter: kDrawingRevisionTableFilter
Dim oSelectedEntity As Object = ThisApplication.CommandManager.Pick( _
    SelectionFilterEnum.kDrawingRevisionTableFilter, _
    "Select the Revision Table to find its insertion point")

' 2. Check if the user cancelled (Hit Esc)
If (oSelectedEntity Is Nothing) Then
    Return
End If

' 3. Cast the object to RevisionTable
Dim oRevTable As RevisionTable = oSelectedEntity

' 4. Get the Position (Point2d)
' Note: For tables, this is usually the corner point defined in the Style
Dim xCoord As Double = oRevTable.Position.X
Dim yCoord As Double = oRevTable.Position.Y

' 5. Display results (converting from cm to document units for clarity)
Dim oUOM As UnitsOfMeasure = oDoc.UnitsOfMeasure
Dim unitString As String = oUOM.GetStringFromType(oUOM.LengthUnits)

MessageBox.Show("Revision Table Position:" & vbLf & _
                "X: " & oUOM.GetStringFromValue(xCoord, oUOM.LengthUnits) & vbLf & _
                "Y: " & oUOM.GetStringFromValue(yCoord, oUOM.LengthUnits), "Table Position")
