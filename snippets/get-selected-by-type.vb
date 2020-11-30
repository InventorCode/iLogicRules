'Source: https://github.com/InventorCode/iLogicRules
'Title: Get Selected By Type
'Author: nannerdw
'Description: Returns a list of pre-selected objects matching one of the input types.
'objTypes can be a single ObjectTypeEnum or an object implementing IEnumerable(Of ObjectTypeEnum)

'Sub Main is provided as an example for calling GetSelectedByType
Private Sub Main
	
	'Calling GetSelectedByType with an ObjectTypeEnum:
	Dim oPoints = GetSelectedByType(kSketchPointObject)
	
	'Calling GetSelectedByType with an array of ObjectTypeEnum:
	Dim oLines = GetSelectedByType({kSketchLineObject, kSketchLine3DObject})

	MessageBox.Show(
		caption :="GetSelectedByType - Example",
		text:= oPoints.Count & " 2D sketch point(s) selected" & vbCrLf & 
				oLines.Count & " sketch lines(s) selected")
End Sub

Private Function GetSelectedByType(objTypes As Object) As List(Of Object)
'Returns a list of pre-selected objects matching one of the input types.
'objTypes can be a single ObjectTypeEnum or an object implementing IEnumerable(Of ObjectTypeEnum)
	Dim selSet As SelectSet = ThisApplication.ActiveDocument.SelectSet
	
	If TypeOf objTypes Is ObjectTypeEnum Then
		Return (From x In selSet Where objTypes = x.Type).ToList
	Else If TypeOf objTypes Is IEnumerable(Of ObjectTypeEnum) Then
		Return (
			From x In selSet
			Where DirectCast(objTypes, IEnumerable(Of ObjectTypeEnum)).Contains(x.Type)
			).ToList
	Else
		Throw New ArgumentException
	End If
End Function
