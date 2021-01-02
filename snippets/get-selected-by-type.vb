'Source: https://github.com/InventorCode/iLogicRules
'Title: Get Selected By Type
'Author: nannerdw
'Description: Two functions for returning a list of pre-selected objects matching either:
'ObjectTypeEnum or IEnumerable(Of ObjectTypeEnum)

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

''' <summary>
''' Returns a list of pre-selected objects matching <paramref name="objTypes"/>
''' </summary>
''' 
''' <remarks>
''' Source: <seealso href="https://github.com/InventorCode/iLogicRules"/>
''' </remarks>
''' 
Private Function GetSelectedByType(objTypes As IEnumerable (Of ObjectTypeEnum)) As List(Of Object)
'Returns a list of pre-selected objects matching one of the input types
	Return (
		From x In ThisApplication.ActiveDocument.SelectSet
		Where DirectCast(objTypes, IEnumerable(Of ObjectTypeEnum)).Contains(x.Type)
		).ToList
End Function

''' <summary>
''' Returns a list of pre-selected objects matching <paramref name="objType"/>
''' </summary>
'''
''' <remarks>
''' Source: <seealso href="https://github.com/InventorCode/iLogicRules"/>
''' </remarks>
''' 
Private Function GetSelectedByType(objType As ObjectTypeEnum) As List(Of Object)
'Returns a list of pre-selected objects matching the input type
	Return (
		From x In ThisApplication.ActiveDocument.SelectSet
		Where objType = x.Type
		).ToList
End Function
