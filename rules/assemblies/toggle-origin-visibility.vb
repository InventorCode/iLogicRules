'Source: https://github.com/InventorCode/iLogicRules
'Title: Toggle Origin Visibility
'Author: nannerdw
'Description: Toggles the visibility of origin points, axes, and planes for selected components.
'
'Tested in Inventor 2020.3
'
'If one or more of a component's workfeatures are already visible, then all will be toggled off.
'
'While editing in-place from an assembly, components outside of the active edit document can also be selected.
'A component can also be selected by one of its vertices, edges, faces, features, bodies, or workfeatures.
'If nothing is selected, the active edit document's origin workfeatures will be toggled instead.
'
'When a component's workfeature has its visibility changed within the context of an assembly,
'that component's document is modified in memory with the "View rep (dirty)" flag.
'This means that workfeatures in read-only components can't have their visibility toggled.
'See this thread for more info on this limitation:
'https://forums.autodesk.com/t5/inventor-ideas/ability-to-toggle-visibility-of-origin-amp-work-features-on/idc-p/9372835

Option Explicit On
Imports Inventor.ObjectTypeEnum

Private Class RuleMain
	Private Const _transactionName As String = "Toggle Origin Visibility"
	
	Private _app As Inventor.Application = Nothing
	Private _activeDoc As Document = Nothing
	Private _selSet As SelectSet = Nothing
	
	Private Sub Main
		_app = ThisApplication
		
		_activeDoc = _app.ActiveDocument
		Select Case _activeDoc.DocumentType
		Case kPartDocumentObject, kAssemblyDocumentObject
		Case Else
			Exit Sub
		End Select
			
		_selSet = _activeDoc.SelectSet
		
		Dim selectedObjs As List(Of Object) = (_selSet.OfType(Of Object)).ToList
		
		Dim selectedCompDefs As List(Of ComponentDefinition) = (
			From x In Me.GetComponentOccurrences(selectedObjs)
			Where x.Definition.Type <> kWeldsComponentDefinitionObject
			Select DirectCast(x.Definition,ComponentDefinition)
			).Distinct.ToList
		
		Dim oTransaction As Transaction
		Try
			oTransaction = _app.TransactionManager.StartTransaction(_activeDoc, _transactionName)
			
			If selectedObjs.Count = 0 Then
				Dim activeEditDef As ComponentDefinition = _app.ActiveEditDocument.ComponentDefinition
				Me.ToggleVisibility(activeEditDef, Not Me.FoundVisible(activeEditDef))
			Else
				selectedCompDefs.ForEach(Sub(x) Me.ToggleVisibility(x, Not Me.FoundVisible(x)))
			End If
			
			're-select components
			SelectObjects(selectedObjs)
			
			oTransaction.End
			
		Catch ex As Exception
			oTransaction.Abort
			
			're-select components
			SelectObjects(selectedObjs)
			
			Throw
		End Try
	End Sub

	Private Function GetComponentOccurrences(objs As IEnumerable(Of Object)) As List(Of ComponentOccurrence)
		Dim occs As New List(Of ComponentOccurrence)
		
		Dim tryAddToOccs = Sub(occ)
			If Not occs.Contains(occ) Then occs.Add(occ)
		End Sub
		
		For Each obj As Object In objs
			Select Case obj.Type
				
			Case kComponentOccurrenceProxyObject
				tryAddToOccs(obj.NativeObject)
				
			Case kComponentOccurrenceObject
				tryAddToOccs(obj)
				
			Case kRectangularOccurrencePatternObject, _
				 kCircularOccurrencePatternObject, _
				 kFeatureBasedOccurrencePatternObject
				For Each occ As ComponentOccurrence In obj.ParentComponents
					tryAddToOccs(occ)
				Next occ
				
			Case kOccurrencePatternElementObject
				For Each occ As ComponentOccurrence In obj.Occurrences
					Try
						tryAddToOccs(occ)
					Catch ex As MissingMemberException
					End Try
				Next occ
				
			Case Else
				Try
					tryAddToOccs(obj.ContainingOccurrence)
				Catch ex As MissingMemberException
				End Try
				
			End Select
			
		Next obj
		
		Return occs
	End Function

	Private Function FoundVisible(ByRef compDef As ComponentDefinition) As Boolean
		With compDef
			If .WorkPoints.item(1).Visible Then Return True
				
			For i As Integer = 1 To 3
				If .WorkPlanes.item(i).Visible Or .WorkAxes.item(i).Visible Then
					Return True
				End If
			Next
		End With
		
		Return False
	End Function

	Private Sub ToggleVisibility(compDef As ComponentDefinition, newVisState As Boolean)
		Try
			compDef.Workpoints.Item(1).Visible = newVisState
			For i As Integer = 1 To 3
				compDef.WorkPlanes.item(i).Visible = newVisState
				compDef.WorkAxes.item(i).Visible = newVisState
			Next
		Catch ex As Exception
			Select Case ex.HResult
			Case -2147467259 'E_FAIL (Unspecified error)
				'iPart/iAssembly members can only have their origin geometry toggled on/off through the model browser.
				Try
					Me.ToggleVisibilityFromBrowser(compDef.WorkPoints(1), newVisState)
					
					For i As Integer = 1 To 3
						Me.ToggleVisibilityFromBrowser(compDef.WorkPlanes(i),newVisState)
						Me.ToggleVisibilityFromBrowser(compDef.WorkAxes(i),newVisState)
					Next i
				Catch ex1 As Exception
				End Try
			Case Else
			End Select
		End Try
	End Sub
	
	Private Sub ToggleVisibilityFromBrowser(obj As Object, newVisState As Boolean)
		If newVisState <> obj.Visible Then
			_selSet.Clear
			Dim oBrowserNode As BrowserNode = _activeDoc.BrowserPanes("Model").GetBrowserNodeFromObject(obj)
			oBrowserNode.EnsureVisible
			oBrowserNode.DoSelect
			_app.CommandManager.ControlDefinitions.Item("AssemblyVisibilityCtxCmd").Execute
		End If
	End Sub
	
	Private Sub SelectObjects(objs As IEnumerable(Of Object))
		_selSet.Clear
		For Each obj As Object In objs
			Try
				_selSet.Select(obj)
			Catch
			End Try
		Next
	End Sub
End Class
