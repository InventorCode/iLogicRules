Option Explicit On
Imports System.Linq
Imports Inventor.Curve2dTypeEnum
Imports Inventor.ObjectTypeEnum

'#############################
'###   Add Dowel Symbols   ###
'#############################
' Source: https://github.com/InventorCode/iLogicRules
'
' v1.0
'
'Summary:
'	Adds dowel symbols to all pre-selected circular or arc edges.
'
'	These dowel symbols are not true sketch symbols.
'	They are ordinary sketches that update with changes to geometry.
'
'	Dowel symbols can be added to multiple views at once.
'	Incompatible selections will be ignored.

Private Class RuleMain
	#Region "Constants"
		Private Const _transactionName As String = "Add Dowel Symbols"
		Private Const _dowelSketchName As String = "Dowel Symbols"
	#End Region
	
	#Region "Main"
		Private Sub Main
			Dim doc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
			If doc Is Nothing Then
				MessageBox.Show("A drawing document is not active.")
				Exit Sub
			End If
			
			Dim circularCurveSegs =
				From curveSeg In doc.SelectSet.OfType(Of DrawingCurveSegment)
				Where curveSeg.GeometryType = kCircleCurve2d Or curveSeg.GeometryType = kCircularArcCurve2d
			
			If circularCurveSegs.Count = 0 Then
				MessageBox.Show("No circular or arc edges were selected.")
				Exit Sub
			End If
			
			'Group DrawingCurveSegments by their center points (with positions compared as single-precision numbers),
			Dim curveSegsByCenter =
				From curveSeg As DrawingCurveSegment In circularCurveSegs
				Let cenPt = curveSeg.Geometry.Center
				Let cenPos = New Tuple(Of Single, Single)(CSng(cenPt.X), CSng(cenPt.Y))
				Group curveSeg By Key = cenPos Into Group
				Let maxRad = Group .Max(Function(x) x.Geometry.Radius)
				Select New With {
					Key .cenPos = Key, 
					Key .LargestCurveSegment = Group .First(Function(x) x.Geometry.Radius = maxRad) }
			
			'Keep only the curve segment with the largest radius from each concentric group
			Dim largestCurveSegs As List(Of DrawingCurveSegment) = (
				From concentricGroup In curveSegsByCenter 
				Select concentricGroup.LargestCurveSegment
				).ToList
			
			Dim oTransaction As Transaction
			Try
				oTransaction = ThisApplication.TransactionManager.StartTransaction(doc, _transactionName)
				Me.AddDowelSymbols(largestCurveSegs)
				oTransaction.End
			Catch ex As Exception
				oTransaction.Abort
				Throw
			End Try
		End Sub
	#End Region
	
	#Region "Subs & Functions"
		Private Function GetOrCreateSketchByName(dwgView As DrawingView, sketchName As String) As DrawingSketch
			For Each tempSketch As DrawingSketch In dwgView.Sketches
				If tempSketch.Name = sketchName Then Return tempSketch
			Next tempSketch
			
			'Create sketch if does not exist.
			Dim oSketch As DrawingSketch = dwgView.Sketches.Add()
			oSketch.Name = sketchName
			Return oSketch
		End Function
		
		Private Sub AddDowelSymbols(circularCurveSegs As List(Of DrawingCurveSegment))
			Dim curveSegsByView = circularCurveSegs.GroupBy(Function(x) x.Parent.Parent)
			
			'Loop drawing views
			For Each tmpGroup In curveSegsByView
				Dim tmpView As DrawingView = tmpGroup.Key
				
				Dim tmpSketch As DrawingSketch = Me.GetOrCreateSketchByName(tmpView, _dowelSketchName)
				tmpSketch.Edit()
				
				Try
					tmpSketch.DeferUpdates = True
					
					'Loop circular edges in view
					For Each tmpEdge As DrawingCurveSegment In tmpGroup
						Dim oSketchEntity As SketchEntity = tmpSketch.AddByProjectingEntity(tmpEdge.Parent)
						Dim oSketchCircle As SketchCircle
						
						Dim gc As GeometricConstraints = tmpSketch.GeometricConstraints
						
						If oSketchEntity.Type = kSketchCircleObject Then
							oSketchCircle = oSketchEntity
						Else
							'Add circle and constrain it to the projected arc
							With oSketchEntity.Geometry
								oSketchCircle = tmpSketch.SketchCircles.AddByCenterRadius(.Center, .Radius)
							End With
							
							gc.AddCoincident(oSketchCircle, oSketchEntity.StartSketchPoint)
							gc.AddCoincident(oSketchCircle.CenterSketchPoint, oSketchEntity.CenterSketchPoint)
						End If
						
						'Add sketch points
						Dim tg As TransientGeometry = ThisApplication.TransientGeometry
						Dim pts As SketchPoints = tmpSketch.SketchPoints
						
						Dim centerPoint As SketchPoint = oSketchCircle.CenterSketchPoint
						Dim topPoint As SketchPoint
						Dim bottomPoint As SketchPoint
						Dim leftPoint As SketchPoint
						Dim rightPoint As SketchPoint
						
						With oSketchCircle.Geometry
							topPoint = pts.Add(tg.CreatePoint2d(.Center.X, .Center.Y + .Radius))
							bottomPoint = pts.Add(tg.CreatePoint2d(.Center.X, .Center.Y - .Radius))
							leftPoint = pts.Add(tg.CreatePoint2d(.Center.X - .Radius, .Center.Y))
							rightPoint = pts.Add(tg.CreatePoint2d(.Center.X + .Radius, .Center.Y))
						End With
						
						'Add Lines
						Dim lines As SketchLines = tmpSketch.SketchLines
						Dim topLine As SketchLine = lines.AddByTwoPoints (centerPoint, topPoint)
						Dim bottomLine As SketchLine = lines.AddByTwoPoints (centerPoint, bottomPoint)
						Dim leftLine As SketchLine = lines.AddByTwoPoints (centerPoint, leftPoint)
						Dim rightLine As SketchLine = lines.AddByTwoPoints (centerPoint, rightPoint)
						
						With gc
							'Constrain Points
							.AddCoincident (topPoint, oSketchCircle)
							.AddCoincident (bottomPoint, oSketchCircle)
							.AddCoincident (leftPoint, oSketchCircle)
							.AddCoincident(rightPoint, oSketchCircle)
							
							'Constrain Lines
							.AddCollinear(topLine, bottomLine)
							.AddCollinear(leftLine, rightLine)
							.AddPerpendicular(topLine, rightLine)
							.AddHorizontal(rightLine)
						End With
						
						'Find all profiles in current sketch
						Dim oProfile As Profile = tmpSketch.Profiles.AddForSolid(False)
				
						'Fill profiles in current sketch	
						For Each tmpProfilePath As ProfilePath In oProfile
							Dim oPathSegments As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
							Dim ProfileWillBeFilled As Boolean = False
							
							For Each tmpProfileEntity As ProfileEntity In tmpProfilePath
								oPathSegments.Add(tmpProfileEntity.SketchEntity)
								
								If tmpProfileEntity.CurveType <> kCircularArcCurve2d Then Continue For
									
								With tmpProfileEntity.Curve
									'Quadrants 1 and 3 will be filled.
									'The line connecting the endpoints of the arc spanning one of these quadrants will have a negative slope.
									ProfileWillBeFilled = (.EndPoint.X - .StartPoint.X) * (.EndPoint.Y - .StartPoint.Y) < 0
								End With
							Next tmpProfileEntity
							
							Dim oFillRegions As SketchFillRegions = tmpSketch.SketchFillRegions
							If ProfileWillBeFilled Then oFillRegions.Add (tmpSketch.Profiles.AddForSolid(False, oPathSegments))
						Next tmpProfilePath
					Next tmpEdge
					
				Catch ex As Exception
					tmpSketch.DeferUpdates = False
					Throw
				End Try
				
				tmpSketch.DeferUpdates = False
				tmpSketch.ExitEdit
			Next tmpGroup
		End Sub
	#End Region
End Class
