'--------------------------------
'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: Converts selected iMate results to ordinary constraints so that the iMates can be reused
'Note: If the selected constraint is part of a composite iMate, the other iMates in the composite will also be converted into ordinary constraints.
'Tested in Inventor 2020.3
'--------------------------------
Option Explicit On
Imports Inventor.ObjectTypeEnum
Imports System.Linq

Private Class RuleMain
	Private Const _transactionName As String = "Convert iMate Results to Constraints"
	Private _iMateConstraints As New List (Of AssemblyConstraint)

	Private Sub Main
		Dim app As Inventor.Application = ThisApplication
		Dim doc As AssemblyDocument = TryCast(ThisDoc.Document, AssemblyDocument)
		If doc Is Nothing Then Exit Sub
		
			For Each obj As Object In doc.SelectSet
				Me.AddToImateConstraintsList(obj)
			Next
			
			If _iMateConstraints.Count = 0 Then
				MessageBox.Show(
					caption:="No iMate Results Found",
					text:="One or more of the following must be" & vbCrLf & "selected before running this rule:" & vbCrLf & vbCrLf & 
					"An iMate result" & vbCrLf & vbCrLf &
					"A composite iMate result" & vbCrLf &
					"(All child iMates will be converted.)" & vbCrLf & vbCrLf &
					"A child node of a composite iMate result" & vbCrLf &
					"(All siblings will be converted, too.)" & vbCrLf & vbCrLf &
					"A component occurrence containing one or more iMate results" & vbCrLf &
					"(All of the component's iMate results will be converted.)"
					)
				Exit Sub
			End If
			
			Dim oTransaction As Transaction
			
			Try
				oTransaction = app.TransactionManager.StartTransaction(doc, _transactionName)
				
				_iMateConstraints.ForEach(Sub(x) CopyImateConstraint(x))
				
				_iMateConstraints.ForEach(Sub(x) DeleteImateConstraint(x))
				
				doc.Update2
				
				oTransaction.End
				
			Catch ex As Exception
				oTransaction.Abort
				Throw
			End Try
	End Sub

	Private Sub AddToImateConstraintsList(obj As Object)
		If TypeOf obj Is iMateResult Then
			For Each oConstraint As AssemblyConstraint In obj.Constraints
				If Not _iMateConstraints.Contains(oConstraint) Then _iMateConstraints.Add(oConstraint)
			Next
		Else If TypeOf obj Is AssemblyConstraint Then
			If obj.iMateResult IsNot Nothing Then
				If obj.iMateResult.ParentComposite IsNot Nothing Then
					Me.AddToImateConstraintsList(obj.iMateResult.ParentComposite)
				Else
					Me.AddToImateConstraintsList(obj.iMateResult)
				End If
			End If
		Else If TypeOf obj Is ComponentOccurrence Then
			For Each oConstraint As AssemblyConstraint In obj.Constraints
				Me.AddToImateConstraintsList(oConstraint)
			Next
		End If
	End Sub
	
	Friend Sub CopyImateConstraint(oConstraint)
		With oConstraint
			Dim oConstraints As AssemblyConstraints = .Parent.Constraints
			
			Select Case .Type
			
			Case kAngleConstraintObject
				oConstraints.AddAngleConstraint(.EntityOne, .EntityTwo, .Angle.Expression, .SolutionType, .ReferenceVectorEntity)
				
			Case kFlushConstraintObject
				oConstraints.AddFlushConstraint(.EntityOne, .EntityTwo, .Offset.Expression)
				
			Case kInsertConstraintObject
				oConstraints.AddInsertConstraint2(.EntityOne, .EntityTwo, .AxesOpposed, .Distance.Expression, .LockRotation)
			
			Case kMateConstraintObject
				oConstraints.AddMateConstraint2(.EntityOne, .EntityTwo, .Offset.Expression, .EntityOneInferredType, .EntityTwoInferredType, .SolutionType)
				
			Case kRotateRotateConstraintObject
				oConstraints.AddRotateRotateConstraint(.EntityOne, .EntityTwo, .Ratio.Expression, .ForwardDirection)
				
			Case kRotateTranslateConstraintObject
				oConstraints.AddRotateTranslateConstraint(.EntityOne, .EntityTwo, .Ratio.Expression, .ForwardDirection)
				
			Case kSymmetryConstraintObject
				oConstraints.AddSymmetryConstraint(.EntityOne, .EntityTwo, .SymmetryPlane, .EntityOneInferredType, .EntityTwoInferredType, .NormalsOpposed)
				
			Case kTangentConstraintObject
				oConstraints.AddTangentConstraint(.EntityOne, .EntityTwo, .InsideTangency, .Offset.Expression)
				
			Case kTransitionalConstraintObject
				oConstraints.AddTransitionalConstraint(.FaceOne, .FaceTwo)
				
			End Select
		End With	
	End Sub
	
	Private Sub DeleteImateConstraint(oConstraint As AssemblyConstraint)
		Try
			oConstraint.iMateResult.ParentComposite.Delete
		Catch ex As Exception
			Try
				oConstraint.Delete
			Catch ex1 As Exception
			End Try
		End Try
	End Sub
End Class
