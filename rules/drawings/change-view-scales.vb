'Source: https://github.com/InventorCode/iLogicRules
'Title: Change View Scales [in Current Drawing Sheet]
'Author: Matthew D. Jordan
'Description: Changes the scales of views in the drawing sheet to that
' specified by the user.  Either edits selected views, or all views if
' none are selected.

Sub Main()
	
	Const transactionName As String = "Change View Scales"

	Dim oSheets As Sheets
	Dim oSheet As Sheet
	Dim oViews As DrawingViews
	Dim oView As DrawingView

	Dim oDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If oDoc Is Nothing Then Exit Sub
		
	oSheet = oDoc.ActiveSheet

	Dim listItems() As String = New String() {
		"1:1",
		"1:2",
		"6"" = 1'-0""",
		"3"" = 1'-0""",
		"1 1/2"" = 1'-0""",
		"1"" = 1'-0""",
		"3/4"" = 1'-0""",
		"1/2"" = 1'-0""",
		"3/8"" = 1'-0""",
		"1/4"" = 1'-0""",
		"1/8"" = 1'-0""",
		"3/32"" = 1'-0""",
		"1/16"" = 1'-0""",
		"1/32"" = 1'-0""",
		"Custom"
		}

	Dim sScale As String = InputListBox(
	    Prompt := "Select the new scale: ",
	    ListItems := listItems, 
	    DefaultValue := listItems(0), 
	    Title := "Scale All Views",
	    ListName := "Scales:",
		Width := 200,
		Height := 400
	)
	
	If sScale = "" Then Exit Sub ' Dialog was closed.  No transaction will be created.

	If sScale = "Custom" Then sScale = InputBox(
		Title :="Enter a custom view scale", 
		Prompt :="Examples of valid scales:" & vbCrLf & 
			"0.5" & vbCrLf & 
			"1/2" & vbCrLf & 
			"1:2" & vbCrLf &
			"6"" = 1'" & vbCrLf &
			"6"" = 1'-0""")

	Dim oTransaction As Transaction
	
	Try
		oTransaction = ThisApplication.TransactionManager.StartTransaction(oDoc, transactionName)
		
		' See if anything is selected
		Dim oSSet As SelectSet = oDoc.SelectSet
		If oSSet.Count = 0 Then

			'if nothing is selected, then run the SetViewName function on all views
			oViews = oSheet.DrawingViews
		    For Each oView In oViews
				SetViewScale(oView, sScale)
			Next
		Else

			'if there are selected objects, iterate through them and run the SetViewName on the view objects
			For Each cView In oSSet.OfType(Of DrawingView)
				SetViewScale(cView, sScale)
			Next 'Each in oSSet

		End If 'oSSet.count = 0
		
		oTransaction.End
		
	Catch ex As Exception
		oTransaction.Abort
		Throw
		
	End Try
End Sub

Function SetViewScale(ByVal xView As DrawingView, ByVal sScale As String)

	Try
		xView.ScaleString = sScale
	Catch
	End Try

End Function
