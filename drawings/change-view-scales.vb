'Source: https://github.com/InventorCode/iLogicRules
'Title: Change View Scales [in Current Drawing Sheet]
'Author: Matthew D. Jordan
'Description: Changes the scales of views in the drawing sheet to that
' specified by the user.  Either edits selected views, or all views if
' none are selected.

Sub Main()

	Dim oSheets As Sheets
	Dim oSheet As Sheet
	Dim oViews As DrawingViews
	Dim oView As DrawingView

	oDoc = ThisDoc.Document
	oSheet = oDoc.ActiveSheet

'	sScale = InputBox("Enter the new scale in decimal units:", "Title", "1.0")

Dim listItems() As String = New String() { "1:2",
	"1:1", _
	"6"" = 1'-0"" ", _
	"3"" = 1'-0"" ", _
	"1 1/2"" = 1'-0"" ", _
	"1"" = 1'-0"" ", _
	"3/4"" = 1'-0"" ", _
	"1/2"" = 1'-0"" ", _
	"3/8"" = 1'-0"" ", _
	"1/4"" = 1'-0"" ", _
	"1/8"" = 1'-0"" ", _
	"3/32"" = 1'-0"" ", _
	"1/16"" = 1'-0"" ", _
	"1/32"" = 1'-0"" "}

Dim sScale As String = InputListBox(
    Prompt := "Select the new scale: ",
    ListItems := listItems, 
    DefaultValue := listItems(0), 
    Title := "Scale All Views",
    ListName := "Scales:",
	Width := 200,
	Height := 400
)


	' See if anything is selcted
	Dim oSSet As SelectSet = ThisDoc.Document.SelectSet
	If oSSet.count = 0 Then

		'if nothing is selected, then run the SetViewName function on all views
		oViews = oSheet.DrawingViews
	    For Each oView In oViews
			SetViewScale(oView, sScale)
		Next
	Else

		'if there are selected objects, iterate through them and run the SetViewName on the view objects
		For Each temp In oSSet
			Dim cView As DrawingView= TryCast(temp, DrawingView)
			If cView IsNot Nothing Then SetViewScale(cView, sScale)
		Next 'Each in oSSet

	End If 'oSSet.count = 0
End Sub

Function SetViewScale(ByVal xView As DrawingView, ByVal sScale As String)

	Try
		xView.ScaleString = sScale
	Catch
	End Try

End Function
