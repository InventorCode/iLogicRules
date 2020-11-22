Imports System.Text.RegularExpressions

'#####################
'###   Auto-View   ###
'#####################
' v1.5
'
'This script lets you auto-label drawing views with the built-in view labels.  It will pull
'the label names automatically for views that it recognizes (front, top, right, etc).
'Section, Detail, and Auxially views are ignored.  Iso and custom camera views cannot be
'identified automatically, so the user will be prompted for the name.
'
'This routine will only label the views that are currently selected.  If no views are selected,
'all views on the sheet will be labeled.

Sub Main()

	Dim oDoc As Inventor.Document = ThisDoc.Document
	
	Select Case oDoc.DocumentType
	
		Case kDrawingDocumentObject
			AutoView(oDoc)
	
	    Case kDesignElementDocumentObject, kForeignModelDocumentObject, kNoDocument, kPresentationDocumentObject, kSATFileDocumentObject, kUnknownDocumentObject, kAssemblyDocumentObject, kPartDocumentObject
	        MsgBox ("This tool is not compatible with this type of file.")
	        Exit Sub
	
	End Select

End Sub 'Main()'


Sub AutoView(ByRef oDoc As Inventor.Document)

	Dim oSheets As Sheets
	Dim oSheet As Sheet = oDoc.ActiveSheet
	Dim oViews As DrawingViews
	Dim oView As DrawingView
	

	Dim oSet As HighlightSet = oDoc.CreateHighlightSet
	oSet.Color = ThisApplication.TransientObjects.CreateColor(255, 0, 0)

	' See if anything is selcted
	Dim oSSet As SelectSet = ThisDoc.Document.SelectSet
	If oSSet.count = 0 Then

		'if nothing is selected, then run the SetViewName function on all views
		oViews = oSheet.DrawingViews
	    For Each oView In oViews
	    	oSet.AddItem(oView)
	    	
			SetViewName(oView)
			oSet.Remove(oView)
		Next
	Else

		'if there are selected objects, iterate through them and run the SetViewName on the view objects
		For Each temp In oSSet
			Dim cView As DrawingView= trycast(temp, DrawingView)
			If cView IsNot Nothing Then
				oSet.AddItem(oView)
				'oSet.Color = ThisApplication.TransientObjects.CreateColor(255, 0, 0)
				SetViewName(cView)
				oSet.Remove(oView)
			End If 'cView IsNot Nothing
		Next 'Each in oSSet
	End If 'oSSet.count = 0
End Sub


Function SetViewName(ByVal xView As DrawingView)

	xView.ShowLabel = True
	Dim myDrawingViewName As String
	Dim IsFlatPatternView As Boolean
	myDrawingViewName = xView.Label.Text

	If Regex.IsMatch(myDrawingViewName, "(\d+(?:\.\d*)?|\.\d+)") = False Then
		Return False
	End If
	
	If xView.ViewType() = drawingviewtypeenum.kSectionDrawingViewType Then
	ElseIf xView.ViewType() = drawingviewtypeenum.kAuxiliaryDrawingViewType Then
	ElseIf xView.ViewType() = drawingviewtypeenum.kDetailDrawingViewType Then
	Else

		'custom field to ignore certain strings in the view labels, disabled for now
		'If xView.Label.FormattedText Like "*DETAIL*" Then
		'    Else

		Select Case xView.Camera.ViewOrientationType
		    Case ViewOrientationTypeEnum.kBackViewOrientation
	    	    myDrawingViewName = "BACK"
		    Case ViewOrientationTypeEnum.kBottomViewOrientation
		        myDrawingViewName = "BOTTOM"
    		Case ViewOrientationTypeEnum.kFrontViewOrientation
		        myDrawingViewName = "FRONT"
		    Case ViewOrientationTypeEnum.kIsoBottomLeftViewOrientation
    		    myDrawingViewName = "ISO BOTTOM LEFT"
		    Case ViewOrientationTypeEnum.kIsoBottomRightViewOrientation
    		    myDrawingViewName = "ISO BOTTOM RIGHT"
	    	Case ViewOrientationTypeEnum.kIsoTopLeftViewOrientation
	        	myDrawingViewName = "ISO TOP LEFT"
	    	Case ViewOrientationTypeEnum.kIsoTopRightViewOrientation
	        	myDrawingViewName = "ISO TOP RIGHT"
		    Case ViewOrientationTypeEnum.kLeftViewOrientation
	    	    myDrawingViewName = "LEFT"
    		Case ViewOrientationTypeEnum.kRightViewOrientation
	        	myDrawingViewName = "RIGHT"
		    Case ViewOrientationTypeEnum.kTopViewOrientation
	    	    myDrawingViewName = "TOP"
			Case ViewOrientationTypeEnum.kSavedCameraViewOrientation
    	    	myDrawingViewName = "AUX"
			Case ViewOrientationTypeEnum.kArbitraryViewOrientation
			
				If Regex.IsMatch(myDrawingViewName, "(\d+(?:\.\d*)?|\.\d+)") Then
				
					myDrawingViewName = InputBox("Please enter view name.", "View Name", xView.Name)
				End If
'		        myDrawingViewName = InputBox("Please enter view name.", "View Name", xView.Name)
    		Case Else
        		myDrawingViewName = ""
	    End Select 'xView.Camera.ViewOrientationType


		'Change if view orientation was found
		If Not myDrawingViewName = "" Then
    		'change drawing view name
    		xView.Name =  myDrawingViewName 
		End If

		If xView.isflatpatternview = True Then
			xView.Name = "FLAT PATTERN"
		End If

		'olabel = "<StyleOverride Underline='True'>" & oview.Name & "</StyleOverride>"
		olabel = "<Br/><Br/>" & "<StyleOverride Underline='False'>" & "<DrawingViewName/>" & "</StyleOverride>"
		'oPartNumber = "<Property Document='model' PropertySet='Design Tracking Properties' Property='Part Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='5'>PART NUMBER</Property></StyleOverride>"
		'oscale = "<Br/><DrawingViewScale/>"


		'add to the view label'Part number only on flat pattern view
		'If xView.Name="FLAT PATTERN" Then xView.Label.FormattedText = olabel '& oPartNumber' & oscale
		'If Not xView.Name="FLAT PATTERN" Then xView.Label.FormattedText = olabel & oscale

		xView.Label.FormattedText = olabel & oscale

		'End If 'xView.Label.FormattedText Like "*DETAIL*"
	End If 'xView.ViewType() = 


End Function
