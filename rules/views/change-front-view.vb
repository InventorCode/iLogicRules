'#############################
'###   Change Front View   ###
'#############################
' v1.1
'
'MDJ
'Set the Front view to match whatever is selected!  This routine is intended
'to make it easier to set a Front view for cnc parts.  There are two modes:
'	1. If a face or edge is selected, the Front view is set normal to that selection,
'	2. If nothing is selected, the current view is set as front.

'based on https://forums.autodesk.com/t5/inventor-customization/setting-view-cube-orientation-with-ilogic/m-p/8187506'


Sub Main()

	Dim oApp As Application = ThisApplication
	Dim oDoc As Document = ThisDoc.Document

	Select Case oDoc.DocumentType
	
		Case kDrawingDocumentObject, kDesignElementDocumentObject, kNoDocument, kUnknownDocumentObject
	        MsgBox ("This tool is not compatible with this type of file.")
	

	    Case kForeignModelDocumentObject, kPresentationDocumentObject, kSATFileDocumentObject, kAssemblyDocumentObject, kPartDocumentObject
			If ThisApplication.ActiveDocument.FileSaveCounter = 0 Then
				'This file has not been saved
				MsgBox("This file has not been saved yet, aborting script.")
			Else
				ChangeFrontWrapper(oApp, oDoc)
			End If

	        Exit Sub
	
	End Select
End Sub
Public Sub ChangeFrontWrapper(ByRef oApp As Application, ByRef oDoc As Inventor.Document)

	Dim oldSelectionPriority = oDoc.SelectionPriority
	oDoc.SelectionPriority = 67587

	ChangeFront(oApp, oDoc)
	ChangeHome(oApp, oDoc)

	oDoc.SelectionPriority = oldSelectionPriority
	oApp.ActiveDocument.Save

End Sub


Sub ChangeFront(ByRef oApp As Application, ByRef oDoc As Inventor.Document)

    Dim oSSet As SelectSet = oApp.ActiveDocument.SelectSet 'get the current selection set
    Dim tempCollection As Collection

' See if anything is selcted
    If oSSet.count = 0 Then

        'if nothing is selected, make the current view FRONT
		oApp.CommandManager.ControlDefinitions("AppViewCubeViewFrontCmd").Execute

    Else if oSSet.count = 1 Then

        'if there are selected objects, get the documents from them and run your code
		    Dim LookAt As ControlDefinition = oApp.CommandManager.ControlDefinitions("AppLookAtCmd")
		    LookAt.Execute()
		    oSSet.Clear()
		    Dim activeView As View = oApp.ActiveView
		    activeView.SetCurrentAsFront()

			oDoc.Save
	Else
		msgBox("Too many objects selected!  Exiting now.")

    End If 'oSSet.count


End Sub

Public Sub ChangeHome(ByRef oApp As Application, ByRef oDoc As Inventor.Document)
	
		'set to iso view
		Dim oCamera As Camera 
		oCamera = oApp.ActiveView.Camera 
		oCamera.ViewOrientationType = 10760 'Iso Top Left View Orientation 
		oCamera.Apply
		'set current iso as home view
		oApp.CommandManager.ControlDefinitions("AppViewCubeViewHomeFloatingCmd").Execute
		
		'return to front
		oCamera.ViewOrientationType = 10764 'Front View Orientation 
		oCamera.Apply

		oDoc.Save

End Sub


