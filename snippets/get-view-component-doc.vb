'####################################################
'###   Get Component Document From Drawing View   ###
'####################################################
'v1.0
'
'Tested in Inventor 2020.3.3
'
'Summary:
'The GetViewComponentDoc function returns the component (part or assembly) document referenced by a drawing view. 
'If the drawing view refers to an .IPN file, the assembly document referenced by the IPN is returned instead.

Imports Inventor.ObjectTypeEnum 'This may be required in the header of your script.  'It was not necessary in Inventor 2020.3.3

'------------------------------------------------------------------
'Sub Main is provided as an example for calling GetViewComponentDoc
'------------------------------------------------------------------
Sub Main
	Dim doc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If doc Is Nothing Then
		MessageBox.Show("A drawing document is not active.")
		Exit Sub
	End If

	For Each sht As Sheet In doc.Sheets
		Logger.Info(sht.Name)
		
		For Each oView As DrawingView In sht.DrawingViews
			Logger.Info(vbTab & oView.Name)
			
			Dim viewDoc As Document = GetViewComponentDoc(oView)
			Logger.Info(vbTab & vbTab & viewDoc.DisplayName)
			
		Next
	Next sht
	
	ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
End Sub

Function GetViewComponentDoc(oView As DrawingView) As Document
	Dim doc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument
	Select Case doc.DocumentType
	Case kAssemblyDocumentObject, kPartDocumentObject
		Return doc
	Case kPresentationDocumentObject
		Return doc.ActiveExplodedView.ReferencedDocumentDescriptor.ReferencedDocument
	Case Else
		Throw New NotImplementedException
	End Select
End Function
