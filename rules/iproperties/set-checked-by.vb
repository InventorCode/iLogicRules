'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: Fills out the "Checked By" and "Checked Date" iProperties,
'and displays the results in the iLogic log window.
'User name can be set in Application Options - General
'
'Optional RuleArguments:
'	doc As Document: The document whose iProps will be changed.
'	(The document that called this rule will be used if no doc is passed to this rule.)

Option Explicit On

Private Class RuleMain
	Private _doc As Document = Nothing
	
	Private Sub Main
		_doc = RuleArguments("doc")
		If _doc Is Nothing Then _doc = ThisDoc.Document
			
		Dim docName As String
		If _doc.FileSaveCounter > 0 Then
			docName = System.IO.Path.GetFileName(_doc.FullFileName)
		Else
			docName = _doc.Displayname
		End If
		
		Logger.Info(docName)
		
		If Not _doc.IsModifiable Then
			'A generated iPart or iAssembly member file is an example of a file that is not modifiable
			Logger.Error("iProps can't be updated because document is not modifiable")
			Exit Sub
		End If
		
		If _doc.FileSaveCounter > 0 AndAlso (System.IO.File.GetAttributes(_doc.FullFileName) And System.IO.FileAttributes.ReadOnly) Then
			Logger.Error("iProps can't be updated because file is read-only")
			Exit Sub
		End If
		
		Me.UpdateProp("Checked By", ThisApplication.UserName)
		Me.UpdateProp("Date Checked", DateTime.Today())
		
		ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
	End Sub
	
	Private Sub UpdateProp(propName As String, newValue As Object)
		Dim prop As Inventor.Property = _doc.PropertySets("Design Tracking Properties").Item(propName)
		
		If prop.Expression <> newValue Then
			Try
				prop.Expression = newValue
			Catch ex As Exception
				Logger.Error(vbTab & """" & propName & """ iProp could not be modified")
				Exit Sub
			End Try
			Logger.Info(vbTab & """" & propName & """ iProp updated")
		Else
			Logger.Info(vbTab & """" & propName & """ iProp is already up-to-date")
		End If
	End Sub
End Class
