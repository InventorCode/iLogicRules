'#############################
'###   Disable Adaptivity  ###
'#############################
' Source: https://github.com/InventorCode/iLogicRules
'
' Removes all adaptivity for components in this assembly.
' This is a sledgehammer, use wisely.

Dim doc As Document = ThisApplication.ActiveDocument

If doc.DocumentType <> kAssemblyDocumentObject Then
	MessageBox.Show("Please run this rule in an assembly file.  Exiting...", "Disable Adaptivity")
	Return
End If


Dim i As Integer = 0

For Each componentDoc As Document In doc.AllReferencedDocuments
	Try
		If componentDoc.ModelingSettings.AdaptivelyUsedInAssembly = True Then
			componentDoc.ModelingSettings.AdaptivelyUsedInAssembly = False
			i += 1
		End If
	Catch
	End Try
Next

Logger.Info("disable-adaptivity.vb: Edited " & i & " file(s).")
