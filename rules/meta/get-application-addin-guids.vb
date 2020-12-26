'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Summary: Outputs all application addin GUIDs to a text file in the %Temp% folder

Dim outputFolder As String = System.IO.Path.GetTempPath()
Dim outputFilename As String = "AddinGUIDs.txt"

Dim lines As New List(Of String)
lines.Add("{ClientId} DisplayName (Description)")
lines.Add("Curly braces are included in the ClientId string")
lines.Add("")

Dim sortedAddins = ThisApplication.ApplicationAddIns _
	.OfType(Of ApplicationAddIn) _
	.OrderBy(Function(x As ApplicationAddIn) x.DisplayName)

For Each oAddin As ApplicationAddIn In sortedAddins
	With oAddin
		lines.Add(.ClientId & " " & .DisplayName & If (.Description <> "", " (" & .Description & ")", ""))
	End With
Next

Dim filePath As String = System.IO.Path.Combine(outputFolder,outputFilename)
System.IO.File.WriteAllLines(filePath,lines)
Process.Start(filePath) 'Opens the file in its default application