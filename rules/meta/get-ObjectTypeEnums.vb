'Source: https://github.com/InventorCode/iLogicRules
'Title: Get ObjectTypeEnums
'Author: nannerdw
'Description: Creates a text file in %Temp% that lists all ObjectTypeEnums and their values 

Dim outputFolder As String = System.IO.Path.GetTempPath()
Dim outputFilename As String = "ObjectTypeEnums-" & ThisApplication.SoftwareVersion.DisplayVersion & ".txt"
Dim filePath As String = System.IO.Path.Combine(outputFolder, outputFilename)

Dim enumValues As New List (Of String)

For Each enumValue In System.Enum.GetValues(GetType(Inventor.ObjectTypeEnum))
	enumValues.Add(enumValue & " = " & enumValue.ToString)
Next

System.IO.File.WriteAllLines(filePath, enumValues)
Process.Start(filePath) 'Opens the file in its default application