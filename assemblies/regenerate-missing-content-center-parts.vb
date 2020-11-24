'###################################################
'###   Regenerate Missing Content Center Parts   ###
'###################################################
'Matthew D. Jordan
'MIT License

'Based on:
'http://help.autodesk.com/view/INVNTOR/2019/ENU/?caas=caas/discussion/t5/Inventor-Forum/Resolving-Links-to-CC-Components-in-Sub-Asm/td-p/8231970.html
'by jan.priban - Autodesk Employee - 2018-09-27'

' version 1.2

''' <summary>
''' Entry point into the script.  Ensures the current file is an assembly.
''' </summary>
Sub Main()

	If ThisApplication.ActiveDocument.DocumentType = kAssemblyDocumentObject Then
		RegenerateCCOccurrences()
	Else
		msgbox("This command must be run from an assembly. Exiting...")
		Exit Sub
	End If

End Sub


''' <summary>
''' Subroutine that controls the logic for this program.
''' </summary>
Public Sub RegenerateCCOccurrences()

	Dim StartTime As Double
	StartTime = Timer

	'--- Are there missing parts? ---'
	Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
	Dim missingFileCollection As New List(Of DocumentDescriptor)
	Call GetMissingReferences(oDoc, missingFileCollection)

	If (missingFileCollection.Count = 0) Then
		msgbox("There are no missing Content Center parts.  Exiting...")
		Exit Sub
	End If


	'---  CC Families   ---'
	Dim familyCollection As New List(Of ContentFamily)
	Dim family As ContentFamily
	Dim topNode As ContentTreeViewNode = ThisApplication.ContentCenter.TreeViewTopNode

	Dim x As Integer
	For x = 1 To topNode.ChildNodes.Count
		Call AppendFamilyCollection(topNode.ChildNodes.Item(x), familyCollection)
	Next

	Logger.Info("Collection of CC family objects populated in " & Round(Timer - StartTime, 2) & "sec")
	'Logger.Info("Checked families count: " & familyCollection.Count)


	'---   Iterate through parts   ---'
	'---------------------------------'
	Dim missingFilename As String
	Dim partRowIndex As Integer
	Dim ee As Inventor.MemberManagerErrorsEnum


	' Iterate through the occurrences and print the name.
	Dim missingFileRef As DocumentDescriptor
	For Each missingFileRef In missingFileCollection

		'Get occurence file name
		missingFilename = System.IO.Path.GetFileNameWithoutExtension(missingFileRef.FullDocumentName)
		missingFilename = Replace(missingFilename, "_", "/")

		Logger.Info("Missing CC file: " & missingFilename)
		Logger.Info("Missing CC part: " & missingFileRef.DisplayName)

		For Each family In familyCollection

			'If InStr(1, i.DisplayName, cFamily.DisplayName) <> 0 Then
			If Not missingFileRef.DisplayName.Contains(family.DisplayName) Then
				Continue For
			End If

			partRowIndex = FindRow(family, missingFilename)
			Logger.Info("--- Found member: " & missingFileRef.DisplayName & " | in family: " & family.DisplayName & " | in row: " & partRowIndex)

			If partRowIndex <= 0 Then
				Continue For
			End If

			'Create instance of missing CC part
			family.CreateMember(partRowIndex, ee, "Problem", Inventor.ContentMemberRefreshEnum.kUseDefaultRefreshSetting)
			Logger.Info(missingFileRef.DisplayName & "   " & ee.ToString)
			Exit For

		Next
	Next

	Dim secondsElapsed As Double
	secondsElapsed = Round(Timer - StartTime, 2)
	Logger.Info("---------")
	Logger.Info("This code ran successfully in " & SecondsElapsed & " seconds")
	Logger.Info("Checked families count: " & familyCollection.Count)
	Logger.Info("---END---")

End Sub


''' <summary>
''' Append a list of all Content Center ContentFamily objects.
''' </summary>
''' <param name="currentNode">The ContentTreeViewNode to go through.</param>
''' <param name="familyCollection">The List to append.</param>
Sub AppendFamilyCollection(currentNode As ContentTreeViewNode, familyCollection As List(Of ContentFamily))

	'If currentNode.DisplayName = "Features" Then Exit Sub
	'If currentNode.DisplayName = "Structural Shapes" Then Exit Sub

	For Each family In currentNode.Families
		familyCollection.Add(family)
	Next

	For Each childNode In currentNode.ChildNodes
		Call AppendFamilyCollection(childNode, familyCollection)
	Next

End Sub



''' <summary>
''' Returns the Content Family row that matches a filename.
''' </summary>
''' <param name="family">ContentFamily to search.</param>
''' <param name="filename">Filename to search for.</param>
''' <returns>Integer representing the matching CC Family row.</returns>
Function FindRow(ByVal family As ContentFamily, ByVal filename As String) As Integer

	Dim row As ContentTableRow

	For Each row In family.TableRows
		If row.GetCellValue(family.FileNameColumn) = filename Then
			Return row.Index
			Exit For
		End If
	Next

End Function


''' <summary>
''' Get the missing component document descriptors from the assembly and populate into a list.
''' </summary>
''' <param name="document">AssemblyDocument</param>
''' <param name="missingFileCollection"></param>
Private Sub GetMissingReferences(document As AssemblyDocument, ByRef missingFileCollection As List(Of DocumentDescriptor))

	Dim componentDefinition As AssemblyComponentDefinition = document.ComponentDefinition

	' Get all of the leaf occurrences of the assembly.
	Dim leafOccurences As ComponentOccurrencesEnumerator = componentDefinition.Occurrences.AllLeafOccurrences
	Dim occurence As ComponentOccurrence
	Dim documentDescriptor As DocumentDescriptor

	For Each occurence In leafOccurences
		documentDescriptor = occurence.ReferencedDocumentDescriptor

		If Not occurence.ReferencedDocumentDescriptor.ReferenceMissing Then
			Continue For
		End If

		If missingFileCollection.Contains(documentDescriptor) Then
			Continue For
		End If

		'Add the missing file description into the collection
		missingFileCollection.Add(documentDescriptor)
	Next

End Sub