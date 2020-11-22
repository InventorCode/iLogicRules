'#######################################
'###   Save And Replace With Epoch   ###
'#######################################

' version 1.0
' Matthew D. Jordan -(C) 2020
' MIT License

' Renames a seleted part or assembly file using the current epoch milliseconds as the file name.

'TODO: Add prompt for replacing all parts in subassemblies as well.
'TODO: Add support for multiple selections

Imports System.IO

Sub Main()
    Dim oDoc As Inventor.Document
    oDoc = ThisDoc.Document

    If oDoc.IsModifiable = False Then
        MsgBox("This file is not modifiable.")
        Exit Sub
    End If

    'Get the selection from user
    Dim oSelect As Object = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Pick a part occurrence.")
    If oSelect Is Nothing Then Exit Sub

    Select Case oDoc.DocumentType

        Case kAssemblyDocumentObject
            SaveAndReplaceWithEpoch(oSelect)

        Case Else
            MsgBox("This tool is not compatible with this type of file.")
            Exit Sub

    End Select
End Sub


Public Sub SaveAndReplaceWithEpoch(ByRef oSelect As Object)

    'Convert to occurrence
    Dim oOcc As ComponentOccurrence = oSelect

    'Get the document
    Dim oDoc As Document = oOcc.Definition.Document

    'Check the document type -- WHY no allow assemblies????
    If (oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Or oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
    Else
        Throw New System.ArgumentException("Please select a part or assembly.  Routine exiting.")
    End If

    If (oDoc.IsModifiable = False) Then
        Throw New System.ArgumentException("This file is not modifiable.  Routine exiting.")
    End If

    'Get the new name
    Dim oFileName As FileName = New FileName(oDoc)

    'Save under new name
    oDoc.SaveAs(oFileName.FullName, True)

    'Replace all occurrences in the top assy
    oOcc.Replace(oFileName.FullName, True)

    'Rename File
    If RenameToEpochOptions.Rename = True Then
        System.IO.File.Move(oFileName.OldFilePath, oFileName.OldFilePath & ".bak")
    End If
End Sub


Public Class FileName

    Public Property Name As String = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().toString
    Public Property Path As String
    Public Property Extension As String
    Public Property FullName As String
    Public Property OldFilePath As String

    Public Sub New(ByRef oDoc As Inventor.Document)

        'Get values from selected file...
        Me.OldFilePath = oDoc.FullFileName
        'Dim TempPath As String = OldFilePath
        Dim OldName As String = System.IO.Path.GetFileNameWithoutExtension(TempPath)
        Me.Extension = System.IO.Path.GetExtension(OldFilePath)
        Me.Path = System.IO.Path.GetDirectoryName(OldFilePath) + System.IO.Path.DirectorySeparatorChar


        'Append/Prepend Name...
        Me.Name = RenameToEpochOptions.Prepend & Me.Name & RenameToEpochOptions.Append

        'Get file name from user, exit if cancel or nothing is entered.
        Dim TempName As String = GetNameFromUser()
        If (TempName.Equals(vbCancel) Or TempName.Equals(" ")) Then
            Throw New System.ArgumentException("Please enter a filename.  Routine exiting.")
        End If

        'test for empty path
        If (Me.Path = "" Or Me.Path = vbNullString) Then
            Throw New System.ArgumentException("The directory path is empty for some reason.  Routine exiting.")
        End If

        Me.FullName = Me.Path & Me.Name & Me.Extension

    End Sub

    Private Function GetNameFromUser() As String
        Return InputBox("Enter new file name: ", "New file name", Name)
    End Function

End Class


Public Class RenameToEpochOptions

    Shared Property Prepend As String = ""
    Shared Property Append As String = ""
    Shared Property Rename As Boolean = True

    Private Sub New()
    End Sub

End Class