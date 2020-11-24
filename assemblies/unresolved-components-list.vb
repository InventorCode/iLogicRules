Imports System.IO

' ######################################
' ###   List Unresolved Components   ###
' ######################################
'
' v1.0 - MDJ
'
' This routine will go through an assembly and summarize the unresolved files.
' They are saved to an excel file with the same name as the assembly + "Unresolved".
' The excel file will open after the routine is finished.


Sub Main()

	If Not ThisApplication.ActiveDocument.DocumentType = kAssemblyDocumentObject Then Return

	Dim oFile As Inventor.File = ThisApplication.ActiveDocument.File
	Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
''	oDoc.Save
	Dim oDocDef As AssemblyComponentDefinition = oDoc.ComponentDefinition

	Dim oList As String = System.IO.Path.GetDirectoryName(oDoc.FullFileName) & _
	 "\" & System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName) & " Unresolved.xlsx"
	'Dim excelName As String = 
	'Dim oList As String = excelPath & excelName

		'Create/overwrite the existing excel file
		excelApp = CreateObject("Excel.Application")
	    excelApp.Visible = False
		excelApp.DisplayAlerts = False

		wb = excelApp.Workbooks.Add()
			
		wb.SaveAs (oList)

		'close out the excel app
		ws = Nothing
		wb.Close
		excelApp.DisplayAlerts = True

		excelApp.Quit
		'clean up
		wb = Nothing
		excelApp = Nothing

		GoExcel.Open(oList, "Sheet1")
		GoExcel.CellValue(oList,"Sheet1", "A1") = "Missing File Name"
		GoExcel.CellValue(oList, "Sheet1" , "B1") = "Missing File Path"
		GoExcel.CellValue(oList, "Sheet1" , "C1") = "Parent Assembly"

		Logger.Info("Excel file created..." & oList)

	'create an incrementor object to track the Excel row variable across subroutines
	Dim oListLine As New Incrementor(2)
	Logger.Info("incrementor oListLine created with value = " & oListLine.i)

	'call the main logic sub to determine if a referenced file is unresolved
	Call ProcessReferences(oFile, oDoc, oDocDef, oListLine, oList)
		
	Logger.Info("Save and close excel file...")
	GoExcel.Save
	GoExcel.Close

	Dim objExcel = CreateObject("Excel.Application")
	Dim objWorkbook = objExcel.Workbooks.Open(oList)
	objExcel.Application.Visible = True



''	msgbox("List Unresolved Components has finished")
End Sub

Private Sub ProcessReferences(ByVal oFile As Inventor.File, oDoc As AssemblyDocument, oDocDef As ComponentDefinition, ByRef oListLine As Incrementor, oList as String)

		'get the file name for the missing component
		
		Dim fileName As String

        Dim oFileDescriptor As FileDescriptor
        For Each oFileDescriptor In oFile.ReferencedFileDescriptors

            If Not oFileDescriptor.ReferenceMissing Then

                If Not oFileDescriptor.ReferencedFileType = FileTypeEnum.kForeignFileType Then
                    Call ProcessReferences(oFileDescriptor.ReferencedFile, oDoc, oDocDef, oListLine, oList)

                End If
            Else

            	Dim filepath As String = oFileDescriptor.FullFileName

            	Logger.Info("ReferenceMissing for " & filepath)
				GoExcel.CellValue(oList, "Sheet1" , "A" & oListLine.i) = System.IO.Path.GetFileName(filepath)
				GoExcel.CellValue(oList, "Sheet1" , "B" & oListLine.i) = System.IO.Path.GetDirectoryName(filepath)
				GoExcel.CellValue(oList, "Sheet1" , "C" & oListLine.i) = oFileDescriptor.Parent.FullFileName 'oDoc.FullFileName
				oListLine.increment

            End If

        Next

End Sub

Public Class Incrementor

#Region "Declare Private Variables"
    Private _i As Long
#End Region

	Public Sub New (temp As Long)
		_i = temp
	End Sub

	Public Sub increment
		_i = _i + 1
	End Sub


#Region "Properties"
	Public Property i() As Long
        Get
            Return _i
        End Get
        Set(value As Long)
            _i = value
        End Set
    End Property
#End Region

End Class
