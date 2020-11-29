'##########################
'###   Balloon Update   ###
'##########################

Sub Main()

	Dim doc As Inventor.Document = ThisDoc.Document
	Guard.OnlyDrawingDocuments(doc, "This tool is not compatible with this type of file.  Please run from a drawing.")
	FixItemNumbers

End Sub 'Main()'

Sub FixItemNumbers()

	Dim doc As Inventor.DrawingDocument = ThisDoc.Document
	Dim drawingBoms As Inventor.DrawingBOMs = doc.DrawingBOMs
	Dim drawingBom As Inventor.DrawingBOM
	Dim drawingBomRow As DrawingBOMRow
	Dim bomRow As Inventor.BOMRow
	Dim stockNumber As String
	
	For Each drawingBom In drawingBoms

		For Each drawingBomRow In drawingBom.DrawingBOMRows
			
			bomRow = drawingBomRow.BOMRow
			stockNumber = GetStockNumber(bomRow)
			Try
				If (IsFabricatedPart(stockNumber)) Then
					bomRow.ItemNumber = GetSheetView(bomRow)
				End If
			Catch
			End Try
		Next
	Next
	
End Sub


Public Function GetStockNumber(ByRef bomRow As Inventor.BOMRow)

	Dim doc As Inventor.Document = bomRow.ComponentDefinitions(1).Document
	Return PropertyShim.GetProperty(doc, "Stock Number")

End Function

Public Function GetSheetView(ByRef BOMRow As Inventor.BOMRow)

	Dim doc As Inventor.Document = BOMRow.ComponentDefinitions(1).Document
	Return PropertyShim.GetProperty(doc, "Sheet") & PropertyShim.GetProperty(doc, "View")

End Function


Public Function IsFabricatedPart(ByRef input As String) As Boolean
        Return Text.RegularExpressions.Regex.IsMatch(If(input,"error"), "\b\d{3,4}\-[0123456789]\d\d[a-z|A-Z]*\b")
End Function


Public Class Guard

    Shared Sub OnlyDrawingDocuments(doc As Inventor.Document, message As String)
        Dim tempDocument As Inventor.DrawingDocument
        Try
            tempDocument = doc
        Catch ex As Exception
            Throw New SystemException(message)
        End Try
    End Sub

End Class


Public Class PropertyShim

        Shared ReadOnly PropertyLookup As Dictionary(Of String, String) = New Dictionary(Of String,String)(StringComparer.OrdinalIgnoreCase) From 
        {{"Title", "Inventor Summary Information"},
        {"Subject", "Inventor Summary Information"},
        {"Author", "Inventor Summary Information"},
        {"Keywords", "Inventor Summary Information"},
        {"Comments", "Inventor Summary Information"},
        {"Last Saved By", "Inventor Summary Information"},
        {"Revision Number", "Inventor Summary Information"},
        {"Thumbnail", "Inventor Summary Information"},
        {"Category","Inventor Document Summary Information"},
        {"Manager","Inventor Document Summary Information"},
        {"Company","Inventor Document Summary Information"},
        {"Creation Time","Design Tracking Properties"},
        {"Part Number","Design Tracking Properties"},
        {"Project","Design Tracking Properties"},
        {"Cost Center","Design Tracking Properties"},
        {"Checked By","Design Tracking Properties"},
        {"Date Checked","Design Tracking Properties"},
        {"Engr Approved By","Design Tracking Properties"},
        {"Engr Date Approved","Design Tracking Properties"},
        {"User Status","Design Tracking Properties"},
        {"Material","Design Tracking Properties"},
        {"Part Property Revision Id","Design Tracking Properties"},
        {"Catalog Web Link","Design Tracking Properties"},
        {"Part Icon","Design Tracking Properties"},
        {"Description","Design Tracking Properties"},
        {"Vendor","Design Tracking Properties"},
        {"Document SubType","Design Tracking Properties"},
        {"Document SubType Name","Design Tracking Properties"},
        {"Proxy Refresh Date","Design Tracking Properties"},
        {"Mfg Approved By","Design Tracking Properties"},
        {"Mfg Date Approved","Design Tracking Properties"},
        {"Cost","Design Tracking Properties"},
        {"Standard","Design Tracking Properties"},
        {"Design Status","Design Tracking Properties"},
        {"Designer","Design Tracking Properties"},
        {"Engineer","Design Tracking Properties"},
        {"Authority","Design Tracking Properties"},
        {"Parameterized Template","Design Tracking Properties"},
        {"Template Row","Design Tracking Properties"},
        {"External Property Revision Id","Design Tracking Properties"},
        {"Standard Revision","Design Tracking Properties"},
        {"Manufacturer","Design Tracking Properties"},
        {"Standards Organization","Design Tracking Properties"},
        {"Language","Design Tracking Properties"},
        {"Defer Updates","Design Tracking Properties"},
        {"Size Designation","Design Tracking Properties"},
        {"Categories","Design Tracking Properties"},
        {"Stock Number","Design Tracking Properties"},
        {"Weld Material","Design Tracking Properties"},
        {"Mass","Design Tracking Properties"},
        {"SurfaceArea","Design Tracking Properties"},
        {"Volume","Design Tracking Properties"},
        {"Density","Design Tracking Properties"},
        {"Valid MassProps","Design Tracking Properties"},
        {"Flat Pattern Width","Design Tracking Properties"},
        {"Flat Pattern Length","Design Tracking Properties"},
        {"Flat Pattern Area","Design Tracking Properties"},
        {"Sheet Metal Rule","Design Tracking Properties"},
        {"Last Updated With","Design Tracking Properties"},
        {"Sheet Metal Width","Design Tracking Properties"},
        {"Sheet Metal Length","Design Tracking Properties"},
        {"Sheet Metal Area","Design Tracking Properties"},
        {"Material Identifier","Design Tracking Properties"},
        {"Appearance","Design Tracking Properties"},
        {"Flat Pattern Defer Update","Design Tracking Properties"}}

        Public Shared PropertySetLookup As HashSet(Of String) = New HashSet(Of String) From {
            "Inventor Summary Information",
            "Inventor Document Summary Information",
            "Design Tracking Properties",
            "Inventor User Defined Properties"
        }

        ''' <summary>
        ''' Return the specified document property's value. This is the short form of this function.
        ''' Only requires a document and property propertyName. 
        ''' </summary>
        ''' <param name="doc">Inventor Document</param>
        ''' <param name="propertyName">Name of the Property</param>
        ''' <returns></returns>
        Shared Function GetProperty(ByRef doc As Inventor.Document, ByVal propertyName As String) As Object

            Dim setName As String
            Dim documentPropertySets As Inventor.PropertySets = doc.PropertySets
            'Get propertySet for provided propertyName (if exists)
            If PropertyLookup.TryGetValue(propertyName, setName) Then
                Return documentPropertySets.Item(setName).Item(propertyName).Value
            End If

            'Not found in standard properties, search custom properties
            Dim currentPropertySet As Inventor.PropertySet = documentPropertySets.Item("Inventor User Defined Properties")
            Try
                Return currentPropertySet.Item(propertyName).Value
            Catch
            End Try

            'Still not found, search other custom property sets!
            If documentPropertySets.Count >= PropertySetLookup.Count Then
                For Each currentPropertySet In doc.PropertySets
                    If PropertySetLookup.Contains(currentPropertySet.DisplayName)
                        Return ""
                    End If

                    Try
                        Return currentPropertySet.Item(propertyName).Value
                    Catch ex As Exception
                        Return ""
                    End Try
                Next
            End If

            'Still not found, return nothing...
            Return ""

        End Function

        Shared Function GetProperty(ByRef doc As Inventor.Document, ByVal setName As String, ByVal propertyName As String) As Object

            Dim documentPropertySets As Inventor.PropertySets = doc.PropertySets
            Try
                Dim currentPropertySet As Inventor.PropertySet = documentPropertySets.Item(setName)
                Dim currentProperty As Inventor.Property = currentPropertySet.Item(propertyName)
                Return currentProperty.Value
            Catch ex As Exception
                Return ""
            End Try

        End Function
    Shared Sub SetProperty(ByRef doc As Inventor.Document, ByVal propertyName As String, ByVal value As Object)
        
        Dim setName As String
        Dim documentPropertySets As Inventor.PropertySets = doc.PropertySets

        'If the property exists as a built-in property, set the value
        If PropertyLookup.TryGetValue(propertyName, setName) Then
            Try
                documentPropertySets.Item(setName).Item(propertyName).Value = value
                Exit Sub
            Catch
            End Try
            
        End If

        'Not found in standard properties, search custom properties
        Dim currentPropertySet As Inventor.PropertySet = documentPropertySets.Item("Inventor User Defined Properties")
            Try
                currentPropertySet.Item(propertyName).Value = value
                Exit Sub
            Catch
                currentPropertySet.Add(value, propertyname)
            End Try
    End Sub

    Shared Function CustomPropertyExists(currentPropertySet As Inventor.PropertySet, propertyName As String)

        Dim a As Object
        Try
            a = currentPropertySet.Name(propertyName)
            Return True
        Catch
            Return False
        End Try
    End Function

    Shared Function PropertySetExists(ByRef doc As Inventor.Document, ByVal propertySetName As String)
        For Each PropertySet In doc.PropertySets
            If String.Equals(PropertySet.Name, propertySetName, StringComparison.OrdinalIgnoreCase)
                Return True
            End If
        Next
        Return False
    End Function

        Shared Sub SetProperty(ByRef doc As Inventor.Document, ByVal propertySetName As String, ByVal propertyName As String, ByVal value As Object)
            
            Dim documentPropertySets As Inventor.PropertySets = doc.PropertySets
            'If the property set exists, set the value, or add it if needed
            If PropertySetExists(doc, propertySetName) Then
                Try
                    documentPropertySets.Item(propertySetName).Item(propertyName).Value = value
                    Exit Sub
                Catch
                    documentPropertySets.Item(propertySetName).Add(value, propertyname)
                End Try
            
            Else

                'Create the property set, then the property
                Try
                    documentPropertySets.Add(propertySetName)
                    documentPropertySets.Item(propertySetName).Add(value, propertyname)
                Catch ex As Exception

                End Try
            
            End If
        End Sub
End Class
