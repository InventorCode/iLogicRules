Sub Main()
    PlaceFromContentCenter()

End Sub

Public Sub PlaceFromContentCenter()
    Dim asmDoc As AssemblyDocument = ThisApplication.Documents.Add(kAssemblyDocumentObject)

    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

    ' Get the node in the content browser based on the names of the nodes in the hierarchy.
    Dim hexHeadNode As ContentTreeViewNode = ThisApplication.ContentCenter.TreeViewTopNode.ChildNodes.Item("Fasteners").ChildNodes.Item("Bolts").ChildNodes.Item("Hex Head")

    ' Find a specific family.  In this case it's using the display name, but any family
    ' characteristic could be searched for.
    Dim family As ContentFamily
    Dim checkFamily As ContentFamily
    For Each checkFamily In hexHeadNode.Families
        If checkFamily.DisplayName = "DIN EN 24016" Then
            family = checkFamily
            Exit For
        End If
    Next

    Dim i As Integer
    If Not family Is Nothing Then
        ' Place one instance of each member.
        Dim offset As Double
        offset = 0
        Dim row As ContentTableRow
        For Each row In family.TableRows
            ' Create the member (part file) from the table.
            Dim failureReason As MemberManagerErrorsEnum
            Dim failureMessage As String
            Dim memberFilename As String
            memberFilename = family.CreateMember(row, failureReason, failureMessage, kRefreshOutOfDateParts)

            ' Place the part into the assembly.
            Dim transMatrix As matrix = ThisApplication.TransientGeometry.CreateMatrix
            transMatrix.Cell(2, 4) = offset
            Dim Occ As ComponentOccurrence = asmDef.Occurrences.Add(memberFilename, transMatrix)

            ' Compute the position for the next placement based on the size of the part just placed.
            Dim minY As Double
            Dim maxY As Double
            minY = Occ.RangeBox.minPoint.y
            maxY = Occ.RangeBox.maxPoint.y
            offset = offset + ((maxY - minY) * 1.1)
        Next
    End If
End Sub