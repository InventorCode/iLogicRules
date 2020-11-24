'Force Delete
'v 1.0 - MDJ
'
' THIS ROUTINE IS SO GODDAMNED DANGEROUS.  IT WILL DELETE SELECTED PARTS/ASSEMBLIES FROM AN ASSEMBLY VIA THE API.
' THERE MAY BE UNINTENDED CONSEQUENCES FOR USING THIS, ONLY USE AS A LAST RESORT.
'
'

Public Sub Main()
    


    '---   Declarations   ---
    Dim oDoc As Document
    Dim oSSet As SelectSet
    Dim oItem As Object
    Dim debugVar As Boolean
    debugVar = False
    If debugVar = True Then MsgBox ("Select Case")
    
    '---   Determine the filetype   ---
    Select Case ThisApplication.ActiveDocument.DocumentType 'what kind of document?
    
        'current file is a part file...
        Case kPartDocumentObject
        ''    Set oDoc = ThisApplication.ActiveDocument
        ''    Call ChangeToEach(oDoc)
        ''    Call FrameToCNC_props(oDoc)
            
        'current file is an assembly file...
        Case kAssemblyDocumentObject
        'Code to let the user set this for a selected object in an assembly.
            If debugVar = True Then MsgBox ("Is an assembly file...")
                
            oSSet = ThisApplication.ActiveDocument.SelectSet 'get the current selection set
            If oSSet.Count = 0 Then
               'end sub
               '' oSSet.Select (ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select part(s)..."))
            End If

            For Each oItem In oSSet
                'If oItem.Type = ObjectTypeEnum.kComponentOccurrenceObject Or oItem.Type = kComponentOccurrenceProxyObject Then
                    oItem.Delete

                'End If
            Next
        
        Case Else
            MsgBox ("This command can only be run from an assembly or part file.")
    End Select

End Sub