'Source: https://github.com/InventorCode/iLogicRules
'Title: Add External Rule Paths
'Author: Matthew D. Jordan
'Description: Adds a list of entries into the iLogic external rule directories list.


Dim ruleDirList As List(Of String) = New List(Of String)
ruleDirList.Add("C:\Directory\iLogic Rules\")
ruleDirList.Add("C:\Direcotry\iLogic Rules 2\")

For Each i As String In ruleDirList

    Dim ruleDirToAdd As String = i 
    Dim oFileOptions As iLogicFileOptions = iLogicVb.Automation.FileOptions
    Dim ruleDirectoryExists As Boolean = False
    Dim x As String

    'Determine if the external rule directory already exists in the list
    For Each x In oFileOptions.ExternalRuleDirectories
        'Look for a 
        If x = ruleDirToAdd Then
            ruleDirectoryExists = True
        End If
    Next x

    'Add the external rule directory if it does not already exist in the list
    If ruleDirectoryExists = False Then
        'Create a copy of the ExternalRuleDirectories array, and increase its length by 1
        Dim newDirs As String() = oFileOptions.ExternalRuleDirectories
        ReDim preserve newDirs(newDirs.Length)

        'Add the new directory to the new array
        newDirs(newDirs.Length-1) = ruleDirToAdd

        'Replace the old ExternalRuleDirectories array with the new one:
        oFileOptions.ExternalRuleDirectories = newDirs
    End If
Next i
