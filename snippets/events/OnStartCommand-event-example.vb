Private WithEvents oUIEvents As Inventor.UserInputEvents

    Private Sub oUIEvents_OnStartCommand(CommandID As CommandIDEnum) Handles oUIEvents.OnStartCommand
        If CommandID = CommandIDEnum.kInsertPartsListCommand Then
            MsgBox("partslist has been added")
        End If
    End Sub
End Class