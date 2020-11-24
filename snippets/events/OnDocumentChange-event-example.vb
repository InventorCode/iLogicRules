inventor = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
appEvents = Inventor.ApplicationEvents 
AddHandler appEvents.OnDocumentChange, AddressOf partListCreateEvent

Private Sub partListCreateEvent(
                                    DocumentObject As _Document,
                                    BeforeOrAfter As EventTimingEnum,
                                    ReasonsForChange As CommandTypesEnum,
                                    Context As NameValueMap,
                                    ByRef HandlingCode As HandlingCodeEnum)
    HandlingCode = HandlingCodeEnum.kEventNotHandled

    If (BeforeOrAfter = EventTimingEnum.kAfter And Context(1).Equals("CreatePartListLocalBOM")) Then
        MsgBox("a partlist has been created")
    End If
End Sub