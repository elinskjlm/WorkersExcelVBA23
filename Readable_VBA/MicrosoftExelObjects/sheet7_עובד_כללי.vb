Private Sub Worksheet_Activate()

    Application.EnableEvents = False
    turnOffStuff True

    ProtectSheet Me, True

    Application.EnableEvents = True
    turnOffStuff False

End Sub