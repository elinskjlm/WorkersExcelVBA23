Private Sub Worksheet_Activate()
    'On Error GoTo 0

    Application.EnableEvents = False
    turnOffStuff True
    ProtectSheet Me, False

    Dim shouldUpdateValue As Boolean
    shouldUpdateValue = Me.Range("shouldUptateSheetInSheet").value

    If shouldUpdateValue Then
        Me.Range("prevNameCellInSheet").value = Worksheets("îñê øàùé").Range("NameSearchCell").value
        refreshSheet Me.name
    End If

    ProtectSheet Me, True
    Application.EnableEvents = True
    turnOffStuff False

End Sub