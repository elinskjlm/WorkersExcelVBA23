Private Sub Worksheet_Activate()
    'On Error GoTo 0
    
    Application.EnableEvents = False
    turnOffStuff True
    ProtectSheet Me, False
        
    Dim shouldUpdateValue As Boolean
    shouldUpdateValue = Me.Range("shouldUptateSheetInSheet").value
    
    If IsError(Me.Range("a14").value) Then
        Me.Range("c11").value = "תאריך" ' crucial for avoiding errors on cell A14
    End If
    
    If shouldUpdateValue Then
        Me.Range("prevNameCellInSheet").value = Worksheets("מסך ראשי").Range("NameSearchCell").value
        Me.Range("prevOrederCellInSheet").value = Me.Range("c11").value
        refreshSheet Me.name
    End If
    
    ProtectSheet Me, True
    Application.EnableEvents = True
    turnOffStuff False
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("C11")) Is Nothing Then
        
        turnOffStuff True
        ProtectSheet Me, False
        
        If IsError(Me.Range("a14").value) Then
            Application.EnableEvents = False
            Me.Range("c11").value = "תאריך" ' crucial for avoiding errors on cell A14
            Application.EnableEvents = True
        Else
            refreshSheet Me.name
        End If
        
        ProtectSheet Me, True
        turnOffStuff False
        
    End If
End Sub
