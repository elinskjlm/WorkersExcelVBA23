Option Explicit


Private Sub txtComment_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim cboName As ComboBox
    Dim cboScore As ComboBox
    Dim tblScore As ListObject
    Dim clmnScore As Range
    
    Dim cboTheme As ComboBox
    Dim tblTheme As ListObject
    Dim clmnTheme As Range
    
    Set tblScore = ThisWorkbook.Worksheets("âéìéåï èëðé").ListObjects("lsOptRvew")
    Set clmnScore = tblScore.ListColumns(2).DataBodyRange
    
    ' Loop through each ComboBox
    Dim i As Integer
    For i = 1 To 5
        
        Set cboName = Me.Controls("cmbName" & Format(i, "00"))
        Set cboName = AddActiveNamesToComboBox(cboName)
        
        Set cboScore = Me.Controls("cmbScore" & Format(i, "00"))
        cboScore.List = Application.Transpose(clmnScore.value)
        
        Set cboName = Nothing
        Set cboName = Nothing
        Set cboScore = Nothing
    Next i
    
    Set cboTheme = Me.Controls("cmbTheme")
    Set tblTheme = ThisWorkbook.Worksheets("âéìéåï èëðé").ListObjects("lsOptRvwTheme")
    Set clmnTheme = tblTheme.ListColumns(1).DataBodyRange
    
    ' Load the column data into the combo box
    cboTheme.List = Application.Transpose(clmnTheme.value)
    
    Set tblScore = Nothing
    Set clmnScore = Nothing
    Set cboTheme = Nothing
    Set tblTheme = Nothing
    Set clmnTheme = Nothing
    
    
End Sub


Private Sub btnAdd_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim genArr() As Variant
    Dim formattedDate As Date
    
    ' Set reference to the "lates" table
    Set ws = ThisWorkbook.Worksheets("áé÷åøåú")
    Set tbl = ws.ListObjects("tbPerfReview")
    
    Dim score As String

    ' Check that a valid date has been entered #TODO edit comment
    If Me.txtDate.value = "" Or Not IsDate(Me.txtDate.value) Then
        MsgBox "äæï úàøéê áôåøîè dd/mm/yyyy." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úàøéê"
        Me.txtDate.SetFocus
        Exit Sub
    Else
        ' Convert the input value to a date in DD/MM/YYYY format
        formattedDate = Format(CDate(Me.txtDate.value), "DD/MM/YYYY")
    End If
    
    ' Check that txtAchmash is not empty
    If Me.txtAchmash.value = "" Then
        MsgBox "ëúåá àú ùí äàçîù.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àçîù"
        Me.txtAchmash.SetFocus
        Exit Sub
    End If


    ' Declare an array to store the selected names
    Dim selectedNames() As Variant
    Dim count As Integer ' Variable to keep track of the number of selected names
    
    ' Check names and scores, one by one
    Dim i As Integer
    For i = 1 To 5
        Dim cboName As ComboBox
        Dim cboScore As ComboBox
        
        Set cboName = Me.Controls("cmbName" & Format(i, "00")) ' Assign the ComboBox control to a variable
        Set cboScore = Me.Controls("cmbScore" & Format(i, "00"))
        
    
        ' Check name and score
        If cboName.value <> "" Then ' Check if there any value in the name box
            ' Check if the name is valid
            If IsError(Application.Match(cboName.value, cboName.List, 0)) Then ' If invalid, display a messagebox and exit the sub
                MsgBox "ùí ìà ú÷éï áîàáèç îñ' " & i & "." & vbNewLine & "áçø ùí îàáèç îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùí îàáèç"
                cboName.SetFocus
                Exit Sub
            Else ' If valid, check also score
                Dim tempArr() As Variant
                ReDim tempArr(1 To 2)
                If cboScore.value <> "" And Not IsError(Application.Match(cboScore.value, cboScore.List, 0)) Then
                    tempArr(1) = cboName.value
                    tempArr(2) = cboScore.value
                ReDim Preserve selectedNames(0 To count)
                selectedNames(count) = tempArr
                count = count + 1
                Else
                    MsgBox "öéåï ìà ú÷éï áîàáèç îñ' " & i & "." & vbNewLine & "áçø öéåï îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú öéåï"
                    cboScore.SetFocus
                    Exit Sub
                End If

            End If
        End If
        
        Set cboName = Nothing
        Set cboScore = Nothing
    Next i
    
    ' Check if the selectedNames array is empty
    If count = 0 Then
        ' If empty, prompt the user and focus on the first ComboBox
        MsgBox "éù ìäæéï ìôçåú ùí àçã. áçø îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùí îàáèç"
        Me.cmbName01.SetFocus
        Exit Sub
    End If
    ' ===============================



    ' Check that txtTheme is not empty
    If Me.cmbTheme.value = "" Then
        MsgBox "ëúåá àú ðåùà äáé÷åøú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ðåùà"
        Me.cmbTheme.SetFocus
        Exit Sub
    ElseIf IsError(Application.Match(Me.cmbTheme.List, Me.cmbTheme.value, 0)) Then
        MsgBox "ðåùà áé÷åøú ìà ú÷éï" & vbNewLine & "áçø ðåùà îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ðåùà"
        Me.cmbTheme.SetFocus
        Exit Sub
    End If

    ' Check that txtDetail is not empty
    If Me.txtDetail.value = "" Then
        MsgBox "ëúåá ôéøåè äáé÷åøú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ôéøåè"
        Me.txtDetail.SetFocus
        Exit Sub
    End If
    
    ' Check that txtCons is not empty
    If Me.txtCons.value = "" Then
        MsgBox "ëúåá ãáøéí ìùéôåø.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùéôåø"
        Me.txtCons.SetFocus
        Exit Sub
    End If
    
    ' Check that txtPros is not empty
    If Me.txtPros.value = "" Then
        MsgBox "ëúåá ãáøéí ìùéîåø.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùéîåø"
        Me.txtPros.SetFocus
        Exit Sub
    End If
    
    
    If Me.txtComment.value = "" Then
        Me.txtComment.value = "-"
    End If
    
    ProtectSheet ws, False
    
    Dim name As Variant
    For Each name In selectedNames
        Set newRow = tbl.ListRows.Add
        score = name(2)
        newRow.Range(1) = formattedDate
        newRow.Range(3) = Me.txtAchmash.value
        newRow.Range(4) = name(1)
        newRow.Range(5) = Me.cmbTheme.value
        newRow.Range(6) = Me.txtDetail.value
        newRow.Range(7) = Me.txtCons.value
        newRow.Range(8) = Me.txtPros.value
        newRow.Range(10) = score
        newRow.Range(11) = Me.txtComment.value
        Set newRow = Nothing
    Next name
    
    ProtectSheet ws, True
    
    ProtectSheet ThisWorkbook.Worksheets("òåáã - áé÷åøåú"), False
    ThisWorkbook.Worksheets("òåáã - áé÷åøåú").Range("W1").value = ""
    ProtectSheet ThisWorkbook.Worksheets("òåáã - áé÷åøåú"), True
    
    
    
    MsgBox "äáé÷åøú ðøùîä áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú áé÷åøú"
    
    Set ws = Nothing
    Set tbl = Nothing
    
    
    ClearUserFormInputs Me
    ActiveWorkbook.Save


End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim txDte As MSForms.TextBox
    Set txDte = Me.txtDate
    ' Here we only try to parse the date, but ignoring if it fails.
    ' Assingment is only to make the function run. Because
    ' "When calling a function, you need to assign its return value to a variable or use it in an expression"
    Dim parsedDate As Date
    parsedDate = tryParseDate(txDte)
    
    
    'On Error Resume Next
    'Dim dateString As String
    
    'dateString = Replace(Me.txtDate, ".", "/") ' Replace dots with slashes
    'Me.txtDate = Format(DateValue(dateString), "dd/mm/yyyy") ' Use custom date format
    Set txDte = Nothing
    
End Sub


