Option Explicit

Private Sub UserForm_Initialize()
    turnOffStuff True
    
    ' Loop through each ComboBox

    Dim cboName As ComboBox
    Dim cboScore As ComboBox
    Dim tblScore As ListObject
    Dim clmnScore As Range
    
    Dim cboTheme As ComboBox
    Dim tblTheme As ListObject
    Dim clmnTheme As Range
    
    Set tblScore = ThisWorkbook.Worksheets("âéìéåï èëðé").ListObjects("lsOptDrill")
    Set clmnScore = tblScore.ListColumns(2).DataBodyRange
    
    ' Loop through each ComboBox
    Dim i As Integer
    For i = 1 To 5
        
        Set cboName = Me.Controls("cmbName" & Format(i, "00"))
        Set cboName = AddActiveNamesToComboBox(cboName)
        
        Set cboScore = Me.Controls("cmbScore" & Format(i, "00"))
        cboScore.List = Application.Transpose(clmnScore.value)
        
        Set cboName = Nothing
        Set cboScore = Nothing
    Next i
    

    Set tblScore = Nothing
    Set clmnScore = Nothing
    Set cboTheme = Nothing
    Set tblTheme = Nothing
    Set clmnTheme = Nothing
    
    
    turnOffStuff False
    
End Sub


Private Sub btnAdd_Click()

    turnOffStuff True
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    '----Dim genArr() As Variant
    Dim formattedDate As Date
    
    ' Set reference to the "tbDrills" table
    Set ws = ThisWorkbook.Worksheets("úøâéìéí")
    Set tbl = ws.ListObjects("tbDrills")
    
    Dim score As String

    ' Check that a valid date has been entered #TODO edit comment
    If Me.txtDate.value = "" Or Not IsDate(Me.txtDate.value) Or Not tryParseDate(Me.txtDate) Then
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


    ' ================================
    ' Current Version
    ' Declare an array to store the selected names
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
    If Me.txtTheme.value = "" Then
        MsgBox "ëúåá àú úøçéù äúøâéì.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ðåùà"
        Me.txtTheme.SetFocus
        Exit Sub
    End If

    ' Check that txtDetail is not empty
    If Me.txtDetail.value = "" Then
        MsgBox "ëúåá ôéøåè äúøâéì.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ôéøåè"
        Me.txtDetail.SetFocus
        Exit Sub
    End If
    
    ' Check that txtCons is not empty
    If Me.txtCons.value = "" Then
        MsgBox "ëúåá ãáøéí ìùéôåø." & vbNewLine & "àí àéï îä ìäåñéó, àôùø ìëúåá à.î.ì.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùéôåø"
        Me.txtCons.SetFocus
        Exit Sub
    End If
    
    ' Check that txtPros is not empty
    If Me.txtPros.value = "" Then
        MsgBox "ëúåá ãáøéí ìùéîåø." & vbNewLine & "àí àéï îä ìäåñéó, àôùø ìëúåá à.î.ì.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùéîåø"
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
        newRow.Range(5) = Me.txtTheme.value
        newRow.Range(6) = Me.txtDetail.value
        newRow.Range(7) = Me.txtCons.value
        newRow.Range(8) = Me.txtPros.value
        newRow.Range(10) = score
        newRow.Range(11) = Me.txtComment.value
    Next name

    ProtectSheet ws, True

    ProtectSheet ThisWorkbook.Worksheets("òåáã - úøâéìéí"), False
    ThisWorkbook.Worksheets("òåáã - úøâéìéí").Range("AC1").value = ""
    ProtectSheet ThisWorkbook.Worksheets("òåáã - úøâéìéí"), True



    MsgBox "äúøâéì ðøùí áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úøâéì"

    Set ws = Nothing
    Set tbl = Nothing
    Set newRow = Nothing

    
    'ws.Columns("A:A").NumberFormat = "m/d/yyyy"

    ClearUserFormInputs Me
    ActiveWorkbook.Save
    
    turnOffStuff False

End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    ' no turnOffStuff
    
    
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



