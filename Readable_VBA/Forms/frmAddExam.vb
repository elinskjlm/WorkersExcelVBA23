Option Explicit

Private Sub txtComment_Change()

End Sub

Private Sub UserForm_Initialize()

    turnOffStuff True
    
    ' Add names to Names dropdown list
    Dim cbo As ComboBox
    Set cbo = Me.cmbName 'Assign the ComboBox control to a variable
    'Call the AddActiveNamesToComboBox function to populate the ComboBox
    Set cbo = AddActiveNamesToComboBox(cbo)
    
    Set cbo = Nothing
    turnOffStuff False
    
End Sub

Private Sub btnAdd_Click()
    
    turnOffStuff True
    
    Dim ws As Worksheet
    Dim newRow As ListRow
    Dim formattedDate As Date
    Dim tbl As ListObject
    
    
    ' Set reference to the "tbShetchExam" table
    Set ws = ThisWorkbook.Worksheets("áçðé ùèç")
    Set tbl = ws.ListObjects("tbShetchExam")

    ' Check that a valid date has been entered
    If Me.txtDate.value = "" Or Not IsDate(Me.txtDate.value) Or Not tryParseDate(Me.txtDate) Then
        MsgBox "äæï úàøéê áôåøîè dd/mm/yyyy.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úàøéê"
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
      
    ' Check that a valid name has been selected
    If Me.cmbName.value = "" Or IsError(Application.Match(Me.cmbName.value, Me.cmbName.List, 0)) Then
        MsgBox "áçø ùí îàáèç îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùí îàáèç"
        Me.cmbName.SetFocus
        Exit Sub
    End If
    
    ' Check that either optBtnPass or optBtnFail is selected
    If Me.optBtnPass.value = False And Me.optBtnFail.value = False Then
        MsgBox "áçø äàí òáø àå ðëùì", vbMsgBoxRtlReading + vbMsgBoxRight, "áçéøú úåöàú áçéðä"
        Exit Sub
    End If

    If Me.txtComment.value = "" Then
        Me.txtComment.value = "-"
    End If
    
    
    ProtectSheet ws, False
    
    ' Add new row to the "tbShetchExam" table
    ' #TODO Check for
    Set newRow = tbl.ListRows.Add
    newRow.Range(1) = formattedDate
    newRow.Range(2) = Me.txtAchmash.value
    newRow.Range(3) = Me.cmbName.value
    
    If Me.optBtnPass.value = True Then
        newRow.Range(5) = "òáø" ' should take from âéìéåï èëðé, #TODO
    ElseIf Me.optBtnFail.value = True Then
        newRow.Range(5) = "ðëùì" ' should take from âéìéåï èëðé, #TODO
    End If
    newRow.Range(7) = Me.txtComment.value
    
    ProtectSheet ws, True
     
    MsgBox "äáåçï ðøùí áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú áåçï ùèç"
    
    ' Clear input fields
    ClearUserFormInputs Me
    
    
    
    Set ws = Nothing
    Set tbl = Nothing
    Set newRow = Nothing
    
    ' Save file changes
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

    
        
    Set txDte = Nothing
    
    ' if an error occurs, continue execution rather than stopping and showing an error message
    'On Error Resume Next
    
    ' Declare a variable to hold the date string
    'Dim dateString As String
    
    ' Replace any dots in the input field with slashes, and assign the result to the dateString variable
    'dateString = Replace(Me.txtDate, ".", "/")
    
    ' Convert the dateString variable to a date format, and set the txtDate field to display the formatted date in the dd/mm/yyyy format
    'Me.txtDate = Format(DateValue(dateString), "dd/mm/yyyy")
    
    
End Sub



