Option Explicit
' Variables to be used across the subs of this userForm
Private validName As Boolean
Private validDate As Boolean
Private validShift As Boolean
Private RGB_LIGHT_BLUE As Long
Private RGB_WHITE As Long

Private Sub UserForm_Initialize()
    ' Initialize the form, load names to cmbName
    validName = False
    validDate = False
    validShift = False
    RGB_LIGHT_BLUE = RGB(193, 221, 222)
    RGB_WHITE = RGB(255, 255, 255)

    Dim cboName As ComboBox
    Set cboName = Me.Controls("cmbName")
    
    loadValues cboName, "AH", "AI1", "name"
    
    Set cboName = Nothing
    
End Sub

Private Sub btnChooseName_Click()
    ' Check if a valid name is chosen, if so - load dates to cmbDate
    Dim ws As Worksheet
    
    Dim cboName As ComboBox
    Dim cboDate As ComboBox
    
    Set ws = ThisWorkbook.Sheets("âéìéåï èëðé")
    
    Set cboName = Me.Controls("cmbName")
    Set cboDate = Me.Controls("cmbDate")
    
    
    
    
    ' Check that a valid name has been selected
    If Me.cmbName.value = "" Or IsError(Application.Match(Me.cmbName.value, Me.cmbName.List, 0)) Then
        ' If there is no valid name chosen
        MsgBox "áçø ùí îàáèç îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
        Me.cmbName.SetFocus
        GoTo toEnd
    Else
        ' If the name chosen is valid
        validName = True
        ws.Range("AI1").value = cboName.value

        cboName.Locked = True
        cboName.BackColor = RGB_LIGHT_BLUE
        
        loadValues cboDate, "AJ", "AK1", "date"
    End If

toEnd:
    Set ws = Nothing
    Set cboName = Nothing
    Set cboDate = Nothing

End Sub

Private Sub btnChooseDate_Click()
    ' Check if a valid date is chosen, if so - load shifts to cmbShift
    Dim ws As Worksheet
    
    Dim cboDate As ComboBox
    Dim cboShift As ComboBox
    
    Set ws = ThisWorkbook.Sheets("âéìéåï èëðé")
    
    Set cboDate = Me.Controls("cmbDate")
    Set cboShift = Me.Controls("cmbShift")

    
    ' Check that a valid date has been selected
    If Me.cmbDate.value = "" Then
        ' If there is no valid date chosen
        MsgBox "áçø úàøéê îäøùéîä äðôúçú." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
        Me.cmbDate.SetFocus
        GoTo toEnd
    ElseIf ws.Range("AJ2").value = "" Then
        If cboDate.value = CDate(ws.Range("AJ1").value) Then
            ' If the date chosen is valid
            validDate = True
            ws.Range("AK1").value = CDate(cboDate.value)
    
            cboDate.Locked = True
            cboDate.BackColor = RGB_LIGHT_BLUE
            loadValues cboShift, "AL", "AM1", "shift"
        Else
            MsgBox "áçø úàøéê îäøùéîä äðôúçú." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
            Me.cmbDate.SetFocus
            GoTo toEnd
        End If
    ElseIf IsError(Application.Match(cboDate.value, cboDate.List, 0)) Then
    
        MsgBox "áçø úàøéê îäøùéîä äðôúçú." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
        Me.cmbDate.SetFocus
        GoTo toEnd

    Else
        ' If the date chosen is valid
        validDate = True
        ws.Range("AK1").value = CDate(cboDate.value)

        cboDate.Locked = True
        cboDate.BackColor = RGB_LIGHT_BLUE
        loadValues cboShift, "AL", "AM1", "shift"
    End If
    
    

toEnd:
    Set ws = Nothing
    Set cboDate = Nothing
    Set cboShift = Nothing

End Sub


Private Sub btnChooseShift_Click()
    ' Check if a valid shift is chosen, if so - load the rest of the information
    Dim ws As Worksheet
    Dim cboShift As ComboBox
    Dim cboDate As ComboBox
    
    Set ws = ThisWorkbook.Sheets("âéìéåï èëðé")
    Set cboDate = Me.Controls("cmbDate")
    Set cboShift = Me.Controls("cmbShift")

    ' Check that a valid shift has been selected
    'If Me.cmbShift.value = "" Or IsError(Application.Match(Me.cmbShift.value, Me.cmbShift.List, 0)) Then
    If Me.cmbShift.value = "" Or IsError(Application.Match(Me.cmbShift.value, ws.Range("AL1#"), 0)) Then
        ' If there is no valid shift chosen
        MsgBox "áçø îùîøú îäøùéîä äðôúçú." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
        Me.cmbShift.SetFocus
        GoTo toEnd
    Else
        ' If the shift chosen is valid
        validShift = True
        ws.Range("AM1").value = cboShift.value
        
        cboShift.Locked = True
        cboShift.BackColor = RGB_LIGHT_BLUE
    End If

    ' Check if it is a NoShow event, or Late, and load it appropriately
    If IsNumeric(ws.Range("AR11").value) Then
        Me.optBtnLate.value = True
        Me.txtMinutes.value = ws.Range("AR11").value
    ElseIf IsNumeric(ws.Range("AS11").value) Then
        Me.optBtnNoShw.value = True
        Me.txtMinutes.value = ""
    Else
        Me.optBtnLate.value = False
        Me.optBtnNoShw.value = False
        Me.txtMinutes.value = ""
    End If
    
    ' Load the comment
    Me.txtComment.value = ws.Range("AT11").value
    
toEnd:
    Set ws = Nothing
    Set cboShift = Nothing


End Sub


Private Sub btnMakeChange_Click()
    
    Dim result As Integer
     
    Dim wsLate As Worksheet
    Dim wsTech As Worksheet
    Dim tbl As ListObject
    Dim genArr() As Variant
    Dim searchColumn As Range
    Dim searchTerm As String
    Dim foundCell As Range
    
    ' Set reference to the "lates" table
    Set wsLate = ThisWorkbook.Worksheets("àéçåøéí")
    Set wsTech = ThisWorkbook.Worksheets("âéìéåï èëðé")
    Set tbl = wsLate.ListObjects("tbLate")
    ReDim genArr(1 To 2)

    If validName And validDate And validShift Then
    
        If optBtnNoShw Then
            genArr(1) = "çåø"
            Me.txtMinutes.value = ""
            
            
        ElseIf optBtnLate Then
            If Me.txtMinutes.value = "" Or Not IsNumeric(Me.txtMinutes.value) Then
                MsgBox "äæï àú îñôø ã÷åú äàéçåø.", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
                Me.txtMinutes.SetFocus
                GoTo toEnd
            Else
                genArr(1) = Me.txtMinutes.value
            End If
        
        Else
            MsgBox "áçø áéï çåø ìáéï àéçåø", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
            optBtnNoShw.SetFocus
            GoTo toEnd
            
        End If
        
        
        
        If Me.txtComment.value = "" Then
            genArr(2) = "-"
        Else
            genArr(2) = Me.txtComment.value
        End If
    
    
    
            
        result = MsgBox("äàí àúä áèåç ùáøöåðê ìùîåø àú äòãëåï?" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "àéùåø ùéðåééí")
        
        If result = vbYes Then
            
    
            Set searchColumn = tbl.ListColumns("îñ' ùåøä").DataBodyRange
            searchTerm = wsTech.Range("AZ11").value
            
            Set foundCell = searchColumn.Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Check if the search term was found
            If Not foundCell Is Nothing Then
                ' Get the row number of the found cell within the table
                Dim rowNum As Long
                rowNum = foundCell.row - tbl.HeaderRowRange.row
                
                ProtectSheet ThisWorkbook.Sheets("àéçåøéí"), False
                ProtectSheet ThisWorkbook.Sheets("òåáã - àéçåøéí"), False
                ' Edit the values in the found row
                With tbl.ListRows(rowNum).Range
                    .Cells(6).value = genArr(1)
                    .Cells(9).value = genArr(2)
                End With
                
                ProtectSheet ThisWorkbook.Sheets("òåáã - àéçåøéí"), True
                ProtectSheet ThisWorkbook.Sheets("àéçåøéí"), True
                
            Else

                MsgBox "ùâéàä." & vbNewLine & "ðøàä ùîñôø äùåøä ìòøéëä - ìà ÷ééí áèáìä", vbMsgBoxRtlReading + vbMsgBoxRight, "òøéëú îéãò"
            End If
        
            MsgBox "äùéðåééí ðøùîå áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "òøéëú îéãò"
    
        End If
    Else
        MsgBox "äùìí îéãò çñø èøí áçéøä áòãëåï." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
    End If
toEnd:
        Set wsLate = Nothing
        Set wsTech = Nothing
        Set tbl = Nothing
        btnDelChange_Click
        ActiveWorkbook.Save

End Sub


Private Sub btnDelInfo_Click()

    Dim result As Integer
     
    Dim wsLate As Worksheet
    Dim wsTech As Worksheet
    Dim tbl As ListObject
    Dim genArr() As Variant
    Dim searchColumn As Range
    Dim searchTerm As String
    Dim foundCell As Range
    
    ' Set reference to the "lates" table
    Set wsLate = ThisWorkbook.Worksheets("àéçåøéí")
    Set wsTech = ThisWorkbook.Worksheets("âéìéåï èëðé")
    Set tbl = wsLate.ListObjects("tbLate")
    ReDim genArr(1 To 2)

    
    If validName And validDate And validShift Then
    
        If optBtnNoShw Then
            genArr(1) = "çåø"
            Me.txtMinutes.value = ""
            
            
        ElseIf optBtnLate Then
            If Me.txtMinutes.value = "" Or Not IsNumeric(Me.txtMinutes.value) Then
                MsgBox "äæï àú îñôø ã÷åú äàéçåø.", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
                Me.txtMinutes.SetFocus
                GoTo toEnd
            Else
                genArr(1) = Me.txtMinutes.value
            End If
        
        Else
            MsgBox "áçø áéï çåø ìáéï àéçåø", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
            optBtnNoShw.SetFocus
            GoTo toEnd
            
        End If
        
        
        
        If Me.txtComment.value = "" Then
            genArr(2) = "-"
        Else
            genArr(2) = Me.txtComment.value
        End If
    
    
    
            
        result = MsgBox("äàí àúä áèåç ùáøöåðê ìîçå÷ àú äàéøåò?" & vbNewLine & "ìà ðéúï ìùçæø àú äîéãò ìàçø äîçé÷ä!" & vbNewLine & "ìà éòæåø ìñâåø áìé ìùîåø ùéðåééí, àå ììçåõ òì ctrl+Z.", vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "àéùåø ùéðåééí")
        
        If result = vbYes Then
            
    
            Set searchColumn = tbl.ListColumns("îñ' ùåøä").DataBodyRange
            searchTerm = wsTech.Range("AZ11").value
            
            Set foundCell = searchColumn.Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Check if the search term was found
            If Not foundCell Is Nothing Then
                ' Get the row number of the found cell within the table
                Dim rowNum As Long
                rowNum = foundCell.row - tbl.HeaderRowRange.row
                
                ProtectSheet ThisWorkbook.Sheets("àéçåøéí"), False
                ProtectSheet ThisWorkbook.Sheets("òåáã - àéçåøéí"), False
                ' Edit the values in the found row
                With tbl.ListRows(rowNum).Range
                    tbl.ListRows(rowNum).Delete
                End With
                ThisWorkbook.Worksheets("òåáã - àéçåøéí").Range("AB1").value = ""
                ProtectSheet ThisWorkbook.Sheets("òåáã - àéçåøéí"), True
                ProtectSheet ThisWorkbook.Sheets("àéçåøéí"), True
            Else
                MsgBox "ùâéàä." & vbNewLine & "ðøàä ùîñôø äùåøä ìîçé÷ä - ìà ÷ééí áèáìä", vbMsgBoxRtlReading + vbMsgBoxRight, "òøéëú îéãò"
            End If
        
            MsgBox "äùéðåééí ðøùîå áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "òøéëú îéãò"
    
        End If
    Else
        MsgBox "äùìí îéãò çñø èøí áçéøä áîçé÷ä." & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
    End If
toEnd:
    Set wsLate = Nothing
    Set wsTech = Nothing
    Set tbl = Nothing
    btnDelChange_Click
    ActiveWorkbook.Save

End Sub


Private Sub btnDelChange_Click()
    Dim cboName As ComboBox
    Dim cboDate As ComboBox
    Dim cboShift As ComboBox
    
    Set cboName = Me.Controls("cmbName")
    Set cboDate = Me.Controls("cmbDate")
    Set cboShift = Me.Controls("cmbShift")

    cboName.Locked = False
    cboDate.Locked = False
    cboShift.Locked = False
    
    cboName.BackColor = RGB_WHITE
    cboDate.BackColor = RGB_WHITE
    cboShift.BackColor = RGB_WHITE
    
    
    ClearUserFormInputs Me
    
    Set cboName = Nothing
    Set cboDate = Nothing
    Set cboShift = Nothing
    
    
End Sub

Private Sub loadValues(cmb As ComboBox, clList As String, cellDrop As String, flag As String)

    'On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim dvList As String
    Dim arrList As Variant
    Dim dateValue As Variant

    'Set the worksheet where the data validation list is located
    Set ws = ThisWorkbook.Sheets("âéìéåï èëðé")
    
    ' Clear the existing items in the combo box
    cmb.Clear
    
    If Not IsEmpty(ws.Range(clList & "2").value) Then
        cmb.List = Sheets("âéìéåï èëðé").Range(clList & "1#").value
        If flag = "date" Then
            Dim i As Long
            For i = cmb.ListCount - 1 To 0 Step -1
                
                ''''cmb.List(i) = CDate(Format(cmb.List(i), "dd/mm/yyyy"))
                'cmb.List(i) = CDate(cmb.List(i))

                cmb.List(i) = Format(cmb.List(i), "dd/mm/yyyy")

            Next i
            cmb.ListIndex = 0
        End If
    Else
        If flag = "date" Then
            cmb.AddItem Format(ws.Range(clList & "1").value, "dd/mm/yyyy")
        Else
            cmb.AddItem ws.Range(clList & "1").value
        End If
        cmb.ListIndex = 0
    End If




ErrHandler:
    Set ws = Nothing


End Sub
