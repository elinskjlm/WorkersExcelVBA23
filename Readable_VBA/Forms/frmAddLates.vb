Option Explicit



Private Sub UserForm_Initialize()
    
    ' Add items to Shift dropdown list
    Me.cmbShift.List = Application.Transpose(Sheets("âéìéåï èëðé").Range("lsOptShifts").value)
    
    Dim i As Integer
    ' Add items to Name dropdown list
    For i = 1 To 5
        ' Add names to Names dropdown list
        Dim cbo As ComboBox
        Set cbo = Me.Controls("cmbName" & Format(i, "00")) 'Assign the ComboBox control to a variable
        'Call the AddActiveNamesToComboBox function to populate the ComboBox
        Set cbo = AddActiveNamesToComboBox(cbo)
        
    Next i
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim wbName As String
    wbName = "äåñôú ãéååç àéçåøéí.xlsm"
    

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Sub btnAdd_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim genArr() As Variant
    Dim formattedDate As Date
    
    ' Set reference to the "lates" table
    Set ws = ThisWorkbook.Worksheets("àéçåøéí")
    Set tbl = ws.ListObjects("tbLate")
    'ReDim genArr(1 To 1)
    


    If Me.txtDate.value = "" Or Not IsDate(Me.txtDate.value) Then
        MsgBox "äæï úàøéê áôåøîè dd/mm/yyyy." & vbNewLine & "ùåøä øé÷ä", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úàøéê"
        Me.txtDate.SetFocus
        Exit Sub
    Else
        ' Convert the input value to a date in DD/MM/YYYY format
        formattedDate = Format(CDate(Me.txtDate.value), "DD/MM/YYYY")
    End If

    If Me.cmbShift.value = "" Or IsError(Application.Match(Me.cmbShift.value, Me.cmbShift.List, 0)) Then
        MsgBox "áçø îùîøú îäøùéîä äðôúçú." & vbNewLine & "ùåøä øé÷ä", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú îùîøú"
        Me.cmbShift.SetFocus
        Exit Sub
    End If



    Dim numCols As Integer
    ' Number of columns in the table. No one should just add or remove columns in the table without making sure it ok with the code.
    Const ORIGINAL_NUM_COLS As Integer = 13
    Const NUM_NAME_BOXES As Integer = 5 'Number of rows (names) in the form.
    numCols = tbl.ListColumns.count
    
    Dim dataToInsert() As Variant
    Dim i, j, row As Integer
    row = 1

    Dim hasValue As Boolean ' any value to write to the table
    hasValue = False


    ' If someonr has tempered with the number of columns - we can't predict the outcome of adding the array,
    ' so we inform the user about it.
    'If numCols = ORIGINAL_NUM_COLS Then

        ' Declacre variables to be used in the loop.
        Dim nameBox As ComboBox
        Dim optBtnNoShw, optBtnLate As Object ' temp type!!!!!!!!!!!!!!!
        Dim minutesBox As Object ' temp type!!!!!!!!!!!!!!!
        Dim lateValue As String ' will contain the value to put in ????? ???'/???
        Dim comment As Object
        




        For i = 1 To NUM_NAME_BOXES
        
            Dim tempArr() As Variant
            ReDim tempArr(1 To ORIGINAL_NUM_COLS)
            ' Assign objects for this iteration
            Set nameBox = Me.Controls("cmbName" & Format(i, "00"))
            Set optBtnNoShw = Me.Controls("optBtn" & Format(i, "00") & "NoShw")
            Set optBtnLate = Me.Controls("optBtn" & Format(i, "00") & "Late")
            Set minutesBox = Me.Controls("txtMinutes" & Format(i, "00"))
            Set comment = Me.Controls("txtComment" & Format(i, "00"))
            
        
            ' Check the name. Empty - next iteration; Valid - continue this iter.; Invalid - prompt user and exit sub.
            If nameBox.value = "" Then
                GoTo Next_i_Iteration
    
            ElseIf IsError(Application.Match(nameBox.value, nameBox.List, 0)) Then
                MsgBox "áçø ùí îàáèç îäøùéîä äðôúçú." & vbNewLine & "áùí îñôø " & i, vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùí îàáèç"
                nameBox.SetFocus
                Exit Sub
                
            Else
                If Not hasValue Then
                    hasValue = True
                End If
                
                tempArr(1) = formattedDate
                tempArr(4) = Me.cmbShift.value
                tempArr(5) = nameBox.value
                
            End If

            ' Check the option buttons:
            ' Late - check correct input of minutes and assign it or prompt the user accordingly;
            ' NoShow - assign it and continue this iter.;
            ' None - prompt user and exit sub.
            If optBtnNoShw Then
                lateValue = "çåø"
                ' #TODO delete minutes late???

            ElseIf optBtnLate Then
                
                If minutesBox.value = "" Or Not IsNumeric(minutesBox.value) Then
                    MsgBox "äæï àú îñôø ã÷åú äàéçåø." & vbNewLine & "áùí îñôø " & i, vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ã÷åú àéçåø"
                    minutesBox.SetFocus
                    Exit Sub
                Else
                    lateValue = minutesBox.value
                End If
            
            Else
                MsgBox "áçø áéï çåø ìáéï àéçåø" & vbNewLine & "áùí îñôø " & i, vbMsgBoxRtlReading + vbMsgBoxRight, "áçéøú çåø àå àéçåø"
                optBtnNoShw.SetFocus
                Exit Sub
                
            End If
            tempArr(6) = lateValue
            
            If comment = "" Then
                comment = "-"
            End If
            
            tempArr(9) = comment
            ReDim Preserve genArr(1 To row)
            genArr(row) = tempArr
             
            'dataToInsert(row, 3) = lateValue
            
            row = row + 1
            
            
            Set nameBox = Nothing
            Set optBtnNoShw = Nothing
            Set optBtnLate = Nothing
            Set minutesBox = Nothing
            Set comment = Nothing

Next_i_Iteration:
        Next i

        If hasValue Then
            ProtectSheet ws, False
            
            Dim lastRowNumber As Long
            Dim firstColumnNumber As Long
            
            lastRowNumber = tbl.ListRows.count + tbl.HeaderRowRange.row
            firstColumnNumber = tbl.Range.Columns(1).Column
        
            For j = 1 To UBound(genArr, 1)
                
                ws.Cells(lastRowNumber + j, firstColumnNumber).value = genArr(j)(1)
                ws.Cells(lastRowNumber + j, firstColumnNumber + 3).value = genArr(j)(4)
                ws.Cells(lastRowNumber + j, firstColumnNumber + 4).value = genArr(j)(5)
                ws.Cells(lastRowNumber + j, firstColumnNumber + 5).value = genArr(j)(6)
                ws.Cells(lastRowNumber + j, firstColumnNumber + 8).value = genArr(j)(9)

            Next j
            
            ProtectSheet ws, True
        End If
    
        ProtectSheet ThisWorkbook.Worksheets("òåáã - àéçåøéí"), False
        ThisWorkbook.Worksheets("òåáã - àéçåøéí").Range("AB1").value = ""
        ProtectSheet ThisWorkbook.Worksheets("òåáã - àéçåøéí"), True
        
        MsgBox "äàéçåøéí ðøùîå áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéçåøéí"
        ' Clear form and Save file changes
        ClearUserFormInputs Me
        ActiveWorkbook.Save

  '  Else
       ' MsgBox "ä÷åã òøåê ìëê ùéù " & ORIGINAL_NUM_COLS & " òîåãåú." & vbNewLine & "òãëï àú îñôøé äòîåãåú á÷åã, åàæ àú äòøê ORIGINAL_NUM_COLS á÷åã.", vbMsgBoxRtlReading + vbMsgBoxRight, "ùâéàä áîñôø äòîåãåú"
        
   ' End If
    
    Set ws = Nothing
    Set tbl = Nothing


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

