Option Explicit

Private Sub UserForm_Initialize()
    ' Add items to Shift dropdown list
    Me.cmbRoll.List = Application.Transpose(Sheets("âéìéåï èëðé").Range("lsOptRoles").value)
    
    ' Set default value
    Me.cmbRoll.value = "îàáèç"
    
    optBtnActive.value = True
    
    optBtnNotMedic.value = True
    
    'optBtnUncertified.value = True



End Sub


Private Sub btnAdd_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim allNames As Range
    
    Set ws = ThisWorkbook.Worksheets("ôøèé òåáãéí")
    Set tbl = ws.ListObjects("tbEmploDetails")
    ' newRow is set in the end
    Set allNames = ws.Range("clmnFullNamesFP")
      
    Dim famName, priName, fullName As String
    Dim idNum, formattedID As String
    Dim telNum As String
    Dim emailAddr As String
    
    Dim streetName, streetNum, apartmentNum, region As String
    Dim status As String
    Dim cntcTel, cntcAddr, cntcName, cntcRelation As String

    Dim formattedStartDate As Date
    Dim formattedMedicalDate As Date
    Dim medic, certified, roll As String
   
   
    Dim clmnIndexFam, clmnIndexPri, clmnIndexID, clmnIndexTel As Long
    Dim clmnIndexEmail, clmnIndexStName, clmnIndexStNum, clmnIndexRegion As Long
    Dim clmnIndexStatus, clmnIndexCntTel, clmnIndexCntAddr, clmnIndexCntName, clmnIndexCntRelat As Long
    Dim clmnIndexDateStart, clmnIndexMedic, clmnIndexCertified, clmnIndexAllowed As Long
    Dim clmnIndexRoll, clmnIndexDateMedic As Long
    
    clmnIndexFam = tbl.ListColumns("ùí îùôçä").Index
    clmnIndexPri = tbl.ListColumns("ùí ôøèé").Index
    clmnIndexID = tbl.ListColumns("îñ' æäåú").Index
    clmnIndexTel = tbl.ListColumns("îñ' èìôåï").Index
    clmnIndexEmail = tbl.ListColumns("ëúåáú îééì").Index
    clmnIndexStName = tbl.ListColumns("ëúåáú - ùí øçåá").Index
    clmnIndexStNum = tbl.ListColumns("ëúåáú - îñ' øçåá").Index
    clmnIndexRegion = tbl.ListColumns("ëúåáú - ùëåðä/éùåá").Index
    clmnIndexStatus = tbl.ListColumns("äàí ôòéì").Index
    clmnIndexCntTel = tbl.ListColumns("àéù ÷ùø-èìôåï").Index
    clmnIndexCntAddr = tbl.ListColumns("àéù ÷ùø-ëúåáú").Index
    clmnIndexCntName = tbl.ListColumns("àéù ÷ùø - ùí îìà").Index
    clmnIndexCntRelat = tbl.ListColumns("àéù ÷ùø - èéá ÷ùø").Index
    clmnIndexDateStart = tbl.ListColumns("úàøéê úçéìú úòñå÷ä ðåëçéú").Index
    clmnIndexMedic = tbl.ListColumns("äàí çåáù").Index
    clmnIndexCertified = tbl.ListColumns("äàí îåñîê ò.ùìéèä").Index
    clmnIndexAllowed = tbl.ListColumns("äàí îåøùä ò.ùìéèä").Index
    clmnIndexRoll = tbl.ListColumns("úô÷éã").Index
    clmnIndexDateMedic = tbl.ListColumns("úàøéê äö' áøéàåú ðåëçéú").Index
   
   
    ' Check that txtFamName is not empty
    If Me.txtFamName.value = "" Then
        MsgBox "ëúåá ùí îùôçä.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùí"
        Me.txtFamName.SetFocus
        Exit Sub
    Else
        famName = Application.WorksheetFunction.Trim(Me.txtFamName.value)
        Me.txtFamName.value = famName
    End If
    
    
    ' Check that txtPriName is not empty
    If Me.txtPriName.value = "" Then
        MsgBox "ëúåá ùí ôøèé.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ùí"
        Me.txtPriName.SetFocus
        Exit Sub
    Else
        priName = Application.WorksheetFunction.Trim(Me.txtPriName.value)
        Me.txtPriName.value = priName
    End If
    
    fullName = famName & " " & priName
    fullName = Application.WorksheetFunction.Trim(fullName)

    ' Check valid ID
    idNum = RemoveNonDigits(Me.txtIDNum.value)
    If idNum = "" Then
        MsgBox "ëúåá îñ' úòåãú æäåú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ú.æ."
        Me.txtIDNum.SetFocus
        Exit Sub
    Else
        If checkID(idNum) Then
            formattedID = Format(idNum, "[$-,100]0-0000000-0")
            Me.txtIDNum.value = formattedID
        Else
            MsgBox "îñ' úòåãú æäåú ìà ú÷éï. áãå÷ ùåá.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ú.æ."
            Me.txtIDNum.SetFocus
            Exit Sub
        End If
    End If

    ' =============>>>>>>>>>>>>>>>>>>>>>>

    ' Check for double names
    Dim clmnName As Range
    Dim clmnID As Range
    Dim dataArray() As Variant
    
    ' Set the ranges for the two columns
    Set clmnName = tbl.ListColumns("ùí îìà").DataBodyRange
    Set clmnID = tbl.ListColumns("îñ' æäåú").DataBodyRange
    
    ' Assign column values to the array
    dataArray = Range(clmnName, clmnID).value
    
    Dim foundCell As Range
    Set foundCell = clmnName.Find(fullName, LookIn:=xlValues, LookAt:=xlWhole) ' Search for "abc" in column1
    
    
    If Not foundCell Is Nothing Then
        Dim correspondingValue As Variant
        correspondingValue = WorksheetFunction.Index(clmnID, foundCell.row - clmnName.Cells(1).row + 1) ' Get the corresponding value from column2
        If correspondingValue = idNum Then ' Don't use formattedID, but rather idNum because of the "-"
            MsgBox "ðøàä ùäîàáèç ëáø ÷ééí áîòøëú." & vbNewLine & "îàáèç áòì ùí æää åîñôø æäåú æää ëáø øùåí." & _
            "ìà ðéúï ìøùåí àú àåúå àãí ôòîééí.", vbMsgBoxRtlReading + vbMsgBoxRight, "ëôéìåú!"
            Me.txtFamName.SetFocus
            Exit Sub
        Else
            MsgBox "ðøàä ùäîàáèç ëáø ÷ééí áîòøëú." & vbNewLine & "÷ééí îàáèç áòì ùí æää, òí îñôø æäåú ùåðä." & _
            "ðãøù ìúú ìîàáèç äçãù ùí ðôøã ùéæåää ø÷ òéîå, áöåøä áøåøä.", vbMsgBoxRtlReading + vbMsgBoxRight, "ëôéìåú!"
            Exit Sub
        End If
    End If
    
    ' =============>>>>>>>>>>>>>>>>>>>>>>
    
    
    ' Check valid Tel
    If Me.txtTelNum.value = "" Then
        MsgBox "ëúåá îñ' èìôåï ùì äîàáèç.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú èìôåï"
        Me.txtTelNum.SetFocus
        Exit Sub
    Else
        telNum = RemoveNonDigits(Me.txtTelNum.value)
        If ValidatePhoneNumber(telNum) Then
            telNum = Left(telNum, Len(telNum) - 7) & "-" & Right(telNum, 7)
            Me.txtTelNum.value = telNum
        Else
            MsgBox "îñ' èìôåï ðøàä ìà ú÷éï. ðñä ìú÷ï.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú èìôåï"
            Me.txtTelNum.SetFocus
            Exit Sub
        End If
    End If
    
  
    
    ' Check valid Email
    If Me.txtEmail.value = "" Then
        If Me.chkBypass.value = False Then
            MsgBox "ëúåá ëúåáú àéîééì.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéîééì"
            Me.txtEmail.SetFocus
            Exit Sub
        End If
    Else
        emailAddr = Replace(Me.txtEmail.value, " ", "")
        If ValidateEmailAddress(emailAddr) Then
            Me.txtEmail.value = emailAddr
        Else
            MsgBox "ëúåáú àéîééì ðøàéú ìà ú÷éðä. ðñä ìú÷ï.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéîééì"
            Me.txtEmail.SetFocus
            Exit Sub
        End If
    End If
    
    ' Check that txtStName is not empty
    If Me.txtStName.value = "" Then
        MsgBox "ëúåá àú ùí äøçåá.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ëúåáú"
        Me.txtStName.SetFocus
        Exit Sub
    Else
        streetName = Application.WorksheetFunction.Trim(Me.txtStName.value)
        Me.txtStName.value = streetName
    End If
    
    ' Check that txtRegion is not empty
    If Me.txtRegion.value = "" Then
        MsgBox "ëúåá àú ùí äùëåðä (àí áé-í), àå äééùåá (àí ìà áé-í).", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ëúåáú"
        Me.txtRegion.SetFocus
        Exit Sub
    Else
        region = Application.WorksheetFunction.Trim(Me.txtRegion.value)
        Me.txtRegion.value = region
    End If
    ' Use additinal info of address
    If Me.txtBuildingNum.value <> "" Then
        streetNum = Me.txtBuildingNum.value
    End If
    
    If Me.txtDateStart.value = "" Or Not IsDate(Me.txtDateStart.value) Then
        MsgBox "äæï úàøéê úçéìú úòñå÷ä áôåøîè dd/mm/yyyy.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úàøéê úçéìú úòñå÷ä"
        Me.txtDateStart.SetFocus
        Exit Sub
    Else
        formattedStartDate = Format(CDate(Me.txtDateStart.value), "DD/MM/YYYY")
    End If
    
    ' Check that a valid roll has been selected
    If Me.cmbRoll.value = "" Or IsError(Application.Match(Me.cmbRoll.value, Me.cmbRoll.List, 0)) Then
        MsgBox "áçø úô÷éã îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úô÷éã"
        Me.cmbRoll.SetFocus
        Exit Sub
    Else
        roll = Me.cmbRoll.value
    End If
    
    
    If Me.txtDateMedical.value = "" Or Not IsDate(Me.txtDateMedical.value) Then
        MsgBox "äæï úàøéê äöäøú áøéàåú áôåøîè dd/mm/yyyy.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úàøéê äöäøú áøéàåú"
        Me.txtDateMedical.SetFocus
        Exit Sub
    Else
        formattedMedicalDate = Format(CDate(Me.txtDateMedical.value), "DD/MM/YYYY")
    End If
    

    If Me.optBtnActive.value = True Then
        status = "ôòéì"
    ElseIf Me.optBtnInactive.value = True Then
        status = "ìà ôòéì"
    Else
        MsgBox "áçø ñèèåñ", vbMsgBoxRtlReading + vbMsgBoxRight, "áçéøú ñèèåñ"
        Exit Sub
    End If
    
    
    If Me.optBtnMedic.value = True Then
        medic = "çåáù"
    ElseIf Me.optBtnNotMedic.value = True Then
        medic = "ìà çåáù"
    Else
        MsgBox "áçø äàí çåáù", vbMsgBoxRtlReading + vbMsgBoxRight, "áçéøú äàí çåáù"
        Exit Sub
    End If
    
    If Me.chkBypass.value = False Then
        ' Check that txtContactName is not empty
        If Me.txtContactName.value = "" Then
            MsgBox "ëúåá ùí îìà ùì àéù ä÷ùø ìî÷øä äöåøê.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéù ÷ùø îùðé"
            Me.txtContactName.SetFocus
            Exit Sub
        Else
            cntcName = Application.WorksheetFunction.Trim(Me.txtContactName.value)
            Me.txtContactName.value = cntcName
        End If
        ' Check that txtContactRelation is not empty MAYBE LET IT BE OPTIONAL?
        If Me.txtContactRelation.value = "" Then
            MsgBox "ëúåá àú ñåâ ä÷ùø ùì àéù ä÷ùø ìî÷øä äöåøê.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéù ÷ùø îùðé"
            Me.txtContactRelation.SetFocus
            Exit Sub
        Else
            cntcRelation = Application.WorksheetFunction.Trim(Me.txtContactRelation.value)
            Me.txtContactRelation.value = cntcRelation
        End If
        
    
        ' Check valid ContactTel
        If Me.txtContactTelNum.value = "" Then
            MsgBox "ëúåá îñ èìôåï àéù ä÷ùø ìî÷øä äöåøê.", vbMsgBoxRtlReading + vbMsgBoxRight, "ääæðú àéù ÷ùø îùðé"
            Me.txtContactTelNum.SetFocus
            Exit Sub
        Else
            cntcTel = RemoveNonDigits(Me.txtContactTelNum.value)
            If ValidatePhoneNumber(cntcTel) Then
                cntcTel = Left(cntcTel, Len(cntcTel) - 7) & "-" & Right(cntcTel, 7)
                Me.txtContactTelNum.value = cntcTel
            Else
                MsgBox "îñ èìôåï ðøàä ìà ú÷éï. ðñä ìú÷ï.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú èìôåï"
                Me.txtContactTelNum.SetFocus
                Exit Sub
            End If
        End If
        
        ' Check that txtContactAddress is not empty MAYBE LET IT BE OPTIONAL?
        If Me.txtContactAddress.value = "" Then
            MsgBox "ëúåá àú äëúåáú ùì àéù ä÷ùø ìî÷øä äöåøê.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéù ÷ùø îùðé"
            Me.txtContactAddress.SetFocus
            Exit Sub
        Else
            cntcAddr = Application.WorksheetFunction.Trim(Me.txtContactAddress.value)
            Me.txtContactAddress.value = cntcAddr
        End If
    End If
 
    ProtectSheet ws, False
 
    Set newRow = tbl.ListRows.Add
    newRow.Range(clmnIndexFam) = famName
    newRow.Range(clmnIndexPri) = priName
    newRow.Range(clmnIndexID) = formattedID
    newRow.Range(clmnIndexTel) = telNum
    newRow.Range(clmnIndexEmail) = emailAddr
    newRow.Range(clmnIndexStName) = streetName
    newRow.Range(clmnIndexStNum) = streetNum
    newRow.Range(clmnIndexRegion) = region
    newRow.Range(clmnIndexStatus) = status
    newRow.Range(clmnIndexCntTel) = cntcTel
    newRow.Range(clmnIndexCntAddr) = cntcAddr
    newRow.Range(clmnIndexCntName) = cntcName
    newRow.Range(clmnIndexCntRelat) = cntcRelation
    newRow.Range(clmnIndexDateStart) = formattedStartDate
    newRow.Range(clmnIndexMedic) = medic
    newRow.Range(clmnIndexCertified) = "ìà îåñîê"
    newRow.Range(clmnIndexAllowed) = "ìà îåøùä"
    newRow.Range(clmnIndexRoll) = roll
    newRow.Range(clmnIndexDateMedic) = formattedMedicalDate
    
    
    ProtectSheet ws, True
    
    MsgBox "äôøèéí ðøùîå áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú îàáèç çãù"
    
    
    Set ws = Nothing
    Set tbl = Nothing
    Set newRow = Nothing
    Set allNames = Nothing

    Set clmnName = Nothing
    Set clmnID = Nothing
    
    Set foundCell = Nothing



    ClearUserFormInputs Me
    ActiveWorkbook.Save


End Sub




Private Sub txtDateStart_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    ' assignment is a hack to make it work
    Dim a As Variant
    a = tryParseDate(Me.txtDateStart)
  
End Sub

Private Sub txtDateMedical_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    ' assignment is a hack to make it work
    Dim a As Variant
    a = tryParseDate(Me.txtDateMedical)
  
End Sub

Private Sub txtIDNum_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
 ' should format it 0-0000000-0
 ' if possible

End Sub

Private Sub txtTelNum_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
 ' should format it 000-0000000
 ' if possible

End Sub

Private Sub txtContactTelNum_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
 ' should format it 000-0000000
 ' if possible

End Sub
