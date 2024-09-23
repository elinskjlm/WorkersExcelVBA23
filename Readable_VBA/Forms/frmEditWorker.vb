Option Explicit



' Variables to be used across the subs of this userForm
Private validName As Boolean
Private RGB_LIGHT_BLUE As Long
Private RGB_WHITE As Long






Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Initialize()
    ' Initialize the form, load names to cmbName
    MultiPage1.value = 0

    validName = False

    RGB_LIGHT_BLUE = RGB(193, 221, 222)
    RGB_WHITE = RGB(255, 255, 255)

    Dim cboName As ComboBox
    Dim cboRoll As ComboBox
    
    Set cboName = Me.Controls("cmbName")
    Set cboRoll = Me.Controls("cmbRoll")
    
    cboName.List = Sheets("âéìéåï èëðé").Range("BE1#").value
    cboRoll.List = Application.Transpose(Sheets("âéìéåï èëðé").Range("lsOptRoles").value)
    
    


End Sub

Private Sub btnChooseName_Click()

    Dim wsTech As Worksheet
    Set wsTech = ThisWorkbook.Sheets("âéìéåï èëðé")

    ' Check if a valid name is chosen, if so - load dates to cmbDate
    If Me.cmbName.value = "" Or IsError(Application.Match(Me.cmbName.value, Me.cmbName.List, 0)) Then
        ' If there is no valid name chosen
        MsgBox "áçø ùí îàáèç îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
        Me.cmbName.SetFocus
        GoTo toEnd
    Else
        ' If the name chosen is valid
        validName = True
        wsTech.Range("AO15").value = Me.cmbName.value

        Me.cmbName.Locked = True
        Me.cmbName.BackColor = RGB_LIGHT_BLUE
    End If
    
    Me.Controls("cmbRoll").value = wsTech.Range("AY18")
    Me.Controls("txtTelNum").value = Format(wsTech.Range("AO18"), "000-0000000")
    Me.Controls("txtEmail").value = wsTech.Range("AP18")
    Me.Controls("txtStName").value = wsTech.Range("BA18")
    Me.Controls("txtBuildingNum").value = wsTech.Range("BB18")
    Me.Controls("txtRegion").value = wsTech.Range("AQ18")

    
    If wsTech.Range("AR18").value = "ôòéì" Then
        Me.Controls("optBtnActive").value = True
    Else
        Me.Controls("optBtnInactive").value = True
    End If

    If wsTech.Range("AV18").value = "çåáù" Then
        Me.Controls("optBtnMedic").value = True
    Else
        Me.Controls("optBtnNotMedic").value = True
    End If
    
    If wsTech.Range("AW18").value = "îåñîê" Then
        Me.Controls("chkCertCams").value = True
    Else
        Me.Controls("chkCertCams").value = False
    End If
    
    If wsTech.Range("AX18").value = "îåøùä" Then
        Me.Controls("chkAllowedCams").value = True
    Else
        Me.Controls("chkAllowedCams").value = False
    End If

    
    Me.Controls("txtLastName").value = wsTech.Range("AL18")
    Me.Controls("txtPrivateName").value = wsTech.Range("AM18")
    Me.Controls("txtIDNum").value = Format(wsTech.Range("AN18"), "0-0000000-0")
    Me.Controls("txtDateStart").value = Format(wsTech.Range("AU18"), "dd/mm/yyyy")
    Me.Controls("txtDateMedical").value = Format(wsTech.Range("AZ18"), "dd/mm/yyyy")
    Me.Controls("txtContactName").value = wsTech.Range("BC18")
    Me.Controls("txtContactRelation").value = wsTech.Range("BD18")
    Me.Controls("txtContactTelNum").value = Format(wsTech.Range("AS18"), "000-0000000")
    Me.Controls("txtContactAddress").value = wsTech.Range("AT18")
    
    

toEnd:
    Set wsTech = Nothing


End Sub


Private Sub btnMakeChange_Click()
    ' ###################################################################################################
    ' ###################################################################################################
    ' ###################################################################################################
    ' ###################################################################################################
    
    
    ' Do not allow îåøùä without îåñîê V
    ' Check all values before change
    ' Avoid double names
    ' Deal with 0s on empty cells!
    
    
    Dim wsTech As Worksheet
    Dim wsWorkers As Worksheet
    Dim tblWorkers As ListObject
    
    Set wsWorkers = ThisWorkbook.Worksheets("ôøèé òåáãéí")
    Set tblWorkers = wsWorkers.ListObjects("tbEmploDetails")
    Set wsTech = ThisWorkbook.Sheets("âéìéåï èëðé")

    Dim searchTerm As String
    Dim searchColumn As Range
    Dim foundCell As Range
    
    Dim result As Integer
    Dim changedBoxes As String
    
    Dim status, medic, cert, allowed As String

    


    Dim stayedActive As Boolean
    Dim stayedInactive As Boolean
    Dim changedStaus As Boolean
    stayedActive = Me.Controls("optBtnActive") = True And wsTech.Range("AR18").value = "ôòéì"
    stayedInactive = Me.Controls("optBtnInactive") = True And wsTech.Range("AR18").value = "ìà ôòéì"
    changedStaus = Not (stayedActive Or stayedInactive)

    Dim stayedMedic As Boolean
    Dim stayedNotMedic As Boolean
    Dim changedMedic As Boolean
    stayedMedic = Me.Controls("optBtnMedic").value = True And wsTech.Range("AV18").value = "çåáù"
    stayedNotMedic = Me.Controls("optBtnNotMedic").value = True And wsTech.Range("AV18").value = "ìà çåáù"
    changedMedic = Not (stayedMedic Or stayedNotMedic)
    
    Dim stayedCert As Boolean
    Dim stayedNotCert As Boolean
    Dim changedCert As Boolean
    stayedCert = Me.Controls("chkCertCams").value = True And wsTech.Range("AW18").value = "îåñîê"
    stayedNotCert = Me.Controls("chkCertCams").value = False And wsTech.Range("AW18").value = "ìà îåñîê"
    changedCert = Not (stayedCert Or stayedNotCert)

    Dim stayedAllowed As Boolean
    Dim stayedNotAllowed As Boolean
    Dim changedAllowed As Boolean
    stayedAllowed = Me.Controls("chkAllowedCams").value = True And wsTech.Range("AX18").value = "îåøùä"
    stayedNotAllowed = Me.Controls("chkAllowedCams").value = False And wsTech.Range("AX18").value = "ìà îåøùä"
    changedAllowed = Not (stayedAllowed Or stayedNotAllowed)
    
    Dim telNum As String
    Dim emailAddr As String
    Dim streetName As String
    
    If validName Then
    
        If changedStaus Then
            ' we relay on the fact that the user can't make both False or both True
            changedBoxes = changedBoxes & "• ñèèåñ (ôòéì/ìà ôòéì)" & vbCrLf
        End If
        If optBtnActive Then
            status = "ôòéì"
        Else
            status = "ìà ôòéì"
        End If
        
        If Me.Controls("cmbRoll").value <> wsTech.Range("AY18") Then
            If Me.cmbRoll.value = "" Or IsError(Application.Match(Me.cmbRoll.value, Me.cmbRoll.List, 0)) Then
                MsgBox "áçø úô÷éã îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú úô÷éã"
                Me.cmbRoll.SetFocus
                GoTo toEnd
            Else
                changedBoxes = changedBoxes & "• úô÷éã" & vbCrLf
            End If
        End If
    
    
        If changedMedic Then
            ' we relay on the fact that the user can't make both False or both True
            changedBoxes = changedBoxes & "• äàí çåáù" & vbCrLf
        End If
        If optBtnMedic Then
            medic = "çåáù"
        Else
            medic = "ìà çåáù"
        End If
    
        If changedCert Then
            ' we relay on the fact that the user can't make both False or both True
            changedBoxes = changedBoxes & "• äàí îåñîê òîãåú ùìéèä" & vbCrLf
        End If
        If chkCertCams Then
            cert = "îåñîê"
        Else
            cert = "ìà îåñîê"
        End If
    
        If changedAllowed Then
            ' we relay on the fact that the user can't make both False or both True
            changedBoxes = changedBoxes & "• äàí îåøùä òîãåú ùìéèä" & vbCrLf
        End If
        If chkAllowedCams Then
            allowed = "îåøùä"
        Else
            allowed = "ìà îåøùä"
        End If
    
        If Me.Controls("chkAllowedCams").value = True And Me.Controls("chkCertCams").value = False Then
            MsgBox "ìà éúëï îöá ùì 'îåøùä òîãåú ùìéèä' áìé 'îåñîê òîãåú ùìéèä'", vbMsgBoxRtlReading + vbMsgBoxRight, "èòåú òîãåú ùìéèä"
            Me.Controls("chkCertCams").SetFocus
            GoTo toEnd
        End If
    
    
    
        If Me.Controls("txtTelNum").value <> wsTech.Range("AO18") Then
            If Me.txtTelNum.value = "" Then
                MsgBox "ëúåá îñ' èìôåï ùì äîàáèç.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú èìôåï"
                Me.txtTelNum.SetFocus
                GoTo toEnd
            Else
                telNum = RemoveNonDigits(Me.txtTelNum.value)
                If ValidatePhoneNumber(telNum) Then
                    telNum = Left(telNum, Len(telNum) - 7) & "-" & Right(telNum, 7)
                    Me.txtTelNum.value = telNum
                    changedBoxes = changedBoxes & "• îñ' èì' îàáèç" & vbCrLf
                Else
                    MsgBox "îñ' èìôåï ðøàä ìà ú÷éï. ðñä ìú÷ï.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú èìôåï"
                    Me.txtTelNum.SetFocus
                    GoTo toEnd
                End If
            End If
        End If
    
    



    
        If CStr(wsTech.Range("AP18")) = "0" Then
            If Me.Controls("txtEmail").value = "0" Or Me.Controls("txtEmail").value = "" Then
                 Me.Controls("txtEmail").value = ""
            Else
                emailAddr = Replace(Me.txtEmail.value, " ", "")
                If ValidateEmailAddress(emailAddr) Then
                    Me.txtEmail.value = emailAddr
                    changedBoxes = changedBoxes & "• àéîééì îàáèç" & vbCrLf
                Else
                    MsgBox "ëúåáú àéîééì ðøàéú ìà ú÷éðä. ðñä ìú÷ï.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéîééì"
                    Me.txtEmail.SetFocus
                    GoTo toEnd
                End If
            End If
        Else
            If Me.Controls("txtEmail").value <> wsTech.Range("AP18") Then
                emailAddr = Replace(Me.txtEmail.value, " ", "")
                If ValidateEmailAddress(emailAddr) Then
                    Me.txtEmail.value = emailAddr
                    changedBoxes = changedBoxes & "• àéîééì îàáèç" & vbCrLf
                Else
                    MsgBox "ëúåáú àéîééì ðøàéú ìà ú÷éðä. ðñä ìú÷ï.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú àéîééì"
                    Me.txtEmail.SetFocus
                    GoTo toEnd
                End If
            End If
        End If





        If Me.Controls("txtStName").value <> wsTech.Range("BA18") Then
            If Me.txtStName.value = "" Then
                MsgBox "ëúåá àú ùí äøçåá.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ëúåáú"
                Me.txtStName.SetFocus
                GoTo toEnd
            Else
                changedBoxes = changedBoxes & "• ùí øçåá îàáèç" & vbCrLf
            End If
        End If

        If CStr(Me.Controls("txtBuildingNum").value) <> CStr(wsTech.Range("BB18")) Then
            If Me.txtBuildingNum.value = "" Then
                MsgBox "ëúåá àú îñ' äøçåá.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ëúåáú"
                Me.txtBuildingNum.SetFocus
                GoTo toEnd
            Else
                changedBoxes = changedBoxes & "• îñ' øçåá îàáèç" & vbCrLf
            End If
        End If

        If Me.Controls("txtStName").value = "" And Me.Controls("txtBuildingNum").value <> "" Then
            MsgBox "ìà ìäùàéø îñ' øçåá áìé ùí øçåá", vbMsgBoxRtlReading + vbMsgBoxRight, "èòåú ëúåáú îàáèç"
            Me.Controls("txtStName").SetFocus
            GoTo toEnd
        End If

        If CStr(Me.Controls("txtRegion").value) <> CStr(wsTech.Range("AQ18")) Then
            If Me.txtRegion.value = "" Then
                MsgBox "ëúåá àú äùëåðä/ééùåá.", vbMsgBoxRtlReading + vbMsgBoxRight, "äæðú ëúåáú"
                Me.txtRegion.SetFocus
                GoTo toEnd
            Else
                changedBoxes = changedBoxes & "• ùëåðä/ééùåá îàáèç" & vbCrLf
            End If
        End If

        
        
        
        result = MsgBox("äàí àúä áèåç ùáøöåðê ìùîåø àú äòãëåï?" & vbNewLine & "ðòùå ùéðåééí áòøëéí äáàéí:" & vbNewLine & changedBoxes, vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "àéùåø ùéðåééí")
        
        If result = vbYes Then
        
            Set searchColumn = tblWorkers.ListColumns("îñ' ùåøä").DataBodyRange
            searchTerm = wsTech.Range("AK18").value
            
            Set foundCell = searchColumn.Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Check if the search term was found
            If Not foundCell Is Nothing Then
                ' Get the row number of the found cell within the table
                Dim rowNum As Long
                rowNum = foundCell.row - tblWorkers.HeaderRowRange.row
                
                turnOffStuff True
                ProtectSheet ThisWorkbook.Sheets("ôøèé òåáãéí"), False
                ' Edit the values in the found row
                With tblWorkers.ListRows(rowNum).Range
                    .Cells(5).value = Me.txtTelNum.value
                    .Cells(6).value = Me.txtEmail.value
                    .Cells(8).value = Me.txtRegion.value
                    .Cells(9).value = status
                    .Cells(16).value = medic
                    .Cells(19).value = cert
                    .Cells(20).value = allowed
                    .Cells(21).value = Me.cmbRoll.value
                    .Cells(38).value = Me.txtStName.value
                    .Cells(39).value = Me.txtBuildingNum.value
                End With
                
                ProtectSheet ThisWorkbook.Sheets("ôøèé òåáãéí"), True
                turnOffStuff False
            Else

                MsgBox "ùâéàä." & vbNewLine & "ðøàä ùîñôø äùåøä ìòøéëä - ìà ÷ééí áèáìä", vbMsgBoxRtlReading + vbMsgBoxRight, "òøéëú îéãò"
            End If
        
            MsgBox "äùéðåééí ðøùîå áäöìçä" & vbNewLine & "", vbMsgBoxRtlReading + vbMsgBoxRight, ""
        
        
        
        
        
        Else
            GoTo toEnd
        End If


       






    Else
        MsgBox "áçø ùí îàáèç îäøùéîä äðôúçú.", vbMsgBoxRtlReading + vbMsgBoxRight, "çñø îéãò"
        Me.cmbName.SetFocus
    End If

toEnd:

End Sub
Private Sub btnDelChange_Click()
    
    Me.Controls("cmbName").Locked = False
    Me.Controls("cmbName").BackColor = RGB_WHITE
    ClearUserFormInputs Me

End Sub
