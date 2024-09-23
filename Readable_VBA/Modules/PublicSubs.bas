Attribute VB_Name = "PublicSubs"
Option Explicit

Public Sub turnOffStuff(ByVal off As Boolean)
    
    If off = True Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
    
End Sub

Public Sub ClearUserFormInputs(ByVal uf As UserForm)
    ' no turnOffStuff because it's being called from another code
    Dim ctrl As Control
    
    For Each ctrl In uf.Controls
        Select Case TypeName(ctrl)
            Case "CheckBox", "OptionButton"
                ctrl.value = False
            Case "TextBox", "ComboBox"
                ctrl.value = ""
            Case "ListBox"
                Set ctrl.value = Nothing
        End Select
    Next ctrl
End Sub
Public Sub ToggleSheetVisibility()
    
    turnOffStuff True
    
    Dim result As Integer
        
    Dim ws As Worksheet
    Dim sheetNames() As Variant
    Dim i As Long
 
    Dim shp As Shape
    Set shp = ThisWorkbook.Sheets("מסך ראשי").Shapes("אליפסה 6")
    
    Dim RED As Long
    Dim GREEN As Long
    
    RED = RGB(220, 20, 60)
    GREEN = RGB(50, 205, 50)

    Dim visibility As XlSheetVisibility
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    

    If shp.Fill.ForeColor.RGB = RED Then
        result = MsgBox("האם אתה אחד מהאנשים הבאים:" & vbNewLine & "אבי / בני / גדי?", vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "אישור שינויים")
        If result = vbYes Then
            visibility = xlSheetVisible
            shp.Fill.ForeColor.RGB = GREEN
        Else
            MsgBox "אז אל תיגע, אל תעזור, אל תלחץ, ואל תשנה כלום!!!" & vbNewLine & "אם אתה צריך, פנה לאוליאל. לא לגעת לבד!", vbMsgBoxRtlReading + vbMsgBoxRight, "אל תיגע!!!!!!!!"
        End If
                
    ElseIf shp.Fill.ForeColor.RGB = GREEN Then
        visibility = xlSheetHidden
        shp.Fill.ForeColor.RGB = RED
    Else
        MsgBox "הצבע של הכפתור שונה והכל נדפק" & vbNewLine & "Const RED As Long = RGB(220, 20, 60)" & vbNewLine & "Const GREEN As Long = RGB(50, 205, 50)", vbMsgBoxRtlReading + vbMsgBoxRight, "לא עובד"
    End If

    ' Specify the names of the sheets to toggle
    sheetNames = Array("פרטי עובדים", "בחני שטח", "איחורים", "ביקורות", "תרגילים", "גיליון טכני", "ToDo", "הדפסה לשיחת משמעת", "מידע לגרפים", "עמדות שליטה להדפסה", "אלפון להדפסה")
        
    ' Loop through the sheet names array
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        ws.Visible = visibility
        Set ws = Nothing
    Next i
    
    currentSheet.Activate

    
    
    Set shp = Nothing
    Set currentSheet = Nothing
 
    
    turnOffStuff False
      
End Sub

Public Sub ProtectSheet(sh As Worksheet, toLock As Boolean)
    Dim password As String
    Dim rg As Range
    
    Set rg = ThisWorkbook.Sheets("גיליון טכני").Range("X32")
    password = rg.value
    
    If toLock Then
        sh.Protect password:=password
    Else
        sh.Unprotect password
    End If
End Sub



Public Sub PrintDisiplinToPDF()
      
    turnOffStuff True
    
    refreshPrintingSheet ' replace with appro. *******************
    
    Dim originalSheet As Worksheet
    Dim originalCell As Range
    Set originalSheet = ActiveSheet
    Set originalCell = ActiveCell
     
    Dim dirPath, pdfPath As String
    Dim printSheetName As String
    Dim employeeName As String
    Dim todayString As String
    Dim fileName As String

    Dim ws As Worksheet
    Dim wasHidden As Boolean
    Set ws = ThisWorkbook.Worksheets("הדפסה לשיחת משמעת")
    
    ' Check if the sheet is hidden
    If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        wasHidden = True
        ws.Visible = xlSheetVisible
    Else
        wasHidden = False
    End If
    
    employeeName = Range("NameSearchCell").value
    
    ' Get today's date as a string in "dd.mm.yy" format
    todayString = Format(Date, "dd.mm.yy")
    
    fileName = "סיכום לשיחת משמעת " & employeeName & " " & todayString
    
    ' Set the path for the PDF file
    dirPath = "C:\Users\Moked Kishla\Desktop\סגן\קובץ ניהול עובדים 2023\הדפסות לשיחות שימוע\"
    pdfPath = dirPath & fileName & ".pdf"
    
    ' Specify the sheet name to print
    printSheetName = "הדפסה לשיחת משמעת"
    
    ' Create a new PDF file
    With ThisWorkbook.Worksheets(printSheetName)
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End With
    
    
    ' Added on 10/2023: open the folder of the newly created PDF file, so noone will complain about not finding it.
    ' Check if the folder exists
    If Dir(dirPath, vbDirectory) <> "" Then
        ' Open the folder in Windows Explorer
        Shell "explorer.exe """ & dirPath & """", vbNormalFocus
    End If
    ' End of adding on 10/2023.

 
    If wasHidden Then
        ws.Visible = xlSheetHidden
    Else
        ws.Visible = xlSheetVisible
    End If
    
    originalSheet.Activate
    originalCell.Select
    
    Set originalSheet = Nothing
    Set originalCell = Nothing
    Set ws = Nothing
    

    
    turnOffStuff False

End Sub

' in work, not in use
Public Sub OpenHelpFileForEditing()
    Application.EnableEvents = False

    turnOffStuff True
    
    Dim path As String
    Dim file As String
    
    On Error GoTo 0
    
    path = ThisWorkbook.Sheets("גיליון טכני").Range("b15").value & "\"
    file = ThisWorkbook.Sheets("גיליון טכני").Range("d17").value
    
    If file = "" Then
        MsgBox "בחר קובץ במקום המתאים" & vbNewLine & "(גיליון טכני, פתיחת קובץ עזר לעריכה)", vbMsgBoxRtlReading + vbMsgBoxRight, "שגיאה"
    Else
        Workbooks.Open fileName:=path & file
    End If
    
    turnOffStuff False
    Application.EnableEvents = True
    
End Sub



Sub stopEvents()
    Application.EnableEvents = False
End Sub

Sub continueEvents()
    Application.EnableEvents = True
End Sub


Sub SendWhatsAppMessage()
    Dim result As Integer
    result = MsgBox("האם אתה בטוח שברצונך לשלוח?", vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "אישור שליחה")
    
    If result = vbYes Then

        Dim whatsappURL As String
    
        ' Construct the WhatsApp URL
        whatsappURL = ThisWorkbook.Sheets("סיכום - הצהרות בריאות").Range("L3").value
        
        ' Open the WhatsApp URL in the default browser
        ThisWorkbook.FollowHyperlink whatsappURL
    
    End If
    
End Sub


