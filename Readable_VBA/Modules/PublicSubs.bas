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
    Set shp = ThisWorkbook.Sheets("��� ����").Shapes("������ 6")
    
    Dim RED As Long
    Dim GREEN As Long
    
    RED = RGB(220, 20, 60)
    GREEN = RGB(50, 205, 50)

    Dim visibility As XlSheetVisibility
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    

    If shp.Fill.ForeColor.RGB = RED Then
        result = MsgBox("��� ��� ��� ������� �����:" & vbNewLine & "��� / ��� / ���?", vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "����� �������")
        If result = vbYes Then
            visibility = xlSheetVisible
            shp.Fill.ForeColor.RGB = GREEN
        Else
            MsgBox "�� �� ����, �� �����, �� ����, ��� ���� ����!!!" & vbNewLine & "�� ��� ����, ��� �������. �� ���� ���!", vbMsgBoxRtlReading + vbMsgBoxRight, "�� ����!!!!!!!!"
        End If
                
    ElseIf shp.Fill.ForeColor.RGB = GREEN Then
        visibility = xlSheetHidden
        shp.Fill.ForeColor.RGB = RED
    Else
        MsgBox "���� �� ������ ���� ���� ����" & vbNewLine & "Const RED As Long = RGB(220, 20, 60)" & vbNewLine & "Const GREEN As Long = RGB(50, 205, 50)", vbMsgBoxRtlReading + vbMsgBoxRight, "�� ����"
    End If

    ' Specify the names of the sheets to toggle
    sheetNames = Array("���� ������", "���� ���", "�������", "�������", "�������", "������ ����", "ToDo", "����� ����� �����", "���� ������", "����� ����� ������", "����� ������")
        
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
    
    Set rg = ThisWorkbook.Sheets("������ ����").Range("X32")
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
    Set ws = ThisWorkbook.Worksheets("����� ����� �����")
    
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
    
    fileName = "����� ����� ����� " & employeeName & " " & todayString
    
    ' Set the path for the PDF file
    dirPath = "C:\Users\Moked Kishla\Desktop\���\���� ����� ������ 2023\������ ������ �����\"
    pdfPath = dirPath & fileName & ".pdf"
    
    ' Specify the sheet name to print
    printSheetName = "����� ����� �����"
    
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
    
    path = ThisWorkbook.Sheets("������ ����").Range("b15").value & "\"
    file = ThisWorkbook.Sheets("������ ����").Range("d17").value
    
    If file = "" Then
        MsgBox "��� ���� ����� ������" & vbNewLine & "(������ ����, ����� ���� ��� ������)", vbMsgBoxRtlReading + vbMsgBoxRight, "�����"
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
    result = MsgBox("��� ��� ���� ������� �����?", vbMsgBoxRtlReading + vbMsgBoxRight + vbQuestion + vbYesNo, "����� �����")
    
    If result = vbYes Then

        Dim whatsappURL As String
    
        ' Construct the WhatsApp URL
        whatsappURL = ThisWorkbook.Sheets("����� - ������ ������").Range("L3").value
        
        ' Open the WhatsApp URL in the default browser
        ThisWorkbook.FollowHyperlink whatsappURL
    
    End If
    
End Sub


