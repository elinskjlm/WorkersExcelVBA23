Option Explicit

Private Sub Workbook_Open()
    
    turnOffStuff True
    
    ThisWorkbook.Sheets("מסך ראשי").Select
    Range("a1").Select
    
    Dim shp As Shape
    Set shp = ThisWorkbook.Sheets("מסך ראשי").Shapes("אליפסה 6")
    Dim GREEN As Long
    GREEN = RGB(50, 205, 50)
    
    shp.Fill.ForeColor.RGB = GREEN
    
    ToggleSheetVisibility

    Application.EnableEvents = True
    Set shp = Nothing
    turnOffStuff False
    
End Sub