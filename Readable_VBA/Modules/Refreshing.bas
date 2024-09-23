Attribute VB_Name = "Refreshing"
Option Explicit

Public Sub refreshFilterCellStyle(sheetName As String, rangeFilter As String)

    ProtectSheet Worksheets(sheetName), False

    ' Style the line of the filter function, to be readable
    Dim FILTER_CELL As Range
    Dim leftNeighbor As Range
    Dim downNeighbor As Range
    
    Dim hasDate As Boolean
    Dim hasLeftNeighbor As Boolean
    Dim hasDownNeighbor As Boolean
    
    Set FILTER_CELL = Worksheets(sheetName).Range(rangeFilter) ' the cell of the Filter function
    Set leftNeighbor = FILTER_CELL.Offset(0, 1)
    Set downNeighbor = FILTER_CELL.Offset(1, 0)
    
    hasDate = IsDate(FILTER_CELL.value)
    hasLeftNeighbor = Not IsEmpty(leftNeighbor)
    hasDownNeighbor = Not IsEmpty(downNeighbor)
    
    ' Clean previous style
    With FILTER_CELL.Resize(Rows.count - FILTER_CELL.row + 1, Columns.count - FILTER_CELL.Column + 1)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Rows.AutoFit
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    ' Apply style
    If hasDate And hasLeftNeighbor Then
        Dim rngApplyStyle As Range
        Dim workingRange As Range
        'Dim downRange As Range
        
        'Set rngApplyStyle = FILTER_CELL
        
        Set workingRange = FILTER_CELL.End(xlToRight)
        
        Set workingRange = Range(workingRange, workingRange.End(xlToLeft))
    
        If hasDownNeighbor Then
            Set workingRange = Range(workingRange, workingRange.End(xlDown))
    
        End If
        
        'If Not leftRange.Cells Is Nothing Then
        ' Apply style to cells to the left
        With workingRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
            .Borders(xlInsideHorizontal).TintAndShade = 0
            .Borders(xlInsideHorizontal).Weight = xlHairline
            .Rows.AutoFit
        End With
    End If

    Set FILTER_CELL = Nothing
    Set leftNeighbor = Nothing
    Set downNeighbor = Nothing
    
    ProtectSheet Worksheets(sheetName), True

End Sub
Public Sub UpdatePivotsAndChartsInSheet(sheetName As String)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim chrt As ChartObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "הגיליון '" & sheetName & "' לא נמצא.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through all pivot tables in the sheet
    For Each pt In ws.pivotTables
        pt.RefreshTable
    Next pt
    
    ' Loop through all chart objects in the sheet
    For Each chrt In ws.ChartObjects
        chrt.Chart.Refresh
    Next chrt
    
    Set ws = Nothing
    
End Sub

' A sub for updateing the UI in "עובד - ____" sheets.
' To be called with the apropriate parameters.
Public Sub RefreshEpmlSheet(sheetName As String, sheetSource As String, rangeName As String, _
     tableSource As String, columnFilter As String, pivotTableNames As Variant, chartNames As Variant)

    ' no turnOffStuff because it's being called from another code

    ' Force the cell with the name to be updated  --- is it still needed?
'    Dim formulaStr As String
'    formulaStr = Worksheets(sheetSource).Range(rangeName).Formula
'    Worksheets(sheetSource).Range(rangeName).value = formulaStr
     
    Dim xTable As ListObject
    Dim xChart As ChartObject
    Dim xNameRange As Range
    Dim xCellValue As String
    Dim xChartName As String
    Dim xMatchIndex As Variant
    Dim chName As Variant
    Dim ptName As Variant

    Set xTable = Worksheets(sheetSource).ListObjects(tableSource)
    ' Set xChart = Worksheets(sheetName).ChartObjects(chartNames(0)) ' Assuming only one chart is handled in the example
    
    Set xNameRange = xTable.ListColumns(columnFilter).DataBodyRange
    
    ' Hide chart if there is no data for the given name (instead of letting it show unfiltered data).
    xCellValue = Worksheets(sheetName).Range(rangeName).value
    xMatchIndex = Application.Match(xCellValue, xNameRange, 0)
    
    If IsError(xMatchIndex) Then
        For Each chName In chartNames
            Set xChart = Worksheets(sheetName).ChartObjects(chName)
            xChartName = xChart.name
            Worksheets(sheetName).ChartObjects(xChartName).Visible = False
            Set xChart = Nothing
        Next chName
        GoTo toEnd ' No need to filter pivot table, there is no data for that name.
    Else
        For Each chName In chartNames
            Set xChart = Worksheets(sheetName).ChartObjects(chName)
            xChartName = xChart.name
            Worksheets(sheetName).ChartObjects(xChartName).Visible = True
            Set xChart = Nothing
        Next chName
    End If
    
    ' Filter pivot table
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    Dim xStr As String
    
   
    
    For Each ptName In pivotTableNames
        Set xPTable = Worksheets(sheetName).pivotTables(ptName)
        Set xPFile = xPTable.PivotFields(columnFilter)
        xStr = xCellValue
        xPFile.ClearAllFilters
        xPFile.CurrentPage = xStr
        If xPFile.CurrentPage = "(All)" Then xPFile.CurrentPage = "No Data Found"
        Set xPTable = Nothing
        Set xPFile = Nothing
    Next ptName
    
    
    

    
    UpdatePivotsAndChartsInSheet sheetName
    
toEnd:
    ' ThisWorkbook.ActiveSheet.Range("A1").Select
    Set xTable = Nothing
    Set xNameRange = Nothing



End Sub

' to be deleted dddddddddddddddddddddddddd

Public Sub refreshPrintingSheet()
    ' no turnOffStuff because it's being called from another code
    

    
    ' Style the line of the filter function, to be readable
    ' Define the range where styling will be applied
    Dim printingSheet As Worksheet
    Dim FILTER_CELL As Range
    Dim leftNeighbor As Range
    Dim downNeighbor As Range
    
    Dim hasDate As Boolean
    Dim hasLeftNeighbor As Boolean
    Dim hasDownNeighbor As Boolean
    
    Set printingSheet = ThisWorkbook.Sheets("הדפסה לשיחת משמעת")
    Set FILTER_CELL = printingSheet.Range("A11") ' the cell of the Filter function
    Set leftNeighbor = FILTER_CELL.Offset(0, 1)
    Set downNeighbor = FILTER_CELL.Offset(1, 0)
    
    hasDate = IsDate(FILTER_CELL.value)
    hasLeftNeighbor = Not IsEmpty(leftNeighbor)
    hasDownNeighbor = Not IsEmpty(downNeighbor)
    
    ' Clean previous style
    With FILTER_CELL.Resize(Rows.count - FILTER_CELL.row + 1, Columns.count - FILTER_CELL.Column + 1)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Rows.AutoFit
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Apply style
    If hasDate And hasLeftNeighbor Then
        Dim rngApplyStyle As Range
        Dim workingRange As Range
        'Dim downRange As Range
        
        'Set rngApplyStyle = FILTER_CELL
        
        Set workingRange = FILTER_CELL.End(xlToRight)
        
        Set workingRange = Range(workingRange, workingRange.End(xlToLeft))
    
        If hasDownNeighbor Then
            Set workingRange = Range(workingRange, workingRange.End(xlDown))
    
        End If
        
        'If Not leftRange.Cells Is Nothing Then
        ' Apply style to cells to the left
        With workingRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
            .Borders(xlInsideHorizontal).TintAndShade = 0
            .Borders(xlInsideHorizontal).Weight = xlHairline
            .Rows.AutoFit
        End With
    End If

    Set printingSheet = Nothing
    Set FILTER_CELL = Nothing
    Set leftNeighbor = Nothing
    Set downNeighbor = Nothing
    Set workingRange = Nothing
    Set workingRange = Nothing

End Sub


Sub refreshSheet(sheetName As String)

    Debug.Print "refreshSheet " & sheetName
    Dim updateSheet As String
    Dim nameRange As String
    Dim filterRange As String
    Dim filterColumn As String
       
    Dim pivotTables As Variant
    Dim chartNames As Variant
    
    Dim sourceSheet As String
    Dim sourceTable As String
    
    
    Dim shouldUpdateAndFilterPivot As Boolean
    Dim shouldUpdateFilter As Boolean
    Dim shouldOnlyRefreshPivot As Boolean
    
    shouldUpdateAndFilterPivot = False
    shouldOnlyRefreshPivot = False
    shouldUpdateFilter = False
    
    
    Select Case sheetName
        
        Case "עובד - כללי", "סיכום - בחני שטח", "סיכום - ביקורות", "סיכום - איחורים", "סיכום - תרגילים", "מסך ראשי"
        ' Sheets that we only want to refresh their pivots \ charts
            updateSheet = sheetName
            
            shouldUpdateAndFilterPivot = False
            shouldOnlyRefreshPivot = True ' <- only
            shouldUpdateFilter = False
            
        Case "עובד - איחורים"
            updateSheet = "עובד - איחורים" ' Name of the sheet to update.
            nameRange = "b2" ' Cell in updateSheet with the name of the employee (to filter the pivot by it).
            filterRange = "a14" ' Cell in updateSheet with the filter formula.
            filterColumn = "שם המאחר" ' Name of column in sourceTable for the pivot will be filtered by.
            pivotTables = Array("ptbLate01")  ' Names of each pivot table in updateSheet. ("ptbLate01", "ptbLate02").
            chartNames = Array("chartLate01")  ' Names of each pivot chart in updateSheet. ("chartLate01", "chartLate02").
            sourceSheet = "איחורים" ' Name of the sheet that pivot is based on.
            sourceTable = "tbLate" ' Name of the table in sourceSheet that pivot is based on.
            
            shouldUpdateAndFilterPivot = True
            shouldOnlyRefreshPivot = False ' <- no need in this case, will be called from shouldUpdateAndFilterPivot
            shouldUpdateFilter = True
            
        Case "עובד - ביקורות"
            updateSheet = "עובד - ביקורות" ' Name of the sheet to update.
            nameRange = "b2" ' Cell in updateSheet with the name of the employee (to filter the pivot by it).
            filterRange = "a14" ' Cell in updateSheet with the filter formula.
            filterColumn = "מאבטח" ' Name of column in sourceTable for the pivot will be filtered by.
            pivotTables = Array("ptbRvw01", "ptbRvw02")  ' Names of each pivot table in updateSheet. ("ptbLate01", "ptbLate02").
            chartNames = Array("chartRvw01", "chartRvw02")  ' Names of each pivot chart in updateSheet. ("chartLate01", "chartLate02").
            sourceSheet = "ביקורות" ' Name of the sheet that pivot is based on.
            sourceTable = "tbPerfReview" ' Name of the table in sourceSheet that pivot is based on.
            
            shouldUpdateAndFilterPivot = True
            shouldOnlyRefreshPivot = False ' <- no need in this case, will be called from shouldUpdateAndFilterPivot
            shouldUpdateFilter = True
            
        Case "עובד - תרגילים"
            updateSheet = "עובד - תרגילים" ' Name of the sheet to update.
            nameRange = "b2" ' Cell in updateSheet with the name of the employee (to filter the pivot by it).
            filterRange = "a14" ' Cell in updateSheet with the filter formula.
            filterColumn = "מאבטח" ' Name of column in sourceTable for the pivot will be filtered by.
            pivotTables = Array("ptbDrill01", "ptbDrill02")  ' Names of each pivot table in updateSheet. ("ptbLate01", "ptbLate02").
            chartNames = Array("chartDrill01", "chartDrill02")  ' Names of each pivot chart in updateSheet. ("chartLate01", "chartLate02").
            sourceSheet = "תרגילים" ' Name of the sheet that pivot is based on.
            sourceTable = "tbDrills" ' Name of the table in sourceSheet that pivot is based on.
            
            shouldUpdateAndFilterPivot = True
            shouldOnlyRefreshPivot = False ' <- no need in this case, will be called from shouldUpdateAndFilterPivot
            shouldUpdateFilter = True
            
        Case "הדפסה לשיחת משמעת"
            updateSheet = "הדפסה לשיחת משמעת"
            filterRange = "a11"
            
            shouldUpdateAndFilterPivot = False
            shouldOnlyRefreshPivot = False
            shouldUpdateFilter = True ' <- only
            
        Case Else
            Exit Sub
            
    End Select
    
    
    
'    If shouldUpdateAndFilterPivot Then
'        ' Filter pivots by worker name, hide charts if no data, refresh pivots and charts
'        RefreshEpmlSheet updateSheet, sourceSheet, nameRange, sourceTable, filterColumn, pivotTables, chartNames
'    End If
'
'    If shouldOnlyRefreshPivot Then
'
'    End If

    If shouldUpdateFilter Then
        ' Update the style for the filter lines
        refreshFilterCellStyle updateSheet, filterRange
    End If
    
    
    
    
 End Sub
