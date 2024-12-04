Attribute VB_Name = "basStock"
Option Explicit

Type StockReport
    Ticker As String
    RecordCount As Long
    QtrOpen As Double
    QtrClose As Double
    QtrChange As Currency
    PctChange As Double
    TotStockVol As Double
End Type

Sub WrapperExecuteStockAnalysis()

    Dim ws As Excel.Worksheet
    Dim tmrStart As Single, tmrEnd As Single
    Dim strMsg As String
    
    'Start timer
    tmrStart = Timer
    
    'Hide screen updates
    Application.ScreenUpdating = False
    
    DeleteNewReportColumns
    
    
    'Set ws = Worksheets("A")
    
    'Loop through worksheets in workbook
    For Each ws In ActiveWorkbook.Worksheets
        AnalyzeStockReport ws
    Next ws
    
    'Enable screen updates
    Application.ScreenUpdating = True
    
    'Cleanup object
    Set ws = Nothing
    
    tmrEnd = Timer
    
    strMsg = "Completed quarterly stock report analysis." & vbCrLf & vbCrLf & _
             "Elapsed time: " & CStr(Round(tmrEnd - tmrStart, 2)) & " seconds"
    
    MsgBox strMsg, vbOKOnly + vbInformation, "Status"
    
    
       
    
End Sub

Sub AnalyzeStockReport(ByRef refWS As Excel.Worksheet)

    Dim rngSrc As Excel.Range, rng As Excel.Range
    Dim cond1 As Excel.FormatCondition, cond2 As Excel.FormatCondition
    
    Dim StockReportArray() As StockReport
    Dim lngRow As Long, lngLastRow As Long, lngStockCount As Long, lngRecCount As Long
    Dim i As Long
    Dim strTicker As String, strFormula As String
    
    lngLastRow = refWS.Cells(Rows.Count, 1).End(xlUp).Row
    lngStockCount = 1
    
    '----------------------------------------------------------------------------
    ' Initialize array to store stock analysis results
    '----------------------------------------------------------------------------
    ReDim Preserve StockReportArray(1 To 1)
    
    'Get initial ticker symboland quarter open for very first record
    With StockReportArray(lngStockCount)
        .Ticker = refWS.Cells(2, 1).Value
        .QtrOpen = refWS.Cells(2, 3).Value
    End With
    
    For lngRow = 2 To lngLastRow
    
        'lngRecCount = lngRecCount + 1

        'New Ticker Record in set
        If (StockReportArray(lngStockCount).Ticker <> refWS.Cells(lngRow, 1).Value) Then
            
            With StockReportArray(lngStockCount)
                
                .RecordCount = lngRecCount
                
                'Look at previous row to capture quarter close
                .QtrClose = refWS.Cells(lngRow - 1, 6).Value

                
                'Calculate quarter change
                .QtrChange = .QtrOpen - .QtrClose
                
                'Calculate percent change
                .PctChange = Round(.QtrChange / .QtrOpen, 4)
                
            End With

            
            'increment ticker symbol count
            lngStockCount = lngStockCount + 1
                        
            
            'Reinitialize array
            ReDim Preserve StockReportArray(1 To lngStockCount)
            
            'Retrieve new record and store array
            With StockReportArray(lngStockCount)
                .Ticker = refWS.Cells(lngRow, 1).Value
                .QtrOpen = refWS.Cells(lngRow, 3).Value
                .TotStockVol = refWS.Cells(lngRow, 7).Value
            End With
            
            lngRecCount = 0
            
        Else
            
            With StockReportArray(lngStockCount)

                .TotStockVol = .TotStockVol + refWS.Cells(lngRow, 7).Value
                
                
                'On last record get current qtr close and add to totl stock volume
                If (lngRow = lngLastRow) Then
        
                    .RecordCount = lngRecCount + 1
                    
                    .QtrClose = refWS.Cells(lngRow, 6).Value
                    '.TotStockVol = .TotStockVol + refWS.Cells(lngRow, 7).Value
                
                    'Calculate quarter change
                    .QtrChange = .QtrOpen - .QtrClose
                    
                    'Calculate percent change
                    .PctChange = Round(.QtrChange / .QtrOpen, 4)
                        
                End If
        
            End With
            

        End If
        
        lngRecCount = lngRecCount + 1
        
    Next lngRow

    ExportArray StockReportArray(), "C:\Debug\stock_" & refWS.Name & ".txt"
    
    
    '----------------------------------------------
    'Populate array results to current worksheet
    '----------------------------------------------
    'add column heading
    With refWS
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Quarterly Change"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Total Stock Volume"
    End With
    
    'Iterate through array and populate worksheet
    'Analysis report will start at column I (Index=9)
    For i = 1 To UBound(StockReportArray)
        
        With StockReportArray(i)
            
            refWS.Cells(i + 1, 9).Value = .Ticker
            refWS.Cells(i + 1, 10).Value = .QtrChange
            refWS.Cells(i + 1, 11).Value = .PctChange
            refWS.Cells(i + 1, 12).Value = .TotStockVol
            
        End With
        
    Next i
    
    
    With refWS
    
        'Add 1 to ticker count for coorect row position in range on sheet
        lngStockCount = lngStockCount + 1
        
        
        
        '----------------------------------------------------------------------------
        'Quarterly Change Conditional Formating
        '----------------------------------------------------------------------------
        Set rng = .Range("J2:J" & lngStockCount)
        
        Set cond1 = rng.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set cond2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
        
        cond1.Interior.Color = vbGreen
        cond2.Interior.Color = vbRed
        
        'Format Percent Change
        Set rng = .Range("K2:K" & lngStockCount)
        rng.NumberFormat = "0.00%"
        
        
        'Format total Stock Volume
        Set rng = .Range("L2:L" & lngStockCount)
        rng.NumberFormat = "#,##0_);[Red](#,##0)"
        
        
        
        '----------------------------------------------------------------------------
        ' Greatest % Increase / Greatest % Decrease / Greatest Total Volume
        '----------------------------------------------------------------------------
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        
        'Create formula using Match to return ticker symbol from column I
        .Range("P2").Formula = "=INDEX(I2:I" & lngStockCount & ",MATCH(MAX(K2:K" & lngStockCount & "),K2:K" & lngStockCount & ",0))"
        .Range("P3").Formula = "=INDEX(I2:I" & lngStockCount & ",MATCH(MIN(K2:K" & lngStockCount & "),K2:K" & lngStockCount & ",0))"
        .Range("P4").Formula = "=INDEX(I2:I" & lngStockCount & ",MATCH(MAX(L2:L" & lngStockCount & "),L2:L" & lngStockCount & ",0))"
        
        .Range("Q2").Formula = "=MAX(K2:K" & lngStockCount & ")"
        .Range("Q3").Formula = "=MIN(K2:K" & lngStockCount & ")"
        .Range("Q4").Formula = "=MAX(L2:L" & lngStockCount & ")"
    
    End With
    
    '----------------------------------------------------------------------------
    'Format Headers
    '----------------------------------------------------------------------------
    refWS.Activate
    
    'bold first row
    Range("1:1").Font.Bold = True
    
    'freeze top row pane
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'autofit columns
    ActiveSheet.UsedRange.Columns.AutoFit
    
    '----------------------------------------------------------------------------
    ' Cleanup
    '----------------------------------------------------------------------------
    Set cond1 = Nothing
    Set cond2 = Nothing
    Set rngSrc = Nothing
    
    'Debug.Print refWS.Name
    
    
End Sub


Sub ExportArray( _
    ByRef refArray() As StockReport, _
    ByVal vstrFile As String)

    Dim hFile As Integer
    Dim strFile As String, strLine As String
    Dim i As Long

    hFile = FreeFile
    vstrFile = "C:\Users\Daniel\Desktop\Module_2\Resources\output.txt"
    Open vstrFile For Output As #hFile
    
    For i = 1 To UBound(refArray)
        With refArray(i)
            strLine = .Ticker & "|" & _
                      .QtrOpen & "|" & _
                      .QtrClose & "|" & _
                      .QtrChange & "|" & _
                      .PctChange & "|" & _
                      .TotStockVol
        End With
        
        Print #hFile, strLine
        
    Next i
    Close #hFile

End Sub


Sub DeleteNewReportColumns()

    Dim ws As Excel.Worksheet
    
    Application.ScreenUpdating = False
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Columns("I:Q").Delete

        ws.Activate
        ActiveSheet.Cells(2, 1).Select
    Next ws
    
    ActiveWorkbook.Worksheets(1).Activate
    ActiveSheet.Cells(2, 1).Select
    
    
    Application.ScreenUpdating = True
    
    Set ws = Nothing
    
End Sub


