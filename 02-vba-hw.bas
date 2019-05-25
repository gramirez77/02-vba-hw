' NAME
'    AnalyzeStocks() -- stock market data analyzer
' 
' DESCRIPTION
'    For each worksheet in the workbook, AnalyzeStocks() crunches stock market
'    data in columns A to G and produces two summary tables: one per ticker
'    (columns I to L), and one per year (columns O to Q).
' 
' AUTHOR
'    Gilberto Ramirez (gramirez77@gmail.com)
'    v1.0: May 25, 2019
' 
' REMARKS:
'    This is the second week homework for the UNC Data Analytics Boot Camp.
'    Choice made for this homework: Hard assignment including the challenge.

Public Sub AnalyzeStocks()
    Dim ws As Worksheet
    Dim numberofTickers As Integer
    Dim previousTicker As String
    Dim currentTicker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalStock As Double
    Dim greatestPercIncreaseTicker As String
    Dim greatestPercIncreaseValue As Double
    Dim greatestPercDecreaseTicker As String
    Dim greatestPercDecreaseValue As Double
    Dim greatestTotalVolumeTicker As String
    Dim greatestTotalVolumeValue As Double
    Dim inputRow As Long
    Dim outputRow As Long
    
    ' CHALLENGE: iterate thru each Worksheet object within the current Workbook
    For Each ws In Application.ActiveWorkbook.Sheets
        
        ' most of the inits are here
        numberofTickers = 0
        previousTicker = ""
        currentTicker = ""
        openingPrice = 0.00
        closingPrice = 0.00
        totalStock = 0
        greatestPercIncreaseTicker = ""
        greatestPercIncreaseValue = 0.00
        greatestPercDecreaseTicker = ""
        greatestPercDecreaseValue = 0.00
        greatestTotalVolumeTicker = ""
        greatestTotalVolumeValue = 0
        
        Debug.Print ws.Name
        
        ' write per ticker summary table header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' row pointers set to 2 to skip the header
        inputRow = 2
        outputRow = 2
        Do
            currentTicker = ws.Cells(inputRow, 1).Value
            
            If currentTicker <> previousTicker Then
                ' ticker change detected, so time for action!
                
                ' show progress in status bar since this function might take long
                ' progress updates happen every 50 ticker changes
                If numberofTickers Mod 50 = 0 Then
                    Application.StatusBar = numberofTickers & " tickers found in sheet """ & ws.Name & """..."
                    DoEvents
                End If
                
                If previousTicker <> "" Then
                    
                    numberofTickers = numberofTickers + 1

                     ' closingPrice must be fetched from the previous row
                    closingPrice = ws.Cells(inputRow - 1, 6).Value
                    
                    ' per ticker summary table | "Ticker" field
                    ws.Cells(outputRow, 9).Value = previousTicker
                                   
                    ' per ticker summary table | "Yearly Change" field
                    ws.Cells(outputRow, 10).Value = closingPrice - openingPrice
                    
                    ' per ticker summary table | "Percent Change" field
                    If openingPrice > 0 Then
                        ws.Cells(outputRow, 11).Value = (closingPrice - openingPrice) / openingPrice
                    ElseIf openingPrice = 0 Then
                        ' if openingPrice is 0, "Percent Change" is set to an empty string
                        ' having a number when openingPrice is 0, it does not make sense!
                        ws.Cells(outputRow, 11).Value = ""
                    End If
                    
                     ' per ticker summary table | "Total Stock Volume" field
                    ws.Cells(outputRow, 12).Value = totalStock
                    
                    outputRow = outputRow + 1
                End If
                
                openingPrice = ws.Cells(inputRow, 3).Value
                totalStock =0
                previousTicker = currentTicker
                
            End If
            
            totalStock = totalStock + ws.Cells(inputRow, 7).Value
            inputRow = inputRow + 1
            
        Loop While currentTicker <> ""
        
        ' HARD: per year summary table
        Application.StatusBar = "Calculating the hard challenge portion now...": DoEvents
        ' write row headers for the output
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ' write column headers for the output
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ' iterate thru the per ticker table to calculate the summary results
        inputRow = 2
        greatestPercIncreaseValue = ws.Cells(inputRow, 11)
        greatestPercDecreaseValue = ws.Cells(inputRow, 11)
        greatestTotalVolumeValue = ws.Cells(inputRow, 12)
        Do
            currentTicker = ws.Cells(inputRow, 9)
            If ws.Cells(inputRow, 11) > greatestPercIncreaseValue Then
                greatestPercIncreaseValue = ws.Cells(inputRow, 11)
                greatestPercIncreaseTicker = currentTicker
            End If
            If ws.Cells(inputRow, 11) < greatestPercDecreaseValue Then
                greatestPercDecreaseValue = ws.Cells(inputRow, 11)
                greatestPercDecreaseTicker = currentTicker
            End If
             If ws.Cells(inputRow, 12) > greatestTotalVolumeValue Then
                greatestTotalVolumeValue = ws.Cells(inputRow, 12)
                greatestTotalVolumeTicker = currentTicker
            End If
            inputRow = inputRow + 1
        Loop While currentTicker <> ""
        ' time to post the summary results
        ws.Cells(2, 16).Value = greatestPercIncreaseTicker
        ws.Cells(2, 17).Value = greatestPercIncreaseValue
        ws.Cells(3, 16).Value = greatestPercDecreaseTicker
        ws.Cells(3, 17).Value = greatestPercDecreaseValue
        ws.Cells(4, 16).Value = greatestTotalVolumeTicker
        ws.Cells(4, 17).Value = greatestTotalVolumeValue
        
        ' do conditional formatting for column J (Yearly Change)
        ' * value less than 0, cell color fill set to red
        ' * value greater or equal than 0, cell color fill set to green 
        With ws.Range("J:J")
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
            .Style = "Comma"
            .NumberFormat = "#,##0.00"
        End With
        ws.Range("J1").FormatConditions.Delete
        
        'change styles and number formats for output columns
        With ws.Columns("K:K")
            .Style = "Percent"
            .NumberFormat = "#,##0.00%"
        End With
        With ws.Columns("L:L")
            .Style = "Comma"
            .NumberFormat = "#,##0"
        End With
        With ws.Range("Q2:Q3")
            .Style = "Percent"
            .NumberFormat = "#,##0.00%"
        End With
        With ws.Range("Q4")
            .Style = "Comma"
            .NumberFormat = "#,##0"
        End With
        
        'autofit column width after changes to make it all neat
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit
        
        ' relinquish control of status bar... it's over!
        Application.StatusBar = False
        
    Next
    
End Sub
