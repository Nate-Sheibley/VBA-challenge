Attribute VB_Name = "Module1"
Sub singlesheet()
    Dim header1 As Variant
    header1 = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")
    Dim header2 As Variant
    header2 = Array("Ticker", "Value")
    Dim summary_labels As Variant
    summary_labels = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")

    Dim current_ticker As String
    current_ticker = "Not a ticker"
    Dim next_ticker As String
    Dim opening_val As Single
    Dim closing_val As Single
    Dim ticker_num As Integer
    ticker_num = 1
    
    Dim total_volume As LongLong
    total_volume = 0
    Dim volume As LongLong
    volume = 0
    
    Dim last_row As Long
    last_row = Range("A" & Rows.Count).End(xlUp).Row
    
    'in real data 1 quarter per sheet, quartly check not necessary to handle dates
    Debug.Print "Setting headers"
    'create new column headers
    Range("I1:L1") = header1
    Range("P1:Q1") = header2
    'create summary row labels
    Range("O2") = summary_labels(0)
    Range("O3") = summary_labels(1)
    Range("O4") = summary_labels(2)
            
    Debug.Print "Setting formatting"
    'conditional fo[srmat quarterly change column
    Dim cond1 As FormatCondition
    Dim cond_format_range As Range
    Set cond_format_range = Range("J1").EntireColumn
    'delete existing formatting
    cond_format_range.FormatConditions.Delete
    'greater than 0 red, less than 0 green
    Set cond1 = cond_format_range.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        cond1.Interior.Color = RGB(0, 255, 0)
    Set cond1 = cond_format_range.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        cond1.Interior.Color = RGB(255, 0, 0)
        
    Dim percent_format As String
    percent_format = "0.00%"
    'format percent change column as percent
    Range("K1").EntireColumn.NumberFormat = percent_format
    'format summary statistics
    Range("Q2:Q3").NumberFormat = percent_format
    'Clean formatting on first row
    Range("A1").EntireRow.FormatConditions.Delete
    
    Debug.Print "Entering for loop"
    'set opening value for first ticker
    opening_val = Cells(2, 3).Value
    
    For Row = 2 To last_row
        current_ticker = Cells(Row, 1).Value
        next_ticker = Cells(Row + 1, 1).Value
        'Debug.Print Cells(Row, 7).Value
        volume = Cells(Row, 7).Value
        total_volume = total_volume + volume
        
        If current_ticker <> next_ticker Then
            'increment ticker count at start because must start of second row
            ticker_num = ticker_num + 1
            'order of operations
            'set closing val to current row closing val
            closing_val = Cells(Row, 6).Value
            'calculate quarterly change, perecnt chance
            'write summary row (ticker, quarterly change, percent chance, total vol
            Cells(ticker_num, 9).Value = current_ticker
            'quarterly change this quarter
            Cells(ticker_num, 10).Value = WorksheetFunction.Round(closing_val - opening_val, 2)
            'percent change this quarter
            Cells(ticker_num, 11).Value = closing_val / opening_val - 1
            'total volume traded this quarter
            Cells(ticker_num, 12).Value = total_volume
            total_volume = 0
            'set next_ticker opening value to opening_val
            opening_val = Cells(Row + 1, 3).Value
        End If
        Next Row
    
    Debug.Print "Finished for loop"
    'define summary values to be held
    Dim greatest_perc_inc As Single
    Dim greatest_perc_dec As Single
    Dim greatest_total_vol As LongLong
    'find summary stats
    greatest_perc_inc = WorksheetFunction.Max(Range("k1").EntireColumn)
    greatest_perc_dec = WorksheetFunction.Min(Range("k1").EntireColumn)
    greatest_total_vol = WorksheetFunction.Max(Range("L1").EntireColumn)
    'find and set tickers corresponding to summary stats, from unique ticker column
    Range("P2") = Cells(Application.Match(greatest_perc_inc, Range("K1").EntireColumn, 0), 9).Value
    Range("P3") = Cells(Application.Match(greatest_perc_dec, Range("K1").EntireColumn, 0), 9).Value
    Range("P4") = Cells(Application.Match(greatest_total_vol, Range("L1").EntireColumn, 0), 9).Value
    'set summary stat values
    Range("Q2") = greatest_perc_inc
    Range("Q3") = greatest_perc_dec
    Range("Q4") = greatest_total_vol
End Sub

Sub runSheetsLoop()
Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
            ws.Select
            Debug.Print "Starting worksheet: " & ws.Name
            Call singlesheet
            Debug.Print "Finished worksheet: " & ws.Name
        Next ws
End Sub
