Attribute VB_Name = "Module1"
Sub RunOnAllSheets()

    Dim ws As Worksheet
    Dim lastRow As Long
'Run on each worksheet from ChatGPT
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the worksheet name is Q1, Q2, Q3, or Q4
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            ' Find the last row in column K
            lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            
            ' Call the Test_Ticker function
            Test_Ticker ws
        End If
    Next ws

End Sub

Sub Test_Ticker(ByRef ws As Worksheet)

    ' VARIABLE DEFINITIONS
    Dim i As Long
    Dim tick_sum As LongLong
    Dim open_price As Double
    Dim close_price As Double
    Dim symbol As String
    Dim summary_table_row As Integer
    Dim open_value As Double
    Dim close_value As Double
    Dim greatest_total_volume As LongLong
    Dim greatest_total_ticker As String
    Dim qpc As Double
    Dim qc As Double
    Dim r As Double
    Dim s As Double
    Dim greatest_pincrease_volume As Double
    Dim greatest_pincrease_ticker As String
    Dim greatest_pdecrease_volume As Double
    Dim greatest_pdecrease_ticker As String
    Dim cntall As Long

    ' CREATE HEADERS FOR SUMMARY TABLE
    With ws
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Quarterly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
        .Range("O2").Value = "Greatest % increase"
        .Range("O3").Value = "Greatest % decrease"
        .Range("O4").Value = "Greatest Total Volume"
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
    End With

    ' CALCULATE TICK_SUM FOR EACH TICKER SYMBOL
    summary_table_row = 2
    tick_sum = 0
    cntall = WorksheetFunction.CountA(ws.Range("A:A")) - 1

    ' Loop through test data
    For i = 2 To cntall
        ' Check if this is the first occurrence of a new value in Column A. Code assist from AskBCS Learning Assist
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            open_value = ws.Cells(i, 3).Value
        End If
'Assist from AskBCS Learning Assist for errors in this section
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Populate Summary table
            symbol = ws.Cells(i, 1).Value
            tick_sum = tick_sum + ws.Cells(i, 7).Value
            ws.Cells(summary_table_row, 9).Value = symbol
            ws.Cells(summary_table_row, 12).Value = tick_sum
            close_value = ws.Cells(i, 6).Value

            ' Calculate quarterly change
            qc = close_value - open_value
            ws.Cells(summary_table_row, 10).Value = qc

            ' Calculate quarterly percent change
            qpc = ((close_value - open_value) / open_value) * 100
            ws.Cells(summary_table_row, 11).Value = qpc

            ' Move to the next summary row
            summary_table_row = summary_table_row + 1
            tick_sum = 0
        Else
            tick_sum = tick_sum + ws.Cells(i, 7).Value
        End If
    Next i

    ' FIND GREATEST VALUES
    Dim maxvalue As Double
    Dim lastRow As Long
    Dim maxrow As Long
    maxvalue = -2000

    ' Find greatest % increase
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    For r = 2 To lastRow
        If ws.Cells(r, 11).Value > maxvalue Then
            maxvalue = ws.Cells(r, 11).Value
            maxrow = r
            greatest_pincrease_ticker = ws.Cells(r, 9).Value
        End If
    Next r

    ws.Cells(2, 17).Value = maxvalue
    ws.Cells(2, 16).Value = greatest_pincrease_ticker

    ' Highlight quarterly change based on positive or negative value
    ' Source for code https://www.youtube.com/watch?v=mqRVIL0qQu8&t=455s&ab_channel=DataNik
    
    Dim qc_range As Range
    Set qc_range = ws.Range("J2:J" & lastRow)
    
    
    For Each Cell In qc_range
        If Cell.Value > 0 Then
            Cell.Interior.ColorIndex = 4
        Else
            Cell.Interior.ColorIndex = 3
        End If
    Next

    ' Highlight percent change
    Dim pc_range As Range
    Set pc_range = ws.Range("K2:K" & lastRow)
    For Each Cell In pc_range
        If Cell.Value > 0 Then
            Cell.Interior.ColorIndex = 4
        Else
            Cell.Interior.ColorIndex = 3
        End If
    Next

    ' Find greatest % decrease
    Dim minvalue As Double
    minvalue = 0
    Dim minrow As Long
    For s = 2 To lastRow
        If ws.Cells(s, 11).Value < minvalue Then
            minvalue = ws.Cells(s, 11).Value
            minrow = s
            greatest_pdecrease_ticker = ws.Cells(s, 9).Value
        End If
    Next s

    ws.Cells(3, 17).Value = minvalue
    ws.Cells(3, 16).Value = greatest_pdecrease_ticker

    ' Find greatest volume
    Dim Tvolume As LongLong
    Tvolume = 0
    For t = 2 To lastRow
        If ws.Cells(t, 12).Value > Tvolume Then
            Tvolume = ws.Cells(t, 12).Value
            greatest_total_ticker = ws.Cells(t, 9).Value
        End If
    Next t

    ws.Cells(4, 17).Value = Tvolume
    ws.Cells(4, 16).Value = greatest_total_ticker

End Sub

