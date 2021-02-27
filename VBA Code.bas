Attribute VB_Name = "Module1"
Sub Stock():
'iterate through the worksheets
For Each ws In Worksheets
'Set variables
Dim ticker As Long
Dim yearly_change As Double
Dim opening As Double
Dim closing As Double
Dim percent_change As Double
Dim table As Long
Dim LastRow As Long
ticker = 1
table = 2
opening = ws.Cells(2, 3).Value
total_volume = 0
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'create table headings
ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"
For i = 2 To LastRow:
    If ws.Cells(i + 1, ticker).Value <> ws.Cells(i, ticker).Value Then
        'add ticker symbol to table
        ws.Cells(table, 9).Value = ws.Cells(i, ticker).Value
        'find yearly change
        closing = ws.Cells(i, 6)
        yearly_change = closing - opening
        ws.Cells(table, 10).Value = yearly_change
        'conditional formatting yearly change color
            If ws.Cells(table, 10).Value > 0 Or ws.Cells(table, 10).Value = 0 Then
                ws.Cells(table, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(table, 10).Value < 0 Then
                ws.Cells(table, 10).Interior.ColorIndex = 3
            End If
        'find % change
        percent_change = (yearly_change / opening) * 100
            'bypass divide by zero error
            If opening = 0 And yearly_change = 0 Then
                percent_change = 0
                ws.Cells(table, 11).Value = 0
            ElseIf opening = 0 Then
                percent_change = 100
                ws.Cells(table, 11).Value = 100
            ElseIf ws.Cells(table, 11).Value <> 0 Then
                ws.Cells(table, 11).Value = percent_change
            End If
        'find total volume
        ws.Cells(table, 12).Value = total_volume + ws.Cells(i, 7)
        'move down one row in new table
        table = table + 1
        'reset cell value for opening
        opening = ws.Cells(i + 1, 3).Value
        'reset total volume
        total_volume = 0
    ElseIf ws.Cells(i + 1, ticker).Value = ws.Cells(i, ticker).Value Then
        total_volume = total_volume + ws.Cells(i, 7).Value
    End If
'find greatest_increase and put in table
    If ws.Cells(table, 11).Value > 0 And ws.Cells(table, 11).Value > greatest_increase Then
        ws.Cells(2, 16).Value = ws.Cells(table, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(table, 11).Value
    'cast new value to greatest_increase
        greatest_increase = ws.Cells(table, 11).Value
'find greatest_decrease and put in table
    ElseIf ws.Cells(table, 11).Value < 0 And ws.Cells(table, 11).Value < greatest_decrease Then
        ws.Cells(3, 16).Value = ws.Cells(table, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(table, 11).Value
    'cast new value to greatest_decrease
        greatest_decrease = ws.Cells(table, 11).Value
    End If
        'find greatest_volume and put in table
    If ws.Cells(table, 12).Value > greatest_volume Then
        ws.Cells(4, 16).Value = ws.Cells(table, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(table, 12).Value
        'cast new value to greatest_volume
        greatest_volume = ws.Cells(table, 12).Value
    End If
Next i
Next ws
End Sub
Sub Greatest():
'iterate through the worksheets
For Each ws In Worksheets
'Set variables
Dim ticker_symbol As Integer
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim table As Long
Dim LastRow2 As Long
table = 2
ticker_symbol = 9
greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0
LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
'set headings
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
'find greatest_increase and put in table
For i = 2 To LastRow2
    If ws.Cells(i, 11).Value > 0 And ws.Cells(i, 11).Value > greatest_increase Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    'cast new value to greatest_increase
        greatest_increase = ws.Cells(i, 11).Value
'find greatest_decrease and put in table
    ElseIf ws.Cells(i, 11).Value < 0 And ws.Cells(i, 11).Value < greatest_decrease Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    'cast new value to greatest_decrease
        greatest_decrease = ws.Cells(i, 11).Value
    End If
        'find greatest_volume and put in table
    If ws.Cells(i, 12).Value > greatest_volume Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        'cast new value to greatest_volume
        greatest_volume = ws.Cells(i, 12).Value
    End If
Next i
Next ws
End Sub
