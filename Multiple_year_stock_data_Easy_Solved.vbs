Sub loop_through_all_worksheets()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

Dim lRow As Long
Dim lCol As Long
Dim CountRow As Long

For Each ws In ThisWorkbook.Worksheets
    
    ws.Activate
'do whatever you need

    ws.Range("I:N").Delete
    ws.Cells(1, 9) = ws.Cells(1, 1)
    ws.Cells(1, 10) = "Total Stock Volume"
  
    CountRow = 1

'    Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row

'    Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column

    For i = 2 To lRow                           'When is different, sum column 10 + 7, and correct ticker name'
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(CountRow + 1, 10) = ws.Cells(CountRow + 1, 10).Value + ws.Cells(i, 7).Value
                    ws.Cells(CountRow + 1, 9) = ws.Cells(i, 1).Value
                    CountRow = 1 + CountRow
                Else                            'This is the sum of column 10 + column 7
                    ws.Cells(CountRow + 1, 10) = ws.Cells(CountRow + 1, 10).Value + ws.Cells(i, 7).Value
                    ws.Cells(CountRow + 1, 9) = ws.Cells(i, 1).Value
                End If
    Next i

ws.Range("I:N").Columns.AutoFit                 'To format a couple of columns'

Next

starting_ws.Activate 'activate the worksheet that was originally active

MsgBox ("Fix complete!")
End Sub




