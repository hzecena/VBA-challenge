Sub VBA_challenge_code():
    
Dim row_count As Long
Dim total As Double
Dim start As Long
Dim change As Single
Dim pcent_change As Single
Dim ws As Worksheet
        
For Each ws In Worksheets

' Headers & Formats
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly_Change"
ws.Range("K1") = "Percent_Change"
ws.Range("L1") = "Total_Stock_Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"

'Create a script that will loop through all the stocks for one year
' Initialize values
start = 2
row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row
j = 0
change = 0
total = 0
        
For i = 2 To row_count

    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        total = total + ws.Cells(i, 7)
        If total = 0 Then
        ' Print results
        ws.Range("I" & 2 + j) = ws.Cells(i, 1)
        ws.Range("J" & 2 + j) = 0
        ws.Range("K" & 2 + j) = "%" & 0
        ws.Range("L" & 2 + j) = 0
    Else
        'Find first non zero starting value
        If ws.Cells(start, 3) = 0 Then
            For find_value = start To i
                If ws.Cells(find_value, 3) <> 0 Then
                    start = find_value
                    Exit For
                        End If
                    Next find_value
                End If
                
                change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                pcent_change = Round((change / ws.Cells(start, 3) * 100), 2)
                ' go to the next ticker
                start = i + 1
                    
                ' Print results
                ws.Range("I" & 2 + j) = ws.Cells(i, 1)
                ws.Range("J" & 2 + j) = Round(change, 2)
                ws.Range("K" & 2 + j) = "%" & pcent_change
                ws.Range("L" & 2 + j) = total
                    
                If change > 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                ElseIf change < 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End If
            End If
    
            ' reset variables for the new ticker
            total = 0
            change = 0
            j = j + 1
        Else
            total = total + ws.Cells(i, 7)
        End If
    Next i

ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & row_count)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & row_count)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & row_count))
    
' Worksheet Function called Match
incr_no = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & row_count)), ws.Range("K2:K" & row_count), 0)
decr_no = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & row_count)), ws.Range("K2:K" & row_count), 0)
vol_no = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & row_count)), ws.Range("L2:L" & row_count), 0)
        
ws.Range("P2") = ws.Cells(incr_no + 1, 9)
ws.Range("P3") = ws.Cells(decr_no + 1, 9)
ws.Range("P4") = ws.Cells(vol_no + 1, 9)

Next ws

MsgBox ("Program Complete. Enter When Ready.")

End Sub