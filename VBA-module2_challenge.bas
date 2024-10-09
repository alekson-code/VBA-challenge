Attribute VB_Name = "Module1"
Option Explicit

Sub Stock_market()

Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        Dim RowCount As Long
        Dim SummaryRow As Long
        Dim start As Long
        Dim total As Double
        Dim i As Double
        Dim j As Integer
        'Dim change As Double
        Dim increase_number As Integer
        Dim decrease_number As Integer
        Dim volume_number As Double
        Dim opening_price As Double
        Dim closing_price As Double
        Dim quarterly_change As Double
        Dim stock_percent_change
        'Dim Ticker As String
        
        
        'The proceeding lines create the column headings, as well as the headings horizontal headings used to show the greatest and least perchange change as well as the greatest total volume.
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"


        'This creates a variable which will hold the total stock volume.
        'Dim total_stock_volume As Double
        'total_stock_volume = 0

        'This variable is used to input the ticker value in the correct cell in column I.  The ticker changes at a faster rate in column I than in column A.  This variable is therefore used to adjust the rate tickers are copied over from column A to column I.
        j = 0
        total = 0
        'change = 0
        quarterly_change = 0
        SummaryRow = 2
        Dim find_value As Integer
        'SummaryRow = 2
        
        'Find the last row of data in the summary table
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        
        For i = 2 To RowCount
        'If ticker changes then print results
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'Store results in variables
                total = total + Cells(i, 7).Value
                
                'Handle zero total volume
                If total = 0 Then
                    'print the results
                    Range("I" & 2 + j).Value = Cells(i, 1).Value
                    Range("J" & 2 + j).Value = 0
                    Range("K" & 2 + j).Value = "%" & 0
                    Range("L" & 2 + j).Value = 0
    
                Else
                    'Find First non zero starting value
                    If Cells(SummaryRow, 3) = 0 Then
                        For find_value = SummaryRow To i
                            If Cells(find_value, 3).Value <> 0 Then
                                SummaryRow = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    quarterly_change = ws.Cells(i, 6) - Cells(SummaryRow, 3)
                    stock_percent_change = quarterly_change / Cells(SummaryRow, 3)
                    SummaryRow = i + 1
                    
                    'print the results
                    Range("I" & 2 + j).Value = Cells(i, 1).Value
                    Range("J" & 2 + j).Value = quarterly_change
                    Range("J" & 2 + j).NumberFormat = "0.00"
                    Range("K" & 2 + j).Value = stock_percent_change
                    Range("K" & 2 + j).NumberFormat = "0.00%"
                    Range("L" & 2 + j).Value = total
                    
                    'Conditional formatting for positive and negative Quarterly Change
                    Select Case quarterly_change
                        Case Is > 0
                            Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            Range("J" & 2 + j).Interior.ColorIndex = 0
                        End Select
                    
                End If
            'reset variables for ticker
            total = 0
            'change = 0
            quarterly_change = 0
            j = j + 1
            Else
                total = total + Cells(i, 7).Value
            End If
            
        Next i
        
        'take the max and min and place them in a separate part in the workbook
        Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
        Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & RowCount))

        'returns one less becouse header header not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
        
        Range("P2") = Cells(increase_number + 1, 9)
        Range("P3") = Cells(decrease_number + 1, 9)
        Range("P4") = Cells(volume_number + 1, 9)
        
    'This cycles to the next page in the workbook and repeats all the code hitherto.
    Next ws

End Sub
