Attribute VB_Name = "Module1"
Sub Stocks():

    Dim ws As Worksheet
    
    'Loop through all worksheets
    For Each ws In Worksheets
    
        'Set sheet name variable
        Dim sheet_name As String
        sheet_name = ws.Name
    
        'Acrivate the current worksheet
        Worksheets(sheet_name).Activate
        
        ' Set table headers
        ' ------------------------------------------------
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        '-------------------------------------------------
        
        'Get last row of data (based on first column)
        ' + 1 at end to ensure last ticker data is captured in summarised data
        Dim Last_Row As Long
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
        'Set row number for summarised data, starting in row 2
        Dim Summary As Long
        Summary = 2
        
        'Set Opening Price variable
        Dim Opening_Price As Double
        Opening_Price = 0
        
        'Set Yearly Change variable
        Dim Yearly_Diff As Double
        Yearly_Diff = 0
        
        'Set Percentage Change variable
        Dim Percent_Diff As Double
        Percent_Diff = 0
        
        'Set ticker variable
        Dim Ticker As String
        
        'Set for loop counter variable
        Dim i As Long
        
        'Set Total Stock Volume variable
        Dim Stock_Volume As Double
        Stock_Volume = 0
        
        '------------------CHALLENGE: variables to record greatest increase/decrease and stock volume
        Dim Max_Increase As Double
        Dim Max_Decrease As Double
        Dim Max_Volume As Double
        
        'CHALLENGE: Set table column and row headers
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        
        
            'Create loop to summarise stock data
            For i = 2 To Last_Row
                
                
                'CHALLENGE: if statement to check for
                'greatest increase/decrease and stock volume
                If i = Last_Row Then
                    
                    'CHALLENGE: Find the greatest increase/decrease and stock volume in the summarised data
                    Max_Increase = Application.WorksheetFunction.Max(Range("K2:K" & Summary))
                    Cells(2, 16).Value = Max_Increase
                    Max_Decrease = Application.WorksheetFunction.Min(Range("K2:K" & Summary))
                    Cells(3, 16).Value = Max_Decrease
                    Max_Volume = Application.WorksheetFunction.Max(Range("L2:L" & Summary))
                    Cells(4, 16).Value = Max_Volume
                    
                    'CHALLENGE: Find corresponding ticker value
                    Call Challenge(Max_Increase, Summary, 2, 11)
                    Call Challenge(Max_Decrease, Summary, 3, 11)
                    Call Challenge(Max_Volume, Summary, 4, 12)
                    
                    'CHALLENGE: Set number formats
                    Range("P2:P3").NumberFormat = "0.00%"
                    
                    '------------------END CHALLENGE------------------
                
                
                'Check for different stock ticker and add new line for new ticker
                'ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 6).Value <> 0 Then
                ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    
                    'Store Opening Price
                    Opening_Price = Cells(i, 3).Value
                
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                           
                    'Get ticker value
                    Ticker = Cells(i, 1).Value
                    
                    'Calculate Total Stock Volume
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                    'Calculate Yearly Difference
                    Yearly_Diff = Cells(i, 6).Value - Opening_Price
                    
                    'Calculate Percentage Difference via subroutine to check if Divisor = 0
                    Call Stock_Loop(Percent_Diff, Yearly_Diff, i, Summary)
                    
                    'Output individual ticker data
                    Cells(Summary, 9).Value = Ticker
                    Cells(Summary, 10).Value = Yearly_Diff
                    Cells(Summary, 11).Value = Percent_Diff
                    Cells(Summary, 12).Value = Stock_Volume
                    
                    'Set number formats
                    Cells(Summary, 11).NumberFormat = "0.00%"
                    'Additional number formatting to make stock volume more readable
                    'Cells(summ_row, 12).NumberFormat = "#,##0"
                            
                    'Increment Summary to next row
                    Summary = Summary + 1
                    
                    'Reset Yearly Change and Stock Volume for the next ticker
                    Yearly_Diff = 0
                    Stock_Volume = 0
                    
                Else
                    
                    'Incrementally calculate Stock Volume for the current ticker
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                    'Reset Percentage Change
                    Percent_Diff = 0
                                        
                End If
                
            Next i
            ' ------------------------------------------------
            
    Next ws

End Sub

Sub Stock_Loop(pc_diff_loop, yr_diff_loop, index, summ_row):
                
    'Check if Percentage Difference is being divided by 0
    If (Cells(index, 6).Value - yr_diff_loop) = 0 Then
        'Hard code Percentage Difference to 0
        pc_diff_loop = 0
        
    Else
        'Calculate Percentage Difference
        pc_diff_loop = yr_diff_loop / (Cells(index, 6).Value - yr_diff_loop)
    
    End If
    
    'Format Cell Colours to Yearly Difference summarised data
    If yr_diff_loop > 0 Then
        'Format cell colour green for positive Yearly Difference
        Cells(summ_row, 10).Interior.ColorIndex = 4
    
    ElseIf yr_diff_loop < 0 Then
        'Format cell colour red for negative Yearly Difference
        Cells(summ_row, 10).Interior.ColorIndex = 3
    
    End If
    

End Sub

Sub Challenge(answer, summ_range, output_row, input_column)
    
    'CHALLENGE: Find the corresponding ticker to the max increase/decrease & stock volume
    For j = 2 To summ_range
        If Cells(j, input_column).Value = answer Then
            Cells(output_row, 15).Value = Cells(j, 9).Value
        End If
    Next j
    
End Sub



