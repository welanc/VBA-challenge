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
        Dim Last_Entry As Long
        Last_Entry = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set row number for summarised data, starting in row 2
        Dim Summary As Long
        Summary = 2
        
        'Set Yearly Change variables
        Dim Yearly_Diff As Double
        Yearly_Diff = 0
        
        'Set Percentage Change variables
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
            For i = 2 To Last_Entry
                
                
                'CHALLENGE: if statement to check for
                'greatest increase/decrease and stock volume
                If i = Last_Entry Then
                    
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
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 6).Value <> 0 Then
                    
                    'Calculate Percentage Change
                    Percent_Diff = Yearly_Diff / (Cells(i, 6).Value - Yearly_Diff)
                    
                    'Add cumulative data to ticker via subroutine stockloop
                    Call stock_loop(i, Ticker, Stock_Volume, Percent_Diff, Yearly_Diff, Summary)

                'Check if Percentage Change is being divided by 0 (result = math error)
                'and if so, hard code Percent_Diff to 0
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 6).Value = 0 Then
                    
                    'Hard code Percentage Change to 0
                    Percent_Diff = 0
                    
                    'Add cumulative data to ticker via subroutine stockloop
                    Call stock_loop(i, Ticker, Stock_Volume, Percent_Diff, Yearly_Diff, Summary)
                    
                    
                Else
                    
                    'Incrementally calculate Yearly Change and Stock Volume for the current ticker
                    Yearly_Diff = Yearly_Diff + (Cells(i + 1, 6).Value - Cells(i, 6).Value)
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                    'Reset Percentage Change
                    Percent_Diff = 0
                                        
                End If
                
            Next i
            ' ------------------------------------------------
            
    Next ws

End Sub

Sub stock_loop(index, tick_loop, stock_vol_loop, pc_diff_loop, yr_diff_loop, summ_row):
                
    'Get ticker value
    tick_loop = Cells(index, 1).Value
    
    'Calculate Total Stock Volume
    stock_vol_loop = stock_vol_loop + Cells(index, 7).Value
    
    'Output individual ticker data
    Cells(summ_row, 9).Value = tick_loop
    Cells(summ_row, 10).Value = yr_diff_loop
    Cells(summ_row, 11).Value = pc_diff_loop
    Cells(summ_row, 12).Value = stock_vol_loop
    
    If yr_diff_loop > 0 Then
        Cells(summ_row, 10).Interior.ColorIndex = 4
    
    ElseIf yr_diff_loop < 0 Then
        Cells(summ_row, 10).Interior.ColorIndex = 3
    
    End If
    
    'Set number formats
    Cells(summ_row, 11).NumberFormat = "0.00%"
    'Additional number formatting to make stock volume more readable
    'Cells(summ_row, 12).NumberFormat = "#,##0"
            
    'Increment Summary to next row
    summ_row = summ_row + 1
    
    'Reset Yearly Change and Stock Volume for the next ticker
    yr_diff_loop = 0
    stock_vol_loop = 0
    

End Sub

Sub Challenge(answer, summ_range, output_row, input_column)
    
    'CHALLENGE: Find the corresponding ticker to the max increase/decrease & stock volume
    For j = 2 To summ_range
        If Cells(j, input_column).Value = answer Then
            Cells(output_row, 15).Value = Cells(j, 9).Value
        End If
    Next j
    
End Sub



