Attribute VB_Name = "Module1"
Sub SummarizeStocks()

'Declare Variables
    Dim ticker As String
    Dim initialOpen As Double
    Dim finalClose As Double
    Dim total As Double
    Dim percentChange As Double
    Dim yearlyChange As Double
    Dim maxRow As Long
    Dim summaryRow As Integer
    Dim WS_Count As Integer
    Dim maxStock As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxStockTicker As String
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    
    'Get the number of worksheets for use in looping
    WS_Count = ActiveWorkbook.Worksheets.Count

    'Loop Through all of the worksheets and activate them so we can make changes to each
    For w = 1 To WS_Count
        'Activate the current worksheet so that the changes are made into it
        Worksheets(w).Activate
            
        With ActiveWorkbook.Worksheets(w)
            'Get the Max Row to automaticall adjust loop for sheet size
            maxRow = ActiveSheet.UsedRange.Rows.Count
            
            'Set the Summary Row to start where the summary reports will begin
            summaryRow = 2
            total = 0
            initialOpen = 0
            finalClose = 0
            
            'Set the Max Variables
            maxStock = 0
            maxIncrease = 0
            maxDecrese = -1
            
            'Create Header for summary chart
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
            'Go Down the First Row; Checking Tickers
            For i = 2 To maxRow
            
                ticker = Cells(i, 1).Value
                'Add the current stock along with the past stocks
                total = Cells(i, 7).Value + total
                finalClose = Cells(i, 6).Value
                
                'Keep Track of the Highest Stock, whenever a new high stock is shown, update variable
                If Cells(i, 7).Value > maxStock Then
                    maxStock = Cells(i, 7).Value
                    maxStockTicker = Cells(i, 1).Value
                End If
                'prevents the initialOpen value from being overwritten, until the varialb eis reset upon
                'going to the next ticker
                'if this is the first time with a new ticker, take the open value
                If initialOpen = 0 Then
                    initialOpen = Cells(i, 3).Value + totalOpen
                End If
                
                    'If the next Ticker is different
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                    
                    'Added this if to prevent the code from erroring out when dividing by 0
                    'When the code goes through the last iteration
                    If initialOpen > 0 Then
                        yearlyChange = finalClose - initialOpen
                        percentChange = yearlyChange / initialOpen
                        
                        'Keep Track of the Highest percentChange
                        If percentChange > maxIncrease Then
                            maxIncrease = percentChange
                            maxIncreaseTicker = Cells(i, 1).Value
                        End If
                        'Keep track of the Lowest percentChange
                        If percentChange < 0 Then
                            If (percentChange * -1) > (maxDecrease * -1) Then
                                maxDecrease = percentChange
                                maxDecreaseTicker = Cells(i, 1).Value
                            End If
                        End If
                        
                        Cells(summaryRow, 10).Value = yearlyChange
                        Cells(summaryRow, 11).Value = percentChange
                        'Format the percent Change cells into a 0.00% setup
                        Cells(summaryRow, 11).NumberFormat = "0.00%"
                        
                    'Fill the cell for yearly change either green if positive or red if negative
                        If yearlyChange > 0 Then
                            Cells(summaryRow, 10).Interior.ColorIndex = 4
                        Else
                            Cells(summaryRow, 10).Interior.ColorIndex = 3
                        End If
                        
                    End If
                    
                    'Put values into Summary Chart
                    Cells(summaryRow, 9).Value = ticker
                    Cells(summaryRow, 12).Value = total
            
                    'Reset variables
                    'Go down 1 row in summary chart
                    total = 0
                    initialOpen = 0
                    finalClose = 0
                    summaryRow = summaryRow + 1
                        
                End If
                
            Next i
            
            'Add the Max results chart
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrese"
            Cells(4, 15).Value = "Greatest total Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            
            'Input Max Variables
            Cells(2, 16).Value = maxIncreaseTicker
            Cells(2, 17).Value = maxIncrease
            Cells(2, 17).NumberFormat = "0.00%"
            Cells(3, 16).Value = maxDecreaseTicker
            Cells(3, 17).Value = maxDecrease
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(4, 16).Value = maxStockTicker
            Cells(4, 17).Value = maxStock
            
        End With
        
    Next w
   
End Sub


