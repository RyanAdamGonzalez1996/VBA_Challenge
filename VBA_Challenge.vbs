Sub SummarizeStocks()

'Declare Variables
    Dim ticker As String
    Dim currentDate As Integer
    Dim currentOpen As Double
    Dim currentHigh As Double
    Dim currentLow As Double
    Dim currentClose As Double
    Dim initialOpen As Double
    Dim finalClose As Double
    Dim total As Double
    Dim percentChange As Double
    Dim yearlyChange As Double
    
    Dim maxRow As Long
    Dim summaryRow As Integer
    
'Get the Max Row to automaticall adjust loop for sheet size
    maxRow = ActiveSheet.UsedRange.Rows.Count
'Set the Summary Row to start where the summary reports will begin
    summaryRow = 2
    total = 0
    initialOpen = 0
    finalClose = 0
    
'Go Down the First Row; Checking Tickers
    For i = 2 To maxRow
    
        ticker = Cells(i, 1).Value
        'Add the current stock along with the past stocks
        total = Cells(i, 7).Value + total
        finalClose = Cells(i, 6).Value
        
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
                Cells(summaryRow, 10).Value = yearlyChange
                Cells(summaryRow, 11).Value = percentChange
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

End Sub


