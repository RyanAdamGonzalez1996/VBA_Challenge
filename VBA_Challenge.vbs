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
            
            'Create Header for summary chart
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
            'Go Down the First Row; Checking Tickers
            For I = 2 To maxRow
            
                ticker = Cells(I, 1).Value
                'Add the current stock along with the past stocks
                total = Cells(I, 7).Value + total
                finalClose = Cells(I, 6).Value
                
                'prevents the initialOpen value from being overwritten, until the varialb eis reset upon
                'going to the next ticker
                'if this is the first time with a new ticker, take the open value
                If initialOpen = 0 Then
                    initialOpen = Cells(I, 3).Value + totalOpen
                End If
                
                    'If the next Ticker is different
                If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
                    
                    'Added this if to prevent the code from erroring out when dividing by 0
                    'When the code goes through the last iteration
                    If initialOpen > 0 Then
                        yearlyChange = finalClose - initialOpen
                        percentChange = yearlyChange / initialOpen
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
                
            Next I
            
        End With
        
    Next w
   
End Sub


