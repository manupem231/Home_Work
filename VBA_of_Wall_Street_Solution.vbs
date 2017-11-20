Sub stock_data_analysis()
    ' Variable to hold Row count of an Active sheet
    Dim Row_count As Double
    ' Set an initial variable for holding the total volume per stock
    Dim Total_vol As Double

    ' Keep track of the location for each stock in the summary table
    Dim Summary_Index As Double
    ' Declare Start and End index for holding Opening and Closing Price Index
    Dim Start_Index As Double
    Dim End_Index As Double
    
    
    ' Declare a variable to hold yearly changes of a stock
    Dim Yrly_Chg As Double
    ' Variable to hold cell count of Percentage Change column
    Dim pc_count As Double
    
    ' Variables for holding the greatest percentage increase, decrease and stock volume
    Dim Max_inc As Double
    Dim Max_dec As Double
    Dim Max_vol As Double
            
    ' Loop through each worksheet
    For Each ws In Worksheets
        ' Get the row count from the current/active sheet
        Row_count = ws.UsedRange.Rows.Count
        ' Initialize Total volume to zero
        Total_vol = 0
        ' Set the headers for the result set
        ws.Cells(1, 9).Value = "Ticker" 'Column I
        ws.Cells(1, 10).Value = "Total Stock Volume" 'Column J
        ws.Cells(1, 11).Value = "Yearly Change" 'Column K
        ws.Cells(1, 12).Value = "Percent Change" 'Column L
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Set the text for final result set
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Initialize the Summary Index to 2
        Summary_Index = 2
        Start_Index = 2
            ' Loop through each rows in the current sheet
              For i = 2 To Row_count
                    
                    ' Check if we are still within the same Stock, if it is not...
                    If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                      End_Index = i
                      ' Add to the Total Stock Volume
                      Total_vol = Total_vol + ws.Cells(i, 7).Value
                      ' Print the Stock in the Summary Table
                      ws.Range("I" & Summary_Index).Value = ws.Cells(i, 1).Value
                      ' Print the Total Volume to the Summary Table
                      ws.Range("J" & Summary_Index).Value = Total_vol
                      
                      ' Calculate yearly change by subtracting opening price(start of the year)
                      ' from closing price(Year end)
                        
                      Yrly_Chg = ws.Cells(End_Index, 6).Value - ws.Cells(Start_Index, 3).Value
                      ' Write the result into column K
                      ws.Cells(Summary_Index, 11).Value = Yrly_Chg
                        
                      ' Caculate Percentage change as Yearly change/Opening Price
                      If ws.Cells(Start_Index, 3).Value <> 0 Then
                        ws.Cells(Summary_Index, 12).Value = Yrly_Chg / ws.Cells(Start_Index, 3).Value
                        ' Format Percentage change to %
                        ws.Cells(Summary_Index, 12).NumberFormat = "0.00%"
                      End If
                      ' set the color index to green when yearly change is positive or red when it is negative
                      If Yrly_Chg > 0 Then
                         ws.Cells(Summary_Index, 11).Interior.ColorIndex = 4
                      Else
                         ws.Cells(Summary_Index, 11).Interior.ColorIndex = 3
                      End If
                      ' Excel function
                      ' Cells(Summary_Index, 13).Value = Excel.WorksheetFunction.Sum(ws.Range(Cells(Start_Index, 7), Cells(End_Index, 7)))
                      
                      ' Add one to the summary table row
                      Summary_Index = Summary_Index + 1
                      ' Reset Start Index
                      Start_Index = End_Index + 1
                      ' Reset the Total Volume
                      Total_vol = 0
                      ' If the cell immediately following a row is the same stock...
                    Else
                      ' Add to the Total Volume
                      Total_vol = Total_vol + ws.Cells(i, 7).Value
                   End If
              Next i
              
            ' Print the cell count of Percentage Change column 12 (Range L)
            ' MsgBox (Range("L1").End(xlDown).Row)
             
            pc_count = ws.Range("L1").End(xlDown).Row
            Max_inc = 0
            Max_dec = 0
            Max_vol = 0
            
            ' loop through Percentage Change column and find the greatest increase, decrease and volume
            For i = 2 To pc_count
                If ws.Cells(i, 12).Value > Max_inc Then
                    Max_inc = ws.Cells(i, 12).Value
                    ws.Range("P2").Value = ws.Cells(i, 9).Value 'set the ticker name
                End If
                If ws.Cells(i, 12).Value < Max_dec Then
                    Max_dec = ws.Cells(i, 12).Value
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                End If
                If ws.Cells(i, 10).Value > Max_vol Then
                    Max_vol = ws.Cells(i, 10).Value
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                End If
            Next i
 
            ' set the max values for increase and decrease
            ws.Range("Q2").Value = Max_inc
            ws.Range("Q3").Value = Max_dec
            ws.Range("Q2", "Q3").NumberFormat = "0.00%"
            ' set the max volume
            ws.Range("Q4").Clear
            ws.Range("Q4").Value = Max_vol

    Next ws
    
End Sub
