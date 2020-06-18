Sub LoopThroughWorksheets()
    'Define Variables
    Dim WS As Worksheet
        
    'run through each worksheet
    For Each WS In Worksheets
     Dim ticker As String
     ticker = ""
     Dim open_price As Double
     open_price = 0
     Dim close_price As Double
     close_price = 0
     Dim Yearly_change As Double
     Yearly_change = 0
     Dim Percent_Change As Double
     Percent_Change = 0
     Dim Total_Stock_Volume As Double
     Total_Stock_Volume = 0
     Dim Summary_Table_Row As Long
     Summary_Table_Row = 2
     Dim Lastrow As Long
     Lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
     Dim i As Long
     i = 2
     
    
     'Variables for Challenge
     Dim Increase_Ticker As String
     Increase_Ticker = ""
     Dim Decrease_Ticker As String
     Decrease_Ticker = ""
     Dim Total_Volume_Ticker As String
     Total_Volume_Ticker = ""
     Dim Greatest_Increase As Double
     Greatest_Increase = 0
     Dim Greatest_Decrease As Double
     Greatest_Decrease = 0
     Dim Greatest_Total_Volume As Double
     Greatest_Total_Volume = 0
     Dim j As Long
     j = 2
    
               
        'set Column names,create summary table
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
        
             'set Challenge Summary Table,names
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        
        
                      
        'Loop Through Worksheets
        For i = 2 To Lastrow
            'To fix open price of the month
            If WS.Cells(i, 1).Value <> WS.Cells(i - 1, 1).Value Then
                open_price = WS.Cells(i, 3).Value
            End If
            'To determine the total stock volume for the year
            Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
            
            'To see if it's the same Ticker
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                'Find Values
                ticker = WS.Cells(i, 1).Value
                close_price = WS.Cells(i, 6).Value
                Yearly_change = close_price - open_price
                If open_price = 0 Then
                  Percent_Change = 0
                ElseIf open_price <> 0 Then
                  Percent_Change = (Yearly_change / open_price)
                End If
                WS.Columns("K").NumberFormat = "0.00%"
                              
            
                'Fill the Results Table and color
                WS.Cells(Summary_Table_Row, 9).Value = ticker
                WS.Cells(Summary_Table_Row, 10).Value = Yearly_change
                If (Yearly_change > 0) Then
                    WS.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                ElseIf (Yearly_change <= 0) Then
                    WS.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                End If
                WS.Cells(Summary_Table_Row, 11).Value = Percent_Change
                WS.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset Values
                Yearly_change = 0
                close_price = 0
                open_price = 0
                Total_Stock_Volume = 0
            End If
        Next i
            
        Lastrow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To Lastrow
                        
            If WS.Cells(j, 11).Value > Greatest_Increase Then
               Greatest_Increase = WS.Cells(j, 11).Value
               Increase_Ticker = WS.Cells(j, 9).Value
            End If
    
            If WS.Cells(j, 11).Value < Greatest_Decrease Then
               Greatest_Decrease = WS.Cells(j, 11).Value
               Decrease_Ticker = WS.Cells(j, 9).Value
            End If
        
            If WS.Cells(j, 12).Value > Greatest_Total_Volume Then
                Greatest_Total_Volume = WS.Cells(j, 12).Value
                Total_Volume_Ticker = WS.Cells(j, 9).Value
            End If
    
        Next j
   
            'Fill Greatest_Values_Table
           WS.Cells(2, 17).Value = Greatest_Increase
           WS.Cells(2, 17).NumberFormat = "0.00%"
           WS.Cells(3, 17).Value = Greatest_Decrease
           WS.Cells(3, 17).NumberFormat = "0.00%"
           WS.Cells(4, 17).Value = Greatest_Total_Volume
           WS.Cells(2, 16).Value = Increase_Ticker
           WS.Cells(3, 16).Value = Decrease_Ticker
           WS.Cells(4, 16).Value = Total_Volume_Ticker
           
           Greatest_Increase = 0
           Greatest_Decrease = 0
           Greatest_Total_Volume = 0
        
    Next WS
    
  ' Challenge
End Sub

