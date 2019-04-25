Sub Mykyta_Yarovoi_StockData()
    
    'LOOP through ALL SHEETS
    
    Dim Ws As Worksheet
    For Each Ws In ActiveWorkbook.Worksheets
    Ws.Activate
    
    ' Find the LAST ROW on every Sheet
    
    LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the HEADERS
    
    Ws.Range("I1").Value = "Ticker"
    Ws.Range("J1").Value = "Yearly Change"
    Ws.Range("K1").Value = "Percent Change"
    Ws.Range("L1").Value = "Total Stock Volume"

    ' Set VARIABLES
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Total_Volume As LongLong
        Total_Volume = 0
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim Summary_Table_Column As Integer
        Summary_Table_Column = 1
    Dim i As Long
   
    ' Set OPEN PRICE
    
    Open_Price = Cells(2, Summary_Table_Column + 2).Value
            
    ' Loop through all TICKERS
    
    For i = 2 To LastRow
        
        ' Let's check IF we are STILL within the same TICKER
        
        If Cells(i + 1, Summary_Table_Column).Value <> Cells(i, Summary_Table_Column).Value Then
            
            ' Set TICKER
            
            Ticker = Cells(i, Summary_Table_Column).Value
            Cells(Summary_Table_Row, Summary_Table_Column + 8).Value = Ticker
            
            ' Set CLOSE PRICE
            
            Close_Price = Cells(i, Summary_Table_Column + 5).Value
            
            ' Set YEARLY CHANGE
            
            Yearly_Change = Close_Price - Open_Price
            Cells(Summary_Table_Row, Summary_Table_Column + 9).Value = Yearly_Change
            
            ' Set PERCENT CHANGE
            
            If (Open_Price = 0 And Close_Price = 0) Then
                Percent_Change = 0
                
            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                Percent_Change = 1
                
            Else
                
                Percent_Change = Yearly_Change / Open_Price
                Cells(Summary_Table_Row, Summary_Table_Column + 10).Value = Percent_Change
                Cells(Summary_Table_Row, Summary_Table_Column + 10).NumberFormat = "0.00%"
                
            End If
            
            ' Set TOTAL VOLUME
            
            Total_Volume = Total_Volume + Cells(i, Summary_Table_Column + 6).Value
            Cells(Summary_Table_Row, Summary_Table_Column + 11).Value = Total_Volume
            
            ' Add 1 to the SUMMARY TABLE ROW
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the OPEN PRICE
            
            Open_Price = Cells(i + 1, Summary_Table_Column + 2)
            
            ' Reset the VOLUME
            
            Total_Volume = 0
            
        Else 'if cells are STILL the SAME TICKER
        
            Total_Volume = Total_Volume + Cells(i, Summary_Table_Column + 6).Value
            
        End If
        
    Next i

    ' Determine the LAST ROW of YEARLY CHANGE per Worksheet
    
    YearlyChangeLastRow = Ws.Cells(Rows.Count, Summary_Table_Column + 8).End(xlUp).Row
    
    ' Set the CELL COLORS
    
    For i = 2 To YearlyChangeLastRow
    
        If (Cells(i, Summary_Table_Column + 9).Value > 0 Or Cells(i, Summary_Table_Column + 9).Value = 0) Then
            Cells(i, Summary_Table_Column + 9).Interior.ColorIndex = 10
            
        ElseIf Cells(i, Summary_Table_Column + 9).Value < 0 Then
            Cells(i, Summary_Table_Column + 9).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
    ' Set HEADERS "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" + "Ticker" and "Value"
    
        Cells(2, Summary_Table_Column + 14).Value = "Greatest % Increase"
        Cells(3, Summary_Table_Column + 14).Value = "Greatest % Decrease"
        Cells(4, Summary_Table_Column + 14).Value = "Greatest Total Volume"
        Cells(1, Summary_Table_Column + 15).Value = "Ticker"
        Cells(1, Summary_Table_Column + 16).Value = "Value"
    
        ' Look through each row to Find the GREATEST VALUE and its associate TICKER
    
        For i = 2 To YearlyChangeLastRow
        
            If Cells(i, Summary_Table_Column + 10).Value = Application.WorksheetFunction.Max(Ws.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(2, Summary_Table_Column + 15).Value = Cells(i, Summary_Table_Column + 8).Value
                Cells(2, Summary_Table_Column + 16).Value = Cells(i, Summary_Table_Column + 10).Value
                Cells(2, Summary_Table_Column + 16).NumberFormat = "0.00%"
                
            ElseIf Cells(i, Summary_Table_Column + 10).Value = Application.WorksheetFunction.Min(Ws.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(3, Summary_Table_Column + 15).Value = Cells(i, Summary_Table_Column + 8).Value
                Cells(3, Summary_Table_Column + 16).Value = Cells(i, Summary_Table_Column + 10).Value
                Cells(3, Summary_Table_Column + 16).NumberFormat = "0.00%"
                
            ElseIf Cells(i, Summary_Table_Column + 11).Value = Application.WorksheetFunction.Max(Ws.Range("L2:L" & YearlyChangeLastRow)) Then
                Cells(4, Summary_Table_Column + 15).Value = Cells(i, Summary_Table_Column + 8).Value
                Cells(4, Summary_Table_Column + 16).Value = Cells(i, Summary_Table_Column + 11).Value
                
            End If
        
        Next i
        
 Next Ws
       
End Sub



