Sub VBAChallenge():
For Each ws In Worksheets

Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row


Dim Ticker As String
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
Dim summary_table As Long
summary_table = 2
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Year_Open As Double
Dim Year_Close As Double


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To Last_Row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & summary_table).Value = Ticker
            ws.Range("L" & summary_table).Value = Total_Stock_Volume
            
            
            Total_Stock_Volume = 0
            
            Year_Open = ws.Range("C" & summary_table)
            Year_Close = ws.Range("F" & summary_table)
            
            Yearly_Change = Year_Close - Year_Open
            
            Percent_Change = Yearly_Change / Year_Open
            
            ws.Range("K" & summary_table).NumberFormat = "0.00%"
            
            ws.Range("J" & summary_table).Value = Yearly_Change
            ws.Range("K" & summary_table).Value = Percent_Change
            
            summary_table = summary_table + 1
            
            
            
        Else
        
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        End If
        
        If ws.Range("J" & summary_table).Value >= 0 Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 3

        Else
           ws.Range("J" & summary_table).Interior.ColorIndex = 4
        
        End If
        
   
    
    
    Next i
        
          
Next ws


End Sub

