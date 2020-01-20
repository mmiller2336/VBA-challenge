Attribute VB_Name = "Module1"
Sub Stocks()

' RUN THE CODE TWICE TO MAKE THE COLOR CODING WORK

For Each ws In Worksheets
    Dim ticker As String
    Dim vol As Double
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Worksheet As String
    Range("H1").Value = "Ticker"
    Range("K1").Value = "Total_Stock_Volume"
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    
    
    WorksheetName = ws.Name


    ws.Cells(1, 8).Value = "ticker"
    ws.Cells(1, 9).Value = "Yearly_change"
    ws.Cells(1, 10).Value = "Yearly_percentage"
    ws.Cells(1, 11).Value = "Total Stock Vol"

    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
        

' loop through the stock data
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            vol = ws.Cells(i, 7).Value
            
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - year_open
            If year_open <> 0 Then
            percent_change = yearly_change / year_close
            
            Else
            year_open = 0
            End If
            
             
            vol = vol + ws.Cells(i, 7).Value
            
            ws.Range("H" & Summary_Table_Row).Value = ticker
            ws.Range("I" & Summary_Table_Row).Value = yearly_change
            ws.Range("J" & Summary_Table_Row).Value = percent_change
            ws.Range("K" & Summary_Table_Row).Value = vol
        
            Summary_Table_Row = Summary_Table_Row + 1
        
            
        year_open = ws.Cells(i + 1, 3).Value
        Else
        
            vol = vol + ws.Cells(i, 7).Value
            
        End If
        
        'If ws.Cells(i, 10).Value < 0 Then
        'ws.Cells(i, 10).Interior.ColorIndex = 3
  
        'Else
        'ws.Cells(i, 10).Interior.ColorIndex = 4
     
        'End If
    
       
        ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("J2:J" & lastrow)) * 100
        ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("J3:J" & lastrow)) * 100
        ws.Range("P4") = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))

      
Next i

Next ws



End Sub
