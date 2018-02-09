Sub stockAnalysis()

Dim LastRow As Double
Dim volume_total As Double
Dim stock_begin As Double
Dim ticker As String
Dim stock_end As Double
Dim percent_change As Variant
Dim stock_change As Double



For Each ws In Worksheets

    'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRow = ws.Range("A1").CurrentRegion.Rows.Count


    
    volume_total = 0
    stock_begin = ws.Cells(2, 3).Value
    
    
    summary_table_row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    For i = 2 To LastRow
         
    
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ticker = ws.Cells(i, 1).Value
            
            stock_end = ws.Cells(i, 6).Value
            
            stock_change = stock_end - stock_begin

            volume_total = volume_total + ws.Cells(i, 7).Value
            
                If stock_begin <> 0 Then
                percent_change = Format((stock_end - stock_begin) / stock_begin, "Percent")
                Else: percent_change = 0
                End If

            ws.Range("I" & summary_table_row).Value = ticker

            ws.Range("L" & summary_table_row).Value = volume_total
            
            ws.Range("J" & summary_table_row).Value = stock_change
                
                If (stock_change < 0) Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                Else: ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                End If
            
            
            ws.Range("K" & summary_table_row).Value = percent_change
            

            summary_table_row = summary_table_row + 1
      
            ticker = 0
            volume_total = 0
            stock_begin = ws.Cells(i + 1, 3).Value
            
            ws.Range("K" & summary_table_row).Value = stock_begin

      Else
                         
           volume_total = volume_total + ws.Cells(i, 7).Value

        End If
            

    Next i
     
    
  Next ws

End Sub