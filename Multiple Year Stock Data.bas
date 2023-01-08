Sub stocks()
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 10).Value = "Yearly_change"
    ws.Cells(1, 11).Value = "Percent_change"
    ws.Cells(1, 12).Value = "Total_stock_volume"
    
    
    Dim ticker As String
    Dim vol As LongLong
    vol = 0
    
    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    year_open = 0
    
    Summary_Table_Row = 2
    
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To Last_Row

      If year_open = 0 Then

          year_open = ws.Cells(i, 3).Value
      End If

      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          year_close = ws.Cells(i, 6).Value
          yearly_change = year_close - year_open
          
          ticker = ws.Cells(i, 1).Value


          vol = vol + ws.Cells(i, 7).Value



          ws.Range("I" & Summary_Table_Row).Value = ticker


          ws.Range("J" & Summary_Table_Row).Value = yearly_change

          ws.Range("K" & Summary_Table_Row).Value = ((year_close - year_open) / year_open)

          ws.Range("L" & Summary_Table_Row).Value = vol

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0
          
          year_open = 0
          

      Else

          vol = vol + ws.Cells(i, 7).Value


      End If


    Next i
    
    Last_Row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Amount As Double
    Greatest_Amount = 0
    
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Decrease As Double
    Greatest_Decrease = 0
    
    Dim Greatest_ticker_Volume As String
    Dim Greatest_ticker As Double
    Greatest_ticker = 0
    
    For i = 2 To Last_Row
    
    If ws.Range("K" & i).Value > Greatest_Amount Then
    
    Greatest_Amount = ws.Range("K" & i).Value
    Greatest_Increase_Ticker = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("K" & i).Value < Greatest_Decrease Then
    
    Greatest_Decrease = ws.Range("K" & i).Value
    Greatest_Decrease_Ticker = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("L" & i).Value > Greatest_ticker Then
    
    Greatest_ticker = ws.Range("L" & i).Value
    Greatest_ticker_Volume = ws.Range("I" & i).Value
    
    End If
    
     
    Next i
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"

    ws.Range("o2").Value = Greatest_Increase_Ticker
    ws.Range("p2").Value = Greatest_Amount
    ws.Range("n2").Value = "greatest % increase"
    
    ws.Range("o3").Value = Greatest_Decrease_Ticker
    ws.Range("p3").Value = Greatest_Decrease
    ws.Range("n3").Value = "greatest % Decrease"
    
    ws.Range("o4").Value = Greatest_ticker_Volume
    ws.Range("p4").Value = Greatest_ticker
    ws.Range("n4").Value = "greatest % volume"
 
          
          'Conditional formatting that will highlight positive change in green and negative change in red
          Dim Fmt_Range As Range
          Set Fmt_Range = ws.Range("J2:J" & CStr(Summary_Table_Row - 1))
          Fmt_Range.FormatConditions.Delete
          Fmt_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
          Fmt_Range.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
          Fmt_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
            Formula1:="=0"
          Fmt_Range.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
          

  
          'Conditional formatting that will highlight positive change in green and negative change in red
          
          Set Fmt_Range = ws.Range("K2:K" & CStr(Summary_Table_Row - 1))
          Fmt_Range.FormatConditions.Delete
          Fmt_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
          Fmt_Range.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
          Fmt_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
            Formula1:="=0"
          Fmt_Range.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
          
 
       

Next ws

End Sub
