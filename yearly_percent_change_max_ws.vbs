Sub yearly_percent_change_max_ws()
'---------------------------------------------------------------------------------------------------------------------------
'Find the following outputs for the stocks for one year :
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'----------------------------------------------------------------------------------------------------------------------------
 
 '--------------------------------------------------------------------------------------------------
 'i : the index of iteretions for overall operations, from 1 to the number of diffent tickers.
 'k: the number of row to display the values I will find, from 2 to the number of different tickers +1.
 '--------------------------------------------------------------------------------------------------
 Dim i, k As Long
 
 '------------------------------------------
 ' ticker_now : one of different tickers
 '------------------------------------------
 Dim ticker_now As String
 
 '--------------------------------------------------
 ' row_now : the first row where ticker_now appears.
 '--------------------------------------------------
 Dim row, row_now As Long

 '---------------------------------------------------------------------------------------------------------------------------------
 ' YearlyChange: Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
 ' PercentChange:The percentage change from the opening price at the beginning of a given year to the closing price at the end of
  'that year.
 '---------------------------------------------------------------------------------------------------------------------------------
 Dim YearlyChange, PercentChange As Double
 
 '-------------------------------------------------------------------------------
 'open_price_begin : the opening price at the beginning of a given year
 'close_price_end : the closing price at the end of that year
 '-------------------------------------------------------------------------------
 Dim open_price_begin As Double
 Dim close_price_end As Double
 
 '------------------------------------------
 'total_volume: total volume of ticker_now
 '------------------------------------------
 Dim total_volume As LongLong
 
 ' --------------------------------------------
 ' LOOP THROUGH ALL SHEETS
 ' --------------------------------------------
For Each ws In Worksheets

 
  '-----------------------------------------------------------------------------------
 ' lastrow, lastcolumn: the number of rows, column in current worksheet, respectively
 '-----------------------------------------------------------------------------------
 Dim lastrow, lastcolumn As Long
 
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
 lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
 
 
 
 '------------------------------------------
 ' the names of the columns we will create.
 '------------------------------------------
  ws.Cells(1, lastcolumn + 2) = "Ticker"
  ws.Cells(1, lastcolumn + 3) = "Yearly Change"
  ws.Cells(1, lastcolumn + 4) = "Percent Change"
  ws.Cells(1, lastcolumn + 5) = "Total Stock Volume"
  
  '--------------------------------------------------------------------------------------------
  ' tickerCount :the number of different tickers.
  '--------------------------------------------------------------------------------------------
  tickerCount = 1
  ticker_now = ws.Range("A2").Value
  For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      tickerCount = tickerCount + 1
    End If
  Next i

  
  '--------------------------------------------------------
  ' Start with the first ticker at (2,1)
  '-------------------------------------------------------
  row_now = 2
  k = 1
 
   
  ' Iterate until the last of rows in the current worksheet, starting with row_now
    total_volume = 0
    For row = 2 To lastrow
  '--------------------------------------------------------------------------------
  ' Iterate until we meet the next ticker that is different from ticker_now
  ' If we meet the next different ticker, calculate YearlyChange and PercentChange
  ' from data set just before the next ticker.
  ' And then stop this 'for' sentence.
  '
  ' This iterations also end if we meet the last row.
  ' The reason why I added this condition is that for the case of the last ticker,
  ' there is no more different ticker to satisfy the former condition.
  '--------------------------------------------------------------------------------
         If (ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value) Then
            ticker_now = ws.Cells(row, 1).Value
            open_price_begin = ws.Cells(row_now, 3).Value 'the first value of ticker_now in <open> column
            
            close_price_end = ws.Cells(row, 6).Value 'the last value of ticker_now in <close> column
           
            YearlyChange = close_price_end - open_price_begin
            PercentChange = YearlyChange / open_price_begin
            
            ' Round the values to the desired decimal place.
            YearlyChange = Application.WorksheetFunction.Round(YearlyChange, 2)
            PercentChange = Application.WorksheetFunction.Round(PercentChange, 4)
            k = k + 1
            '-----------------------------
            ' Display the values on cells
            '-----------------------------
            ws.Cells(k, lastcolumn + 2).Value = ws.Cells(row, 1).Value
            ws.Cells(k, lastcolumn + 3).Value = YearlyChange
            ws.Cells(k, lastcolumn + 4).Value = PercentChange
            
            '------------------------------------------------------
            ' Change the format of PercentChange to the form 0.00%
            '------------------------------------------------------
            ws.Cells(k, lastcolumn + 4).NumberFormat = "0.00%"
            
            ws.Cells(k, lastcolumn + 5).Value = total_volume
            
            '-----------------------------------------------------------------------------------------
            ' Highlight positive change in green and negative change in red in "Yearly Change" column.
            '-----------------------------------------------------------------------------------------
            If YearlyChange < 0 Then
              ' ws.Cells(k, lastcolumn + 3).Interior.Color = RGB(255, 0, 0)
               ws.Cells(k, lastcolumn + 3).Interior.ColorIndex = 3
            Else
               'ws.Cells(k, lastcolumn + 3).Interior.Color = RGB(0, 255, 0)
               ws.Cells(k, lastcolumn + 3).Interior.ColorIndex = 4
            End If
            
            '-------------------------------------------------------
            ' Update the current ticker and their position of row
            '-------------------------------------------------------
            row_now = row + 1
            total_volume = 0
            
         Else
           total_volume = total_volume + ws.Cells(row, 7).Value
         End If
  
    Next row
  
 '--------------------------------------------------------------------------------------------------------
 ' Find the greatest increase and decrease of percent changes and the maximum value of total stock volume
 ' It is possible to insert this code in the previous "for" sentence, but separate it for simplicity.
 '--------------------------------------------------------------------------------------------------------
  
'------------------------------------------------------------------------------
 ' great_increase : the greatest increase of percent changes
 ' great_decrease : the greatest decrease of percent changes
 ' max_volume : the maximum value of total stock volume
 ' ticker_increase : the ticker having the greaeast increase of percent change
 ' ticker_decrease : the ticker having the greaeast decrease of percent change
 ' ticker_volume : the ticker having the maximum value of total stock volume
 '-----------------------------------------------------------------------------
 Dim great_increase, great_decrease As Double
 Dim ticker_increase, ticker_decrease, ticker_volume As String
 Dim max_volume As LongLong
 Dim j As Integer
 
 '---------------------
 ' Initial values
 '---------------------
 great_increase = 0
 great_decrease = 0
 max_volume = 0
 
 '------------------------------------------------
 ' Iterate until the number of difference tickers
 '------------------------------------------------
 For j = 2 To tickerCount
    PercentChange = ws.Cells(j, lastcolumn + 4).Value
 '----------------------------------------------------
 ' Find great_increase and the ticker with that value
 '----------------------------------------------------
    If PercentChange > great_increase Then
       great_increase = PercentChange
       ticker_increase = ws.Cells(j, lastcolumn + 2).Value
 '----------------------------------------------------
 ' Find great_decrease and the ticker with that value
 '----------------------------------------------------
    ElseIf PercentChange < great_decrease Then
       great_decrease = PercentChange
       ticker_decrease = ws.Cells(j, lastcolumn + 2).Value
    End If
 '----------------------------------------------------
 ' Find max_volume and the ticker with that value
 '----------------------------------------------------
    total_volume = ws.Cells(j, lastcolumn + 5).Value
    If total_volume > max_volume Then
       max_volume = total_volume
       ticker_volume = ws.Cells(j, lastcolumn + 2).Value
    End If
 Next j
 
 For i = 2 To 3
   ws.Cells(i, lastcolumn + 10).NumberFormat = "0.00%"
 Next i

 ws.Cells(2, lastcolumn + 8) = "Greatest %Increase"
 ws.Cells(3, lastcolumn + 8) = "Greatest %Decrease"
 ws.Cells(4, lastcolumn + 8) = "Greatest Total Volume"
 
 ws.Cells(1, lastcolumn + 9) = "Ticker"
 ws.Cells(1, lastcolumn + 10) = "Value"
 
 ws.Cells(2, lastcolumn + 10).Value = great_increase
 ws.Cells(3, lastcolumn + 10).Value = great_decrease
 ws.Cells(4, lastcolumn + 10).Value = max_volume
 
 ws.Cells(2, lastcolumn + 9).Value = ticker_increase
 ws.Cells(3, lastcolumn + 9).Value = ticker_decrease
 ws.Cells(4, lastcolumn + 9).Value = ticker_volume
 
 Next ws

End Sub
