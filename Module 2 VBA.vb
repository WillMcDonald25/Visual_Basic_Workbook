Sub Ticker():

Dim Ticker_Name As String
Dim Ticker_Total As Double
Ticker_Total = 0
   
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To last_row

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
Ticker_Name = Cells(i, 1).Value
Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
Range("I" & Summary_Table_Row).Value = Ticker_Name
Range("L" & Summary_Table_Row).Value = Ticker_Total
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      Ticker_Total = 0
    Else
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
    End If
Next i
End Sub
-----------------------------------------------------------------------------
Sub Yearlypercent_Change():
Dim Ticker_Name As String
Dim Yearly_Change As Double
Dim percent_change As Double

Yearly_Change = 0

Dim open_price As Double
Dim close_price As Double
Dim price_row As Long
price_row = 2

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To last_row

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


Ticker_Name = Cells(i, 1).Value
open_price = Range("c" & price_row).Value
close_price = Range("F" & i).Value
Yearly_Change = close_price - open_price

If open_price = 0 Then
    percent_change = 0
    Else
    percent_change = Yearly_Change / open_price
    End If

Range("I" & Summary_Table_Row).Value = Ticker_Name
Range("J" & Summary_Table_Row).Value = Yearly_Change
Range("K" & Summary_Table_Row).Value = percent_change

Summary_Table_Row = Summary_Table_Row + 1
price_row = i + 1


Yearly_Change = 0
Else
Yearly_Change = Yearly_Change + (Cells(i, 6).Value - Cells(i, 3).Value)


End If
Next i

End Sub
-------------------------------------------------------------------------
Sub EachWorksheet():

    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Ticker
        Call Yearlypercent_Change
        
    Next
    Application.ScreenUpdating = True


End Sub