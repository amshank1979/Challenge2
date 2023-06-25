Sub Ticker()
    For Each WS In Worksheets
        WS.Activate
        Columns("I:M").Delete
    'Format Spreadsheet
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percent Change"
    [L1] = "Total Stock Volume"
    Columns("A:L").AutoFit
    'Exit Sub
  ' Set an initial variable for holding the Ticker Symbol
  Dim Ticker As String

  ' Set an initial variable for holding the total per TickerSymbol
  Dim Ticker_Total As Double
  Ticker_Total = 0

  ' Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Open_Price_Pointer = 2
  
'Declare last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Set Yearly Change
Dim Opening_Value As Double
Dim Closing_Value As Double

  ' Loop through all symbols purchases
  For i = 2 To lastrow
    Ticker = Cells(i, "A").Value
   

    ' Check if we are still within the same Ticker Symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Add to the Stock Volume Total
      Stock_Volume = Stock_Volume + Cells(i, "G").Value
      opening_price = Cells(Open_Price_Pointer, "C").Value
      closing_price = Cells(i, "F").Value
      
      Cells(Summary_Table_Row, "I").Value = Ticker
      Cells(Summary_Table_Row, "J").Value = closing_price - opening_price
        Cells(Summary_Table_Row, "K").Value = (closing_price - opening_price) / opening_price
        Cells(Summary_Table_Row, "L").Value = Stock_Volume
        
        If (closing_price - opening_price) > 0 Then
            Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
        Else
            Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
        End If
        
        
      Open_Price_Pointer = i + 1
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock_Volume
      Stock_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Stock_Volume
      Stock_Volume = Stock_Volume + Cells(i, "G").Value

    End If

  Next i

 Next WS

    MsgBox ("Finished")
End Sub
