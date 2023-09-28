Sub MultiYear_multiStock_example():
 ' To repeat the multistock program across all worksheets 
 For Each ws in Worksheets
  ' Set an initial variable for holding the ticker name
  Dim Ticker As String

  ' Set an initial variable for holding the Yearly change, Percentage Change 
  ' total volume of stock per ticker 
  Dim Yearly_Change As Double
  Dim Percentage_Change As Double
  Dim Total_Stock As Double
  Total_Stock = 0
  
  'Find the last row 
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Track location for each ticker in the summary table
  Dim Ticker_Summary_Table_Row As Integer
  Ticker_Summary_Table_Row = 4
  
  'Variables for holding the Open Value, Close Value, Greatest % increase 
  'and decrease,greatest total volume and the tickers for these values 
  Dim Open_Value As Double
  Dim Close_Value As Double
  Dim Max_Change as Double
  Dim Min_Change As Double
  Dim Max_Volume as Double
  Dim Min_Ticker AS String
  Dim Max_Ticker As String
  Dim Max_Volume_Ticker As String
  Max_Change = 0
  Min_Change = 0
  Max_Volume=0
  Open_Value= ws.Cells(2,3).Value

  ' Set the rows/column names for the new data tables
  ws.Cells(3,10).Value = "Ticker"
  ws.Cells(3, 11).Value = "Yearly Change"
  ws.Cells(3, 12).Value = "Percent Change"
  ws.Cells(3, 13).Value = "Total Stock Volume"
  ws.Cells(3,17).Value = "Ticker"
  ws.Cells(3,18).Value = "Value "
  ws.Cells(4,16).Value = "Greatest % increase"
  ws.Cells(5,16).Value = "Greatest % decrease"
  ws.Cells(6,16).Value = "Greatest Total Volume"
  'Format the value in the Range to a Percentage value
  ws.Range("L4:L" &LastRow).NumberFormat="0.00%"
  ' Loop to iterate through all rows till the lastrow
  For i = 2 To LastRow
    
    ' This will check the value of the ticker, if it is the same. If not then 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
      ' Set Ticker name in the new table
      Ticker = ws.Cells(i, 1).Value
      ' Add to the Total Stock Volume
      Total_Stock = Total_Stock + ws.Cells(i, 7).Value
      Close_Value=ws.Cells(i,6).Value
      ' Printing the Ticker name in the Ticker Summary Table
      ws.Range("J" & Ticker_Summary_Table_Row).Value = Ticker
      Yearly_Change = Close_Value - Open_Value
      Percentage_Change = (Yearly_Change / Open_Value)
      ' Printing the Yearly Change and Percentage Change to Summary Table
      ws.Range("K" & Ticker_Summary_Table_Row).Value = Yearly_Change
      ws.Range("L" & Ticker_Summary_Table_Row).Value = Percentage_Change
      ' Printing the Total Stock Volume to the Summary Table
      ws.Range("M" & Ticker_Summary_Table_Row).Value = Total_Stock
      
      ' This will set the positive changes to Green and negative to Red
      ' for Yearly and Percentage Change Columns
      if Yearly_Change >=0 then
      ws.Range("K" & Ticker_Summary_Table_Row).Interior.ColorIndex=4
      ws.Range("L" & Ticker_Summary_Table_Row).Interior.ColorIndex=4
      Else
      ws.Range("K" & Ticker_Summary_Table_Row).Interior.ColorIndex=3
      ws.Range("L" & Ticker_Summary_Table_Row).Interior.ColorIndex=3
      End If
      ' To find the Greatest % increase and Greatest % decrease
      if Percentage_Change > Max_Change then
        Max_Change = Percentage_Change
        Max_Ticker =Ticker
      Elseif Percentage_Change < Min_Change then
        Min_Change = Percentage_Change
        Min_Ticker = Ticker
      End if
      ' To find the Greatest Total Volume
      if (Total_stock > Max_Volume) then
        Max_Volume = Total_Stock
        Max_Volume_Ticker = Ticker
      End if
      ' Icrement summary table row to move down to the next row
      Ticker_Summary_Table_Row = Ticker_Summary_Table_Row + 1
      
      ' Reset the Total Volume of stock ,Percentage change
      Total_Stock = 0
      Percentage_Change = 0
      'Reset the Opening value
      Open_Value= ws.Cells(i+1,3).Value
      
    Else
      ' Add to the  Total Stock volume
      Total_Stock = Total_Stock + ws.Cells(i, 7).Value

    End If

  Next i
  'To store the values of Greatest % increase, Greatest % decrease and 
  ' the greatest total volume
  ws.Range("Q4").Value = Max_Ticker
  ws.Range("R4").Value = Max_Change
  ws.Range("Q5").Value = Min_Ticker
  ws.Range("R5").Value = Min_Change
  ws.Range("Q6").Value = Max_Volume_Ticker
  ws.Range("R6").Value = Max_Volume
  ws.Range("R4").NumberFormat = "0.00%"
  ws.Range("R5").NumberFormat = "0.00%"
 ' Move to the next worksheet
 Next ws
End Sub