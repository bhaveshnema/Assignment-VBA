Attribute VB_Name = "Module1"


Sub Button_Run_Multiyear_Stock_for_Workbook()

Dim ws As Worksheet

' Looping through workbook to retrieve worksheets
For Each ws In ThisWorkbook.Sheets
    Call Multiyear_Stock(ws)
Next

End Sub


Sub Multiyear_Stock(ws As Worksheet)

' Declare Variables
' =================================================================
  ' Set an initial variable for holding the Ticker Symbol
    Dim Ticker_Symbol As String

  ' Set an initial variable for holding the total volume
    Dim Total_Volume As Double
    Total_Volume = 0

  ' Set an initial variable for holding the Yearly and Percentage Change
    Dim Yearly_Change As Single
    Dim Percentage_Change As Single
    ' Yearly_Change = 0

  ' Set open and close price variables for yearly change
    Dim Open_Price As Single
    Dim Close_Price As Single
    Open_Price = 0
      
  ' Keep track of the location for each Ticker Symbol in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

  ' Define and determine the Last Row
    Dim Last_Row As Double
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row


' Main Code
' =====================================================================
  ' Loop through all rows
  ' ------------------------
    Dim i As Double
    For i = 2 To Last_Row
  ' ------------------------

    ' Loop to store Open price
    ' -------------------------------------
      If Open_Price = 0 Then
         Open_Price = ws.Cells(i, 3).Value
      End If
    ' -------------------------------------

    ' Check if we are still within the same Ticker Symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      Ticker_Symbol = ws.Cells(i, 1).Value

      ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

      ' Print the ticker symbol in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the Total Volume to the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = Total_Volume

      ' Store close price and calculate yearly change
      Close_Price = ws.Cells(i, 6).Value
      Yearly_Change = Close_Price - Open_Price
      ws.Range("K" & Summary_Table_Row).Value = Yearly_Change

      ' Conditional formatting
      ' -------------------------------------
        If Yearly_Change < 0 Then
          ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        Else
          ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
      ' --------------------------------------

      ' Calculate the percentage change
      ' --------------------------------------
      If Open_Price = 0 Then
          ws.Range("L" & Summary_Table_Row).Value = "NA"
          ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 6
      Else
      Percentage_Change = Yearly_Change / Open_Price
      ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
      ws.Range("L" & Summary_Table_Row).Value = Percentage_Change
      
      End If
      ' --------------------------------------
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the values
      Total_Volume = 0
      Open_Price = 0

    ' If the cell immediately following a row is the same Ticker Symbol...
    Else

      ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    End If

  Next i

End Sub

Sub Button_Run_Multiyear_Stock_for_Single()
 
     Dim target As Worksheet
     Set target = ActiveSheet
     Call Multiyear_Stock(target)
    
End Sub



