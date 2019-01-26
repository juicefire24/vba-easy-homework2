VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockeasy()

    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "total volume"
    
  ' Set an initial variable for holding the ticker
  Dim ticker_label As String

  ' Set an initial variable for holding the total volume
  Dim total_volume As Double
  total_volume = 0
  
     LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all total_volume
  For i = 2 To LastRow

    ' Check if we are still within the ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
      ticker = Cells(i, 1).Value

      ' Add to the total_volume
      total_volume = total_volume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      Range("i" & Summary_Table_Row).Value = ticker

      ' Print the total_volume to the Summary Table
      Range("j" & Summary_Table_Row).Value = total_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker Total
      total_volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the total_volume
     total_volume = total_volume + Cells(i, 7).Value

    End If
    
    Next i
    
End Sub

