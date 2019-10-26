Sub StockData()

'Start worksheet loop
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
  ws.Activate
  
  'Add column headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("O1").Value = "Ticker"
  ws.Range("P1").Value = "Value"
  
  'Add row labels
  ws.Range("N2").Value = "Greatest % Increase"
  ws.Range("N3").Value = "Greatest % Decrease"
  ws.Range("N4").Value = "Greatest Total Volume"

  
  ' Set an initial variable for holding the ticker name
  Dim Ticker As String

  ' Set an initial variable for holding the total volume per ticker
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Set variable to track current ticker name
  Dim CurrentTicker As String
  CurrentTicker = ""
  
  ' Set variable for maximum percent change
  Dim MaxPercent As Double
  MaxPercent = 0
  
  ' Set variable for minimum change
  Dim MinPercent As Double
  MinPercent = 0
  
  ' Set variable for greatest total volume
  Dim MaxVolume As Double
  MaxVolume = 0
  
  ' Set variable for CurrentTicker open value
  Dim OpenValue As Double
  
       ' Loop through all stock volume
    For I = 2 To lastrow
        
        ' Add to the Volume Total
        Volume_Total = Volume_Total + ws.Cells(I, 7).Value
        
        'Determine if on first row of ticker
        If ws.Cells(I, 1).Value <> CurrentTicker Then
            
            OpenValue = ws.Cells(I, 3).Value
            CurrentTicker = ws.Cells(I, 1).Value
                    
        ' Determine if on last row of ticker
        ElseIf ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then

            ' Set the Ticker name
            Ticker = ws.Cells(I, 1).Value

            ' Print the Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
    
            ' Print the Volume Amount to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
            
            If Volume_Total > MaxVolume Then
            
            MaxVolume = Volume_Total
            
            ws.Range("P4").Value = MaxVolume
            
            ws.Range("O4").Value = CurrentTicker
            
            End If
            
            ' Print the Yearly Change to the Summary Table
            ' Get closing price
            Dim ClosingValue As Double
            ClosingValue = ws.Cells(I, 6).Value
            
            ' Print annual price change
            ws.Range("J" & Summary_Table_Row).Value = ClosingValue - OpenValue
            
            If OpenValue > 0 Then
            Dim PercentChange As Double
            PercentChange = ws.Cells(Summary_Table_Row, 10).Value / OpenValue
                ' Print percent change
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                If PercentChange > MaxPercent Then
                    
                    MaxPercent = PercentChange
                    ws.Range("P2").Value = MaxPercent
                    ws.Range("O2").Value = CurrentTicker
                
                ElseIf PercentChange < MinPercent Then
                
                    MinPercent = PercentChange
                    ws.Range("P3").Value = MinPercent
                    ws.Range("O3").Value = CurrentTicker
                    
                End If
                
                           
            ' If OpenValue divided by is 0
            ElseIf OpenValue = 0 Then
            ws.Range("K" & Summary_Table_Row).Value = 0
            
            End If
            
            ' Set conditional color formatting
            If ws.Range("K" & Summary_Table_Row).Value > 0 Then
            
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            
            ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
            
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
          
            ' Reset the Volume Total
            Volume_Total = 0
        
        End If

        
  Next I
  
  'Set formatting for Greatest table
    ws.Range("P2").Style = "Percent"
    ws.Range("P2").NumberFormat = "0.00%"
    
    ws.Range("P3").Style = "Percent"
    ws.Range("P3").NumberFormat = "0.00%"
    
  'Auto fit data into columns
  ws.Columns("A:P").EntireColumn.AutoFit
  
Next

starting_ws.Activate 'activate the worksheet that was originally active


End Sub