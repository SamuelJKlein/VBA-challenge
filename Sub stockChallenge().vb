Sub stockChallenge()


'variable to save current ticker
Dim CurrentTicker As String
    
'Variable to give each company a different row
Dim DisplayRow As Integer
    
'Variables for saving numbers for each company
Dim YearStart As Double
Dim YearEnd As Double
Dim YearlyChange As Double
Dim YearlyChangePercent As Double
Dim TotalVol As LongLong
TotalVol = 0
    
'Data begins on row 2
DisplayRow = 2



For Each ws In Worksheets

  
    
    
    'Get number of used rows
    Dim RowCount As LongLong
    RowCount = ws.Range("A2").End(xlDown).Row
    
 
    

    
    'Looping through all used rows
  
    Dim i As LongLong
    For i = 2 To RowCount
    TotalVol = TotalVol + ws.Cells(i, 7).Value
    
    'If a cell on column 1 is not equal to row above it we save the opening price as YearStart
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        YearStart = ws.Cells(i, 3).Value
        End If
        
    'If a cell on column 1 is not equal to row below it we save the closing price as YearEnd
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        YearEnd = ws.Cells(i, 6).Value
        
        'displaying all the data for the company we just looped through
        Worksheets(1).Cells(DisplayRow, 9).Value = ws.Cells(i, 1).Value
        YearlyChange = YearEnd - YearStart
        Worksheets(1).Cells(DisplayRow, 10).Value = YearlyChange
       
        YearlyChangePercent = (YearlyChange / YearStart)
        Worksheets(1).Cells(DisplayRow, 11).Value = YearlyChangePercent
         
        Worksheets(1).Cells(DisplayRow, 12).Value = TotalVol
        
        'changing the display row for the next company
        DisplayRow = DisplayRow + 1
        
        'Find greatest increase
        If YearlyChangePercent > Worksheets(1).Cells(2, 17).Value Then
        Worksheets(1).Cells(2, 17).Value = YearlyChangePercent
        Worksheets(1).Cells(2, 16).Value = ws.Cells(i, 1).Value
        End If
        
        'Find greatest decrease
        If YearlyChangePercent < Worksheets(1).Cells(3, 17).Value Then
        Worksheets(1).Cells(3, 17).Value = YearlyChangePercent
        Worksheets(1).Cells(3, 16).Value = ws.Cells(i, 1).Value
        End If
        
        'Find greatest total volume
        If TotalVol > Worksheets(1).Cells(4, 17).Value Then
        Worksheets(1).Cells(4, 17).Value = TotalVol
        Worksheets(1).Cells(4, 16).Value = ws.Cells(i, 1).Value
        End If
        
        'reset volume
        TotalVol = 0
        End If
    Next i
Next ws

 'FORMATTING
  
  'add column headers
    Worksheets(1).Cells(1, 9).Value = "Ticker"
    Worksheets(1).Cells(1, 10).Value = "Yearly Change"
    Worksheets(1).Cells(1, 11).Value = "Percent Change"
    Worksheets(1).Cells(1, 12).Value = "Total Stock Volume"
    
    Worksheets(1).Cells(2, 15).Value = "Greatest % Increase"
    Worksheets(1).Cells(3, 15).Value = "Greatest % Decrease"
    Worksheets(1).Cells(4, 15).Value = "Greatest Total Volume"
    Worksheets(1).Cells(1, 16).Value = "Ticker"
    Worksheets(1).Cells(1, 17).Value = "Value"
    
    'Set columns to auto fit contents
    Worksheets(1).Columns(10).AutoFit
    Worksheets(1).Columns(11).AutoFit
    Worksheets(1).Columns(12).AutoFit
    Worksheets(1).Columns(15).AutoFit
    Worksheets(1).Columns(16).AutoFit
    Worksheets(1).Columns(17).AutoFit
    
    'format row 11 as percentage
    Worksheets(1).Columns(11).NumberFormat = "0.00%"
    Worksheets(1).Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Get number of used rows for summary list
    Dim RowCount2 As Integer
    RowCount2 = Worksheets(1).Range("I2").End(xlDown).Row
    
    'Apply Conditional Formatting to columns J and K
    Range(Worksheets(1).Cells(2, 10), Worksheets(1).Cells(RowCount2, 11)).FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.ColorIndex = 4
    Range(Worksheets(1).Cells(2, 10), Worksheets(1).Cells(RowCount2, 11)).FormatConditions.Add(xlCellValue, xlLess, 0).Interior.ColorIndex = 3

End Sub