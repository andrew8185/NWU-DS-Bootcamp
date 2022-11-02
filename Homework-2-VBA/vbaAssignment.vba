Sub StockFormatter()

'Declare object variables
    Dim wb As Workbook
    
    Dim ws As Worksheet
    
'Where the values are pasted into
 'Range(tempstring).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True
    Dim Rng As Range

'loop through all the sheets
    For Each ws In Worksheets


' Add the word Ticker to the First Column Header
        ws.Cells(1, 9).Value = "Ticker"

' Add the word Yearly Change to the First Column Header
        ws.Cells(1, 10).Value = "Yearly Change"

' Add the word Percent Change to the First Column Header
       ws.Cells(1, 11).Value = "Percent Change"

' Add the word Total Stock Volume to the First Column Header
      ws.Cells(1, 12).Value = "Total Stock Volume"
       
       


Set wb = ThisWorkbook
'Set ws = ThisWorksheet
Set Rng = ws.Range("I2")

RowCount = Cells(Rows.Count, "A").End(xlUp).Row
tempstring = "A2:A" & CStr(RowCount)
Dim UniqueTicker As String


'Variable define type
Dim Endrow As Long
Dim startprice As Long
Dim TotalVolume As Double
Dim TickerCount As Integer
Dim BeginningPrice As Double
Dim EndingPrice As Double
Dim PriceChange As Double
Dim PercentChange As Double


Endrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Variable initialize
TickerCount = 2
UniqueTicker = ws.Cells(2, 1)
TotalVolume = ws.Cells(2, 7)
BeginningPrice = ws.Cells(2, 3)
EndPrice = ws.Cells(6, 6)



For x = 2 To Endrow

'Getting the yearly change
    If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
          
        EndPrice = ws.Cells(x, 6)
        PriceChange = EndPrice - BeginningPrice
        PercentChange = PriceChange / BeginningPrice
        
      
        
        'print to worksheet
        ws.Cells(TickerCount, 9) = UniqueTicker
        ws.Cells(TickerCount, 12) = TotalVolume
        ws.Cells(TickerCount, 10) = PriceChange
        ws.Cells(TickerCount, 11) = PercentChange
            If ws.Cells(TickerCount, 10).Value > 0 Then
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
            End If
            
        'RE-assign a uniqueticker
        UniqueTicker = ws.Cells(x + 1, 1)
        TotalVolume = ws.Cells(x + 1, 7)
        
        'Counter
        TickerCount = TickerCount + 1
        
        BeginningPrice = ws.Cells(x + 1, 3)
        
    Else
        
        TotalVolume = TotalVolume + ws.Cells(x + 1, 7)

 
    
    End If
    
    'Formatting
    
    
 
Next x

    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Columns("A:L").AutoFit

Next

End Sub



