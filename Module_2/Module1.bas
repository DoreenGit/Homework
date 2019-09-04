Attribute VB_Name = "Module1"
Sub High_Low2()
Dim SheetYear As Integer
Dim Ticker As String
Dim NextTicker As String
Dim StockOpen As Double
Dim StockClose As Double
Dim StockHigh As Double
Dim StockLow As Double
Dim LastCol As Long
Dim LastRow As Long
Dim R As Long
Dim C As Long
Dim SummaryR As Long: SummaryR = 1
Dim SummaryC As Long
Dim SummaryTicker As String
Dim SummaryVolume As Double
Dim YearlyChange As Double
Dim PercentChange As Variant
Dim Rng As Range
Dim MaxPct As Double
Dim MinPct As Double
Dim MaxVolume As Variant

Dim MaxRow As Integer
Dim MinRow As Integer
Dim MaxVolRow As Integer

Dim MaxPctLR As String
Dim MinPctLR As String
Dim MaxVolLR As String





SheetCount = Application.Sheets.Count

For SheetYear = 1 To SheetCount
Worksheets(SheetYear).Activate
Columns("I:Q").Delete
'***** Find last column and row
With ActiveSheet
    LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column '*Find last column
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    SummaryC = LastCol + 2
    'SummaryC = LastCol + 2   '*Set summary 2 columns over from last column of data
End With

'****Write Summary Headers
    Cells(1, SummaryC) = "Ticker"
    Cells(1, SummaryC + 1) = "Yearly Change"
    Cells(1, SummaryC + 2) = "Percent Change"
    Cells(1, SummaryC + 3) = "Total Stock Volume"
    Cells(1, SummaryC + 7) = "Ticker"
    Cells(1, SummaryC + 8) = "Value"
    Cells(2, SummaryC + 6) = "Greatest % Increase"
    Cells(3, SummaryC + 6) = "Greatest % Decrease"
    Cells(4, SummaryC + 6) = "Greatest Total Volume"
    


NewTicker = True
SummaryR = 2
'ActiveSheet Loop
'For Sheet = 1 to
 
    '***Ticker Loop
     For R = 2 To LastRow
        '***Stock Open Loop
         Do Until Cells(R, 3).Value <> 0  '*Bypass any zero open price records
            R = R + 1
         Loop
        '*******************
         Ticker = Cells(R, 1).Value
         
         If NewTicker = True Then   '*If New Ticker
            StockOpen = Cells(R, 3).Value  '*Capture Stock Open Price
            NewTicker = False
         End If
         
         StockClose = Cells(R, 6).Value
         SummaryVolume = SummaryVolume + Cells(R, 7).Value
         NextTicker = Cells(R + 1, 1).Value
      

         If Ticker <> NextTicker And NextTicker <> "" Then
            SummaryTicker = Ticker  'Write Summary Ticker Name
            With ActiveSheet
                '*Write Ticker to summary
                .Cells(SummaryR, SummaryC).Value = Ticker
                '*Write Volume to summary
                .Cells(SummaryR, SummaryC + 3).Value = SummaryVolume
                '*Calculate Yearly Change
                YearlyChange = StockClose - StockOpen
                
                If YearlyChange > 0 Then
                    .Cells(SummaryR, SummaryC + 1).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    .Cells(SummaryR, SummaryC + 1).Interior.ColorIndex = 3
                Else
                    .Cells(SummaryR, SummaryC + 1).Interior.ColorIndex = xlNone
                End If
                
                .Cells(SummaryR, SummaryC + 1).Value = YearlyChange
                
                '*Calculate Percentage Change
                PercentChange = FormatPercent(((StockClose - StockOpen) / StockOpen))
                .Cells(SummaryR, SummaryC + 2).NumberFormat = "0.00%"
                .Cells(SummaryR, SummaryC + 2).Value = PercentChange
                '.Cells(SummaryR, SummaryC + 2).NumberFormat = "0.00000"
            End With
            SummaryR = SummaryR + 1
            NewTicker = True
            SummaryVolume = 0
            
         End If
        
Next R

'*****
'**

'Write Max and Mins
With Application.WorksheetFunction
   '*Max
    MaxPct = .Max(Range("K:K"))
    MaxPctLR = "K1:K" + Trim(Str(LastRow))
    MaxRow = Application.WorksheetFunction.Match(MaxPct, Range(MaxPctLR), 0)
    MaxTicker = Cells(MaxRow, SummaryC).Value
   
   '*Min
    MinPct = .Min(Range("K:K"))
    MinPctLR = "K1:K" + Trim(Str(LastRow))
    MinRow = Application.WorksheetFunction.Match(MinPct, Range(MinPctLR), 0)
    MinTicker = Cells(MinRow, SummaryC).Value
    
    MaxVolume = .Max(Range("L:L"))
    MaxVolLR = "L1:L" + Trim(Str(LastRow))
    MaxVolRow = Application.WorksheetFunction.Match(MaxVolume, Range(MaxVolLR), 0)
    MaxVolTicker = Cells(MaxVolRow, SummaryC).Value
    
    
'MaxPct = Application.WorksheetFunction.Max(Range("K:K"))
'MaxRow = Address(Match(Application.WorksheetFunction.Max(Range("K:K")),(K:K),0),1))
    'MinPct = Application.WorksheetFunction.Min(Range("K:K"))
    'MaxVolume = Application.WorksheetFunction.Max(Range("L:L"))
End With
Cells(2, SummaryC + 7) = MaxTicker
Cells(2, SummaryC + 8) = FormatPercent(MaxPct)
Cells(3, SummaryC + 7) = MinTicker
Cells(3, SummaryC + 8) = FormatPercent(MinPct)
Cells(4, SummaryC + 7) = MaxVolTicker
Cells(4, SummaryC + 8) = MaxVolume

ActiveSheet.Columns.AutoFit

Next SheetYear

End Sub
