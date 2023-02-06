VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickers()

'worksheet as variable

'Dim headers() As Variant
'Dim MainWs As Worksheet
'Dim wb As Workbook

'header info?

'For Each MainWs In wb.Sheets
   ' With MainWs
   ' .Rows(1).Value = " "
  '  For i = LBound(headers()) To UBound(headers())
  '  .Cells(1, 1 + i).Value = headers(i)
    
'Next MainWs


'loop each worksheet


' Find last row in sheet
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'headers

Range("O2").Value = "greatest % Increase "
Range("O3").Value = "greatest % Decrease "
Range("O4").Value = "greatest Total Volume "
Range("I1").Value = "Ticker "
Range("J1").Value = "Yearly Change "
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
'Set Variables

Dim tickerName As String
tickerName = " "
Dim totalsv As Double
totalsv = 0
Dim firstOpenPrice As Double
firstOpenPrice = 0
Dim lastClosePrice As Double
lastClosePrice = 0
Dim yearChange As Double
yearChange = 0
Dim percentChange As Double
percentChange = 0
Dim maxTicker As String
maxTicker = " "
Dim minTicker As String
minTicker = " "
Dim greatestPercent As Double
greatestPercent = 0
Dim leastPercent As Double
leastPercent = 0
Dim greatestVolumeTicker As String
greatestVolumeTicker = " "
Dim greatestVolume As Double
greatestVolume = 0


'variable to hold the rows in the columns(for new columns added)
Dim stockrows As Integer
stockrows = 2 'first row to populate in new columns = 2

'capure first ticker first open
         firstOpenPrice = Cells(stockrows, 3).Value
    Cells(stockrows, 15).Value = firstOpenPrice
       
    
         

'loop throughthe rows and check for changes in the tickers from row 2 until the last row
For Row = 2 To lastRow
    
    'check for changes in ticker (comparison)
    If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value And Cells(Row + 1, 1).Value <> " " Then

        'set the ticker name
          tickerName = Cells(Row, 1).Value
          
          'capture last close
           lastClosePrice = Cells(stockrows, 6).Value
           
           'yearly change
          yearChange = lastClosePrice - firstOpenPrice
          
          'percent change (cant divide by 0)
          If firstOpenPrice <> 0 Then
             percentChange = (yearChange / firstOpenPrice)
          End If
          
          'add total stock volume
          totalsv = totalsv + Cells(Row, 7).Value
          
          
          
          
          'display ticker name on current row of total columns in column I
          Cells(stockrows, 9).Value = tickerName
          
           ' display the total on the current row of the total columns in column L
          Cells(stockrows, 12).Value = totalsv
          
          'display year change in column J
          
         Cells(stockrows, 10).Value = yearChange
         
    
          
          'displaypercent change in column K
          Cells(stockrows, 11).Value = percentChange
          
          'capture last close
           lastClosePrice = Cells(stockrows, 6).Value
           
          'display last close
          'Cells(stockrows, 14).Value = lastClosePrice
          
           'add 1 to the stock rowsand go to the next row
           
             firstOpenPrice = Cells(stockrows, 3).Value
          stockrows = stockrows + 1
          
          'first open
          firstOpenPrice = Cells(Row + 1, 3).Value
          
          'colors
         If (yearChange > 0) Then
         Cells(stockrows - 1, 10).Interior.ColorIndex = 4
         ElseIf (yearChange <= 0) Then
         Cells(stockrows - 1, 10).Interior.ColorIndex = 3
         End If
         
         


         
       If (percentChange > greatestPercent) Then
           greatestPercent = percentChange
          maxTicker = tickerName
       ElseIf (percentChange < leastPercent) Then
           leastPercent = percentChange
           minTicker = tickerName
       End If
        
       If (totalsv > greatestVolume) Then
       greatestVolume = totalsv
       greatestVolumeTicker = tickerName
       End If
        
       'reset
       percentChange = 0
       totalsv = 0
        
      Range("Q2").Value = (CStr(greatestPercent) & "%")
      Range("Q3").Value = (CStr(leastPercent) & "%")
      Range("P2").Value = maxTicker
       Range("P3").Value = minTicker
       Range("Q4").Value = greatestVolume
      Range("P4") = greatestVolumeTicker
      
        


   Else
      ' if there is no change in the ticker, keep adding to the total
       totalsv = totalsv + Cells(Row, 7).Value
    End If

Next Row

'value in cells

Range("O2").Value = "greatest % Increase "
Range("O3").Value = "greatest % Decrease "
Range("O4").Value = "greatest Total Volume "
End Sub
