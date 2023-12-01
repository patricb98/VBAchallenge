VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks()

'PART 1

'Loop through all sheets
Dim ws As Worksheet
For Each ws In Worksheets

'set variable for ticker
Dim Ticker As String

'Track location of ticker info in summary table
Dim summary_table_row As Integer
summary_table_row = 2

'set variable for holding total volume of stock ticker
Dim stockvolume As Double
stockvolume = 0

'set variables to calc yearly change
Dim yearchange As Double
Dim openingprice As Double
Dim closingprice As Double

'set opening price
openingprice = ws.Cells(2, 3).Value

'set variable for percent change
Dim percentchange As Double

'Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'find last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop through all stocks
For r = 2 To lastrow

'check we are still within the same stock, if it is not -> then
If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
    
    'set stock ticker
    Ticker = ws.Cells(r, 1).Value

    'Print stock ticker in summary table
    ws.Range("I" & summary_table_row).Value = Ticker
    
    'set closing price for stock
    closingprice = ws.Cells(r, 6).Value
    
    'set yearly change
    yearchange = (closingprice - openingprice)
    
    'print yearly change to summary table
    
    ws.Range("J" & summary_table_row).Value = yearchange
      
    'set colour formating
    
    If ws.Range("J" & summary_table_row).Value < 0 Then
        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        
    ElseIf ws.Range("J" & summary_table_row).Value > 0 Then
        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        
    Else
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 0
        
    End If
    
    'set percent change
    
    percentchange = ((closingprice - openingprice) / openingprice)
    
    'print percent change to summary table
    ws.Range("K" & summary_table_row).Value = percentchange
    
    'format percent change to %
    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
   
    'add to stock ticker volume
    stockvolume = stockvolume + ws.Cells(r, 7).Value

    'print stock ticker volume to summary table
    ws.Range("L" & summary_table_row).Value = stockvolume
    
    'rest volume total
    stockvolume = 0
    
    'add one to summary table row
    summary_table_row = summary_table_row + 1
    
    'reset opening price
    openingprice = ws.Cells(r + 1, 3).Value
    
    'if cell immediately following row is the same ticker
    Else
    
    'add to the stock volume total
    stockvolume = stockvolume + ws.Cells(r, 7).Value
    
    End If
    
Next r
    
'BONUS PART

'headers
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'find last row of percent change
lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row

'create loop to find values
For rr = 2 To lastrow2
            
            'find Greatest % Increase
            If ws.Cells(rr, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow2)) Then
                
                'find greatest increase % value
                ws.Range("P2").Value = ws.Cells(rr, 11).Value
                
                'format cell greatest increase % value
                ws.Range("P2").NumberFormat = "0.00%"
                
                'find tick greatest increase %
                ws.Range("O2").Value = ws.Cells(rr, 9).Value

            'find Greatest % Decrease
            ElseIf ws.Cells(rr, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow2)) Then
                
                'find greatest % decrease value
                ws.Range("P3").Value = ws.Cells(rr, 11).Value
                
                'format cell greatest % decrease
                ws.Range("P3").NumberFormat = "0.00%"
                
                'find tick greatest % decrease
                ws.Range("O3").Value = ws.Cells(rr, 9).Value
                
            'find Greatest Total Volume
            ElseIf ws.Cells(rr, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow2)) Then
                
                'find greatest total volume value
                ws.Range("P4").Value = ws.Cells(rr, 12).Value
                
                'find tick greatest total volume value
                ws.Range("O4").Value = ws.Cells(rr, 9).Value
            
            End If
        
        Next rr

Next ws

End Sub

