VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stockchecker()

' Loop to go through all sheets
For Each ws In Worksheets

        ' Print Titles to top row for all the worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
  
Dim Ticker As String
Dim Open_price As Double
Dim Close_price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total As Double
    Total = 0
Dim Starting_row As Double
    Starting_row = 2
       
Dim Table_Row As Integer
    Table_Row = 2
        
        
       ' last row of the sheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To Lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Table_Row).Value = Ticker
            Open_price = ws.Range("C" & Starting_row).Value
            Close_price = ws.Cells(i, 6).Value
            Yearly_Change = Close_price - Open_price
            ws.Range("J" & Table_Row).Value = Yearly_Change
    
            If Open_price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / Open_price
            End If
            
            ' format cells
            ws.Range("K" & Table_Row).NumberFormat = "0.00%"
                        
          
            ws.Range("K" & Table_Row).Value = Percent_Change
            Total = Total + ws.Cells(i, 7).Value
            ws.Range("L" & Table_Row).Value = Total
            
            'Bonus: color code cells based on yearly change
            If ws.Range("J" & Table_Row).Value > 0 Then
                ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                
            ElseIf ws.Range("J" & Table_Row).Value < 0 Then
                ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                
            ElseIf ws.Range("J" & Table_Row).Value = 0 Then
                ws.Range("J" & Table_Row).Interior.ColorIndex = 0
                
            End If
            
       Table_Row = Table_Row + 1
            
            ' Update starting row and reset
            Starting_row = i + 1
            Total = 0

        Else
        Total = Total + ws.Cells(i, 7).Value
        
        End If
    
      Next i
      
      'Find the last row of the percent change
      Lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        For i = 2 To Lastrow
        
        'greater % increase
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If
        
        'greater % decrease
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                
            End If
        
        'greater total vol
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If
            
            'format cells
            ws.Range("Q2:Q3").NumberFormat = "0.00%"

        Next i
      
      'move to next sheet and restart!
    Next ws

End Sub



