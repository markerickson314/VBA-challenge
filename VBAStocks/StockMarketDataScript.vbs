Attribute VB_Name = "Module1"
Sub multiYearStockData()


' Loop through all sheets
Dim ws As Worksheet
For Each ws In Worksheets

' Lastrow function
Dim Lastrow As Long
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'  Set variable for ticker symbol
Dim Ticker As String

' Define price variables
Dim Open_Price As Double
Dim Close_Price As Double
Dim Price_Change As Double
Dim Percent_Change As Double

' Hold total volume
Dim Volume_Total As LongLong
Volume_Total = 0


' Keep track of each ticker symbol in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


' Loop through rows
For i = 2 To Lastrow

    ' If ticker symbol is different from above
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Set Open_Price
        Open_Price = ws.Cells(i, 3).Value
      
    End If
    
    ' If ticker symbol is different from below then
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
        ' Set the Ticker
        Ticker = ws.Cells(i, 1).Value
        
        ' Set ClosePrice
        Close_Price = ws.Cells(i, 6).Value
        
        ' Add to volume
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
        ' Print Ticker to Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        ' Print Price Change for year to Summary Table
        Price_Change = Close_Price - Open_Price
        ws.Range("J" & Summary_Table_Row).Value = Price_Change
        
            ' Conditional Format Percent_Change
            If Price_Change < 0 Then
            
                 ' Fill Cell Red
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            Else
        
                ' If not, fill cell green
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            End If
        
            'If Open Price equals zero
            If Open_Price = 0 Then
            
                'Set Percent_Change to zero
                Percent_Change = 0
        
            Else
                
                'Calculate Percent Change
                Percent_Change = Price_Change / Open_Price
            
            End If
            
        ' Print Percent change for year to Summary Table
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
        ' Print Brand Total to Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Volume_Total
        
        ' Add one to Summary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the Volume Total
        Volume_Total = 0
         
    ' If ticker symbol remains the same
    Else
    
         ' Add to Volume Total
         Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
    End If


Next i

' Print title row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' Format column
ws.Range("K2:K" & Lastrow).NumberFormat = "0.00%"
ws.Columns("I:L").AutoFit

Next ws

End Sub
