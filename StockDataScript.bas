Attribute VB_Name = "Module1"
Sub StockDataScript()

    'Declare the variables for the Chart
    Dim ticker_symbol As String
    Dim year_change As Double
    Dim percent_change As Double
    Dim Total As Double
    Dim opening_price As Double
    
    'BONUS: Declare the variables for the "Greatests" Charts
    Dim greatestPer_Up As Double
    Dim greatestPer_Down As Double
    Dim greatestTotal As Double
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    j = 2
    ticker_symbol = Cells(2, 1).Value
    opening_price = Cells(2, 3).Value
    
    'Enter Header Row Titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    
    For i = 2 To lastrow
    
        Total = Total + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Enter Ticker Symbol in table
            Cells(j, 9).Value = ticker_symbol
            
            'Calculate Yearly Change and enter in table
            year_change = Cells(i, 6).Value - opening_price
            Cells(j, 10).Value = year_change
            
            'Calculate Percent Change(=Yearly Change / Opening Price) and enter in Chart
            If opening_price = 0 Then
                Cells(j, 11).Value = 0
            Else
                percent_change = year_change / opening_price
                Cells(j, 11).Value = Format(percent_change, "Percent")
            End If
            'Enter Total in table
            Cells(j, 12).Value = Format(CStr(Total), "#,###")
            
            'Conditional Format: Yearly Change Cells Color Green if Positive, Red if Negative, no format for 0
            If year_change > 0 Then
                Cells(j, 10).Interior.ColorIndex = 50
            ElseIf year_change < 0 Then
                Cells(j, 10).Interior.ColorIndex = 22
            End If
            
            'Re-Intialize for next Ticker Symbol group
            ticker_symbol = Cells(i + 1, 1).Value
            opening_price = Cells(i + 1, 3).Value
            Total = 0
            j = j + 1
            
        End If
        
    Next i
    
    'BONUS Section
    
    'Enter Column and Row Titles
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volumne"
    
    'Declare Variables to keep track of "Greatests"
    lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
    greatestPer_Up = Cells(2, 11).Value
    greatestPer_Down = Cells(2, 11).Value
    greatestTotal = Cells(2, 12).Value
    
    greatestPer_UpT = Cells(2, 9).Value
    greatestPer_DownT = Cells(2, 9).Value
    greatestTotalT = Cells(2, 9).Value
    
    'For loop to iterate through the data table created previously with if conditions to check for greatest values
    For i = 2 To lastrow2
        
        If Cells(i + 1, 11).Value > greatestPer_Up Then
            greatestPer_Up = Cells(i + 1, 11).Value
            greatestPer_UpT = Cells(i + 1, 9).Value
        End If
        
        If Cells(i + 1, 11).Value < greatestPer_Down Then
            greatestPer_Down = Cells(i + 1, 11).Value
            greatestPer_DownT = Cells(i + 1, 9).Value
        End If
        If Cells(i + 1, 12).Value > greatestTotal Then
            greatestTotal = Cells(i + 1, 12).Value
            greatestTotalT = Cells(i + 1, 9).Value
        End If
     Next i
       
     'Enter values for "Greatests" table
     Range("P2").Value = greatestPer_UpT
     Range("Q2").Value = Format(greatestPer_Up, "Percent")
     Range("P3").Value = greatestPer_DownT
     Range("Q3").Value = Format(greatestPer_Down, "Percent")
     Range("P4").Value = greatestTotalT
     Range("Q4").Value = Format(greatestTotal, "#,###")
    
    Columns("I:Q").AutoFit
    
End Sub



