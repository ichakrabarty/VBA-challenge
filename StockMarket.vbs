Sub StockMarket()
    
   'Dim wb As Worksheet
    
   'For Each wb In Worksheets
    'Put headers for columns with results
    
    Dim header(1 To 4) As Variant
    header(1) = "ticker symbol"
    header(2) = "yearly change"
    header(3) = "percent change"
    header(4) = "total stock volume"
    
    Range("I1 : L1").Value = header()

    'Set variables for columns with results

    Dim ticker_symbol As String
    Dim year_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    
    'Variables for open and closing price of stock
    
    Dim open_price As Double
    Dim close_price As Double
    
    open_price = Cells(2, 3).Value
    
    'Set row variables

    Dim L_Row As Long
    Dim a As Long
    Dim b As Long

    L_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
    b = 2 'Row counter
    
    total_volume = 0 'Total Volume counter
    
    Dim day_volume As Long 'Variable for stock value in a given day
    'Dim day_volumeNext As Long

    'Create for loop to display all results need
    
    
    For a = 2 To L_Row
           
        'Display total stock volume
        
        day_volume = Range("G" & a).Value
        total_volume = day_volume + total_volume
        Range("L" & b).Value = total_volume
           
        If Cells(a + 1, 1).Value <> Cells(a, 1).Value Then
            
            'Display Ticker Symbols
            ticker_symbol = Cells(a, 1).Value
            Range("I" & b).Value = ticker_symbol
            
            'Display Yearly Change
            close_price = Cells(a, 6).Value
            year_change = close_price - open_price
            Range("J" & b).Value = year_change
            
            'Display Percent Change
            percent_change = (1 - (close_price / open_price))
            Range("K" & b).NumberFormat = "0.00%"
            Range("K" & b).Value = percent_change
        
            total_volume = 0 'Reset Total Stock Volume Sum
            
            open_price = Cells(a + 1, 3).Value
            b = b + 1
        End If
         
         'Color Cells in Yearly Change based on positive and negative change
        If Cells(b, 10).Value > 0 Then
            Cells(b, 10).Interior.ColorIndex = 4
        Else
            Cells(b, 10).Interior.ColorIndex = 3
            
        End If
    Next


End Sub

