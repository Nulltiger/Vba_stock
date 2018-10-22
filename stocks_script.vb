Sub stock()
        Dim Total_Volume as Double
        Dim Ticker_Name as String

    'Set an initial variable for volume'

    Total_Volume = 0

    'loop through all Stock tickers'
    For i = 1 to 70000

    ' Check if we are still within the same ticker, if it not. . .'
    If Cells(i +1, 1).Value<> Cells(i, 1).Value Then

        ' Set the Ticker name'
        Ticker_Name = Cells(i, 1).Value

        'Print the Ticker name in the I column'
        Cells(i, 9).Value = Total_Volume

        ' Reset the Total Volume'
        Total_Volume = 0

    'If the cell immediately following a row is the same Ticker. . .'
    Else

        ' Add to the Total_Volume'
        Total_Volume = Total_Volume + Cells(i, 7).Value

    End If
Next i

End Sub