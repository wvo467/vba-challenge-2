Sub vbaChallenge():

    Dim Ticker_Name as String
    Ticker_Name = " "
    Dim Total_Ticker_Volume as Double
    Total_Ticker_Volume = 0
    Dim Open_Price as Double
    Open_Price = 0
    Dim Close_Price as Double
    Close_Price = 0 

    range("I1").value = "Ticker"
    range("J1").value = "Yearly Change"
    range("K1").value = "Percent Change"
    range("L1").value = "Total Stock Volume"

    last_row = WS.Cells(rows.count, 1).End(xlUp).row

    For i = 2 to last_row
        if cells(i + 1, 1).value <> cells(i, 1).value then
        Ticker_Name = cells(i, 1).value
        Close_Price = cells(i, 6).value

        end if

    For i = 2 to last_row
        if range("J1:Last_row" + 1).value > 0 then


        

End Sub