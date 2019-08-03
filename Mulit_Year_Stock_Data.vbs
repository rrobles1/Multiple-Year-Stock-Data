VBA HW *EASY*

Sub Stock_Data()

    Dim total_volume As Double
    total_volume = 0

    Dim ticker As String

    Dim ticker_row As Integer
    ticker_row = 2

    For i = 2 to 70926

        If Cells(i + 1,1).Value <> Cells(i,1).Value Then

        ticker = Cells(i,1).Value 

        total_volume = total_volume + Cells(i,7).Value
        Range("I" & ticker_row).Value = ticker
        Range("J" & ticker_row).Value = total_volume

        ticker_row = ticker_row + 1

        total_volume = 0
        Else 
        total_volume = total_volume + Cells(i,7).Value 
    
        End If

    Next i 

End Sub 