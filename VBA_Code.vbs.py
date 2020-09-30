Sub stocks()

' Assigning variables

    Dim r, i As Integer
    Dim open_value, close_value, yearly_change, percent_change As Double
    Dim start, LastRow As Double
    Dim total As Double

' Writing the headers for the table
    
    Range("I1").Select
    ActiveCell.Value = "Ticker Symbol"
    
    Range("J1").Select
    ActiveCell.Value = "Yearly Change"
 
    Range("K1").Select
    ActiveCell.Value = "Percent Change"

    Range("L1").Select
    ActiveCell.Value = "Total Stock Volume"

' Initializing variables

    ' Total Stock Value Addition
    total = 0
    
    ' Counter for Summary Table
    i = 0
    
    start = 2
    yearly_change = 0

' Looking for the end of the table
    ' LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastRow = Range("A1").End(xlDown).Row
    
' Iterating --starting the cycle
        For r = 2 To LastRow
        
    ' Ticker
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        ' Ticker_Symbol = Cells(i, 1).Value
    
        ' Ticker_Detail = Ticker_Detail + Cells(i, 7).Value
    
            ' Total
            total = total + Cells(r, 7).Value
            Range("I" & 2 + i).Value = Cells(r, 1).Value
            Range("L" & 2 + i).Value = total
            
            ' Yearly Change
            open_value = Cells(start, 3).Value
            close_value = Cells(r, 6).Value
            yearly_change = close_value - open_value
            Range("J" & 2 + i).Value = yearly_change
                
                'Adding Color
                If Range("J" & 2 + i).Value < 0 Then
                Range("J" & 2 + i).Interior.ColorIndex = 3
                End If
                
                If Range("J" & 2 + i).Value >= 0 Then
                Range("J" & 2 + i).Interior.ColorIndex = 4
                End If
                
            ' Percent Change
            If Cells(start, 3).Value = 0 Then
                percent_change = 0
                
            Else
            percent_change = Round((yearly_change / Cells(start, 3).Value * 100), 2)
            Range("K" & 2 + i).Value = percent_change & "%"
            
            End If
            
            ' Restart for new Ticker
            total = 0
            i = i + 1
            start = r + 1
        
        Else
            total = total + Cells(r, 7).Value
            
        End If

Next r
' End of cycle!


' Summary Table Challenge

' Assigning variables
    Dim LastSummaryRow As Integer
    LastSummaryRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "K").End(xlUp).Row
    Dim GIncrease, GDecrease, GTotalVolume As Double
    Dim Ticker_GIncrease, Ticker_GDecrease, Ticker_GTotalVolume As String

' Initializing variables
    GIncrease = 0
    GDecrease = 0
    GTotalVolume = 0

' Writing the headers and row names for the table
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

' Iterating --starting the cycle
    For x = 2 To CLng(LastSummaryRow)

        ' Greatest % Increase
        If GIncrease < CDbl(Cells(x, 11).Value) Then
            GIncrease = CDbl(Cells(x, 11).Value)
            Ticker_GIncrease = Cells(x, 9).Value
        End If
        Range("P2").Value = Ticker_GIncrease
        Range("Q2").Value = GIncrease
        
        ' Greatest % Decrease
        If GDecrease > CDbl(Cells(x, 11).Value) Then
            GDecrease = CDbl(Cells(x, 11).Value)
            Ticker_GDecrease = Cells(x, 9).Value
        End If
        Range("P3").Value = Ticker_GDecrease
        Range("Q3").Value = GDecrease
    
        ' Greatest % Greatest Total Volume
        If GTotalVolume < CLngLng(Cells(x, 12).Value) Then
            GTotalVolume = CLngLng(Cells(x, 12).Value)
            Ticker_GTotalVolume = Cells(x, 9).Value
        End If
        Range("P4").Value = Ticker_GTotalVolume
        Range("Q4").Value = GTotalVolume
        
    Next x

    'Autoadjust column width and format percent
    
        ActiveSheet.Columns("I:Q").AutoFit
        Range("Q2:Q3").Style = "Percent"
        Range("Q2:Q3").NumberFormat = "0.00%"


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

