Attribute VB_Name = "Module1"
Sub Stocks():
For Each ws In Worksheets

    Dim Ticker As String
    Dim VolumeTotal As Double
    Dim OpenNumber As Double
    Dim CloseNumber As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim GreatVolume As Double
    Dim GreatIncrease As Double
    Dim GreatDecrease As Double
    Dim Days As Double
    Dim SummaryTableRow As Integer
    
    VolumeTotal = 0
    OpenNumber = 0
    CloseNumber = 0
   OpenDate = 0
   CloseDate = 0
   GreatestIncrease = 0
   GreatDecrease = 0
   GreatVolume = 0
    SummaryTableRow = 2
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Precentage Change"
    ws.Cells(1, 12) = "Volume Total"
    ws.Cells(2, 14) = "Greatest % Increase"
    ws.Cells(3, 14) = "Greatest % Decrease"
    ws.Cells(4, 14) = "Greatest Total Volume"
    ws.Cells(1, 15) = "Ticker"
    ws.Cells(1, 16) = "Value"
    ws.Cells(1, 14) = " "
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To 797711
            If OpenNumber = 0 And CloseNumber = 0 Then
                OpenNumber = ws.Cells(i, 3)
                OpenDate = ws.Cells(i, 2)
                CloseNumber = ws.Cells(i, 6)
                CloseDate = ws.Cells(i, 2)
            End If
            If OpenDate > ws.Cells(i, 2) Then
                OpenDate = ws.Cells(i, 2)
                OpenNumber = ws.Cells(i, 3)
            End If
     
            If CloseDate < ws.Cells(i, 2) Then
                CloseDate = ws.Cells(i, 2)
                CloseNumber = ws.Cells(i, 6)
            End If
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
            YearChange = CloseNumber - OpenNumber
            PercentChange = ((CloseNumber - OpenNumber) / CloseNumber)
            ws.Range("I" & SummaryTableRow).Value = Ticker
            ws.Range("L" & SummaryTableRow).Value = VolumeTotal
              ws.Range("j" & SummaryTableRow).Value = YearChange
            ws.Range("k" & SummaryTableRow).Value = PercentChange
            SummaryTableRow = SummaryTableRow + 1
                
            VolumeTotal = 0
            OpenNumber = 0
            CloseNumber = 0
            YearChange = 0
            Days = 0
           If ws.Cells(SummaryTableRow, 11).Value > GreatIncrease Then
           GreatIncrease = ws.Cells(SummaryTableRow, 11).Value
           GreatTicker = ws.Cells(SummaryTableRow, 9).Value
     
        End If
        
        If ws.Cells(SummaryTableRow, 11).Value < GreatDecrease Then
           GreatDecrease = ws.Cells(SummaryTableRow, 11).Value
           LeastTicker = ws.Cells(SummaryTableRow, 9).Value
   
        End If
        
        If ws.Cells(SummaryTableRow, 12).Value > GreatVolume Then
           GreatVolume = ws.Cells(SummaryTableRow, 12).Value
           VolumeTicker = ws.Cells(SummaryTableRow, 9).Value
       
        End If
        Else
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
        End If
       
        Next i
    ws.Range("P2").Value = GreatIncrease
    ws.Range("O2").Value = GreatTicker
    ws.Range("P3").Value = GreatDecrease
    ws.Range("O3").Value = LeastTicker
    ws.Range("P4").Value = GreatVolume
    ws.Range("O4").Value = VolumeTicker
    lastRowsummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For i = 2 To lastRowsummary
       If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 2
        End If
    Next i
    ws.Range("K2:K" & lastRowsummary).NumberFormat = "0.00%"
Next ws

End Sub

