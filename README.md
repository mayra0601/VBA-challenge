# VBA-challenge
Sub StockData()

' Set Ws
    Dim Ws As Worksheet
    For Each Ws In Worksheets
    
' Spreadsheet Headers
      Ws.Range("I1").Value = "Ticker"
      Ws.Range("J1").Value = "Yearly Change"
      Ws.Range("K1").Value = "Percent Change"
      Ws.Range("L1").Value = "Total Stock Volume"
      Ws.Range("P1").Value = "Ticker"
      Ws.Range("Q1").Value = "Value"
      Ws.Range("O2").Value = "Greatest % Increase"
      Ws.Range("O3").Value = "Greatest % Decrease"
      Ws.Range("O4").Value = "Greatest Total Volume"

 ' Variables for ticker name
    Dim Ticker As String
    Ticker = " "
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
        
  ' More Variables
    Dim OpenPrice As Double
    OpenPrice = 0
    Dim ClosePrice As Double
    ClosePrice = 0
    Dim YearlyChange As Double
    YearlyChange = 0
    Dim PercentageChange As Double
    PercentageChange = 0
    Dim MAXTICKER As String
    MAXTICKER = " "
    Dim MINTICKER As String
    MINTICKER = " "
    Dim MAXPERCENT As Double
    MAXPERCENT = 0
    Dim MINPERCENT As Double
    MINPERCENT = 0
    Dim MAXVOLUME_TICKER As String
    MAXVOLUME_TICKER = " "
    Dim MAXVOLUME As Double
    MAXVOLUME = 0
'----------------------------------------------------------------
         
  ' summary table and row count for worksheet
        Dim SummaryRow As Long
        SummaryRow = 2
        Dim Lastrow As Long
        Dim i As Long
       Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

        
  ' Ticker's starting point and loop
        OpenPrice = Ws.Cells(2, 3).Value
        For i = 2 To Lastrow
        
  ' ticker symbol
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                Ticker = Ws.Cells(i, 1).Value
                
   ' YearlyChange and PercentageChange
                ClosePrice = Ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice

   ' Percentage Change
                If OpenPrice <> 0 Then
                    PercentageChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
            
            TotalStockVolume = TotalStockVolume + Ws.Cells(i, 7).Value
            
            
  ' Color coding and yearly change info
    Ws.Range("J" & SummaryRow).Value = YearlyChange
    Ws.Range("I" & SummaryRow).Value = Ticker

                If (YearlyChange > 0) Then
                    Ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                ElseIf (YearlyChange <= 0) Then
                    Ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                End If
       
                
  ' Print the Ticker Name
                Ws.Range("K" & SummaryRow).Value = (CStr(PercentageChange) & "%")
                Ws.Range("L" & SummaryRow).Value = TotalStockVolume
                
   ' Increase summary table row
                SummaryRow = SummaryRow + 1

  ' Reset Yearly Change and PercentageChange
                YearlyChange = 0
                ClosePrice = 0
                OpenPrice = Ws.Cells(i + 1, 3).Value
          
   ' calculations in spreadsheet
                If (PercentageChange > MAXPERCENT) Then
                    MAXPERCENT = PercentageChange
                    MAXTICKER = Ticker
                ElseIf (PercentageChange < MINPERCENT) Then
                    MINPERCENT = PercentageChange
                    MINTICKER = Ticker
                End If
                       
                If (TotalStockVolume > MAXVOLUME) Then
                    MAXVOLUME = TotalStockVolume
                    MAXVOLUME_TICKER = Ticker
                End If
                
   ' Reset counters
                PercentageChange = 0
                TotalStockVolume = 0
                
            'Else Statement
            Else
                ' Increase the Total Ticker Volume
                TotalStockVolume = TotalStockVolume + Ws.Cells(i, 7).Value
            End If
      
        Next i

        
     Next Ws


End Sub
