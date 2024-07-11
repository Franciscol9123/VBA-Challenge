Attribute VB_Name = "Module1"
Sub CalculationsandFormating()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Speed up processing
        Application.ScreenUpdating = False
        
        ' Variable initialization
        Dim Ticker As String
        Dim Ticker_Volume As Double
        Dim Ticker_Open As Double
        Dim Ticker_Close As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        Dim LastRow As Long
        Dim Summary_Table_Row As Integer
        
        Ticker_Volume = 0
        Ticker_Open = 0
        Ticker_Close = 0
        Yearly_Change = 0
        Percent_Change = 0
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        
        ' Set column titles in row 1
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
       
        For i = 2 To LastRow
            ' Check for new ticker symbol
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                Ticker_Open = Cells(i, 3).Value
                Ticker_Volume = Cells(i, 7).Value
            End If
            
            ' Check for end of ticker data
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Ticker_Close = Cells(i, 6).Value
                Yearly_Change = Ticker_Close - Ticker_Open
                If Ticker_Open <> 0 Then
                    Percent_Change = Yearly_Change / Ticker_Open
                Else
                    Percent_Change = 0
                End If
                Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
                
                ' Output results
                Range("I" & Summary_Table_Row).Value = Ticker
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                Range("K" & Summary_Table_Row).Value = Percent_Change
                Range("L" & Summary_Table_Row).Value = Ticker_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset for next ticker
                Ticker_Volume = 0
            Else
                Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
            End If
        Next i

        ' Conditional Formatting
        Dim YearRange As Range
        Set YearRange = Range("J2:J" & Summary_Table_Row - 1)
        For Each Cell In YearRange
            If Cell.Value < 0 Then
                Cell.Interior.ColorIndex = 3
            ElseIf Cell.Value > 0 Then
                Cell.Interior.ColorIndex = 4
            End If
        Next

        ' Format Percent Change
        Range("K2:K" & Summary_Table_Row - 1).NumberFormat = "0.00%"

        ' Calculate greatest increase, decrease, and volume
        Dim Per_Inc As Double
        Dim Per_Dec As Double
        Dim Tot_Vol As Double
        Dim Search_Percent As Range
        Dim Search_Volume As Range
        Dim Tic_Inc As Long
        Dim Tic_Dec As Long
        Dim Tic_Tot As Long
        
        Set Search_Percent = Range("K2:K" & Summary_Table_Row - 1)
        Set Search_Volume = Range("L2:L" & Summary_Table_Row - 1)
        
        Per_Inc = Application.WorksheetFunction.Max(Search_Percent)
        Per_Dec = Application.WorksheetFunction.Min(Search_Percent)
        Tot_Vol = Application.WorksheetFunction.Max(Search_Volume)
        
        Cells(2, 16).Value = Per_Inc
        Cells(3, 16).Value = Per_Dec
        Cells(4, 16).Value = Tot_Vol
        
        Tic_Inc = WorksheetFunction.Match(Per_Inc, Search_Percent, 0)
        Tic_Dec = WorksheetFunction.Match(Per_Dec, Search_Percent, 0)
        Tic_Tot = WorksheetFunction.Match(Tot_Vol, Search_Volume, 0)
        
        Range("O2").Value = Cells(Tic_Inc + 1, 9).Value
        Range("O3").Value = Cells(Tic_Dec + 1, 9).Value
        Range("O4").Value = Cells(Tic_Tot + 1, 9).Value
        
        Range("P2:P3").NumberFormat = "0.00%"

        ' Restore screen updating
        Application.ScreenUpdating = True
    Next ws

    MsgBox "Data calculation and formatting completed."
End Sub

