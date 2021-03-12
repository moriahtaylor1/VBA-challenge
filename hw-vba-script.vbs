Sub stockyear():
    Dim CurrentWs As Worksheet
    For Each CurrentWs In Worksheets
        ' create header
        CurrentWs.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        
        ' ' create bonus table
        CurrentWs.Range("O2").Value = "Greatest % Increase"
        CurrentWs.Range("O3").Value = "Greatest % Decrease"
        CurrentWs.Range("O4").Value = "Greatest Total Stock Volume"
        CurrentWs.Range("P1").Value = "Ticker"
        CurrentWs.Range("Q1").Value = "Value"
        
        ' set a variable specifying the column of interest
        Dim col As Integer
        col = 1
        ' set a variable for specifying the row in the summary table
        Dim summRow As Long
        summRow = 1
        ' determine last row of data
        Dim lastRow As Long
        lastRow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
        ' create variables for storage
        Dim openVal As Double
        Dim closeVal As Double
        Dim yearChange As Double
        Dim perChange As Double
        Dim totalStock As Double
        
        ' ' bonus variables
        Dim max_perChange As Double
        Dim min_perChange As Double
        Dim max_totalStock As Double
        max_perChange = 0
        min_perChange = 0
        max_totalStock = 0
        Dim max_perChange_ticker As String
        Dim min_perChange_ticker As String
        Dim max_totalStock_ticker As String
        
        ' loop through rows in the column
        Dim i As Long
        For i = 2 To lastRow
            ' search for when value of previous ticker is different than current ticker
            If CurrentWs.Cells(i - 1, col).Value <> CurrentWs.Cells(i, col).Value Then
                ' go to next row in summary table
                summRow = summRow + 1
                ' add new ticker symbol
                CurrentWs.Cells(summRow, 9).Value = CurrentWs.Cells(i, 1).Value
                ' capture opening price
                openVal = CurrentWs.Cells(i, 3).Value
                ' reset totalStock
                totalStock = 0
            ' search for when value of next ticker is different than current ticker
            ElseIf CurrentWs.Cells(i + 1, col).Value <> CurrentWs.Cells(i, col).Value Then
                ' capture closing price
                closeVal = CurrentWs.Cells(i, 6).Value
                ' calculate yearly change
                yearChange = closeVal - openVal
                ' calculate percent change
                If openVal = 0 Then
                    perChange = NA
                Else
                    perChange = (yearChange / openVal) * 100
                End If
                ' store statistics in summary table
                CurrentWs.Cells(summRow, 10) = yearChange
                CurrentWs.Cells(summRow, 11) = (CStr(Round(perChange, 2)) & "%")
                ' format with color for positive vs negative change
                If yearChange > 0 Then
                    CurrentWs.Range("J" & summRow).Interior.ColorIndex = 4
                ElseIf yearChange < 0 Then
                    CurrentWs.Range("J" & summRow).Interior.ColorIndex = 3
                End If
                
                ' ' bonus: max % increase and max % decrease
                If perChange > max_perChange Then
                    max_perChange = perChange
                    max_perChange_ticker = CurrentWs.Cells(summRow, 9)
                ElseIf perChange < min_perChange Then
                    min_perChange = perChange
                    min_perChange_ticker = CurrentWs.Cells(summRow, 9)
                End If
                    
            End If
            ' update open value if it is 0
            If openVal = 0 And CurrentWs.Cells(i, 3).Value <> 0 Then
                openVal = CurrentWs.Cells(i, 3).Value
            End If
            ' add volume to total stock volume
            totalStock = totalStock + CurrentWs.Cells(i, 7).Value
            ' update total stock summary
            CurrentWs.Cells(summRow, 12).Value = totalStock
            
            ' ' bonus: max total stock value
            If totalStock > max_totalStock Then
                max_totalStock = totalStock
                max_totalStock_ticker = CurrentWs.Cells(summRow, 9).Value
            End If
            
            Next i
        ' insert values into bonus table
        CurrentWs.Range("P2:Q2") = Array(max_perChange_ticker, CStr(max_perChange & "%"))
        CurrentWs.Range("P3:Q3") = Array(min_perChange_ticker, CStr(min_perChange & "%"))
        CurrentWs.Range("P4:Q4") = Array(max_totalStock_ticker, max_totalStock)
        
    Next CurrentWs
End Sub

