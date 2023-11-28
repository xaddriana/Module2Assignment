Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim dataArray As Variant
    Dim i As Long
    Dim lastRow As Long
    Dim curSheet As Worksheet
    Dim maxIncP As Double
    Dim maxDecP As Double
    Dim maxTV As Double
    Dim maxTVT As String
    Dim maxIncT As String
    Dim maxDecT As String
    Dim groupNo As Long
    Dim totalSV As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    For Each curSheet In ActiveWorkbook.Worksheets
        lastRow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row
        If lastRow > 1 Then
            dataArray = curSheet.Range("A1:G" & lastRow).Value
            
            ' Initialize variables
            curSheet.Range("I1").Value = "Ticket"
            curSheet.Range("J1").Value = "Yearly Change"
            curSheet.Range("K1").Value = "Percent Change"
            curSheet.Range("L1").Value = "Total Stock Volume"
            
            maxIncP = 0
            maxDecP = 0
            maxTV = 0
            groupNo = 1
            totalSV = 0
            openPrice = dataArray(2, 3)
            
            For i = 2 To lastRow - 1
                readName = dataArray(i, 1)
                nextName = dataArray(i + 1, 1)
                
                If nextName = readName Then
                    totalSV = totalSV + dataArray(i, 7)
                Else
                    totalSV = totalSV + dataArray(i, 7)
                    closePrice = dataArray(i, 6)
                    
                    ' Output data to the sheet
                    curSheet.Cells(groupNo + 1, 9).Value = readName
                    curSheet.Cells(groupNo + 1, 12).Value = totalSV
                    
                    ' Check for maxTV
                    If totalSV > maxTV Then
                        maxTV = totalSV
                        maxTVT = readName
                    End If
                    
                    ' Calculate and output yearlyChange and percentChange
                    yearlyChange = closePrice - openPrice
                    percentChange = yearlyChange / openPrice
                    
                    curSheet.Cells(groupNo + 1, 10).Value = yearlyChange
                    curSheet.Cells(groupNo + 1, 11).Value = percentChange
                    
                    ' Check for maxIncP and maxDecP
                    If percentChange > maxIncP Then
                        maxIncP = percentChange
                        maxIncT = readName
                    End If
                    
                    If percentChange < maxDecP Then
                        maxDecP = percentChange
                        maxDecT = readName
                    End If
                    
                    ' Apply colors
                    If yearlyChange > 0 Then
                        curSheet.Cells(groupNo + 1, 10).Interior.Color = RGB(0, 255, 0)  ' Lime
                    Else
                        curSheet.Cells(groupNo + 1, 10).Interior.Color = RGB(255, 0, 0)  ' Bright Red
                    End If
                    
                    ' Reset variables for the next group
                    totalSV = 0
                    openPrice = dataArray(i + 1, 3)
                    groupNo = groupNo + 1
                End If
            Next i
            
            ' Output calculated values to the sheet
            curSheet.Range("O2").Value = "Greatest % Increase"
            curSheet.Range("O3").Value = "Greatest % Decrease"
            curSheet.Range("O4").Value = "Greatest Total Volume"
            curSheet.Range("P1").Value = "Ticker"
            curSheet.Range("Q1").Value = "Value"
            curSheet.Range("P2").Value = maxIncT
            curSheet.Range("P3").Value = maxDecT
            curSheet.Range("P4").Value = maxTVT
            curSheet.Range("Q2").Value = maxIncP
            curSheet.Range("Q3").Value = maxDecP
            curSheet.Range("Q4").Value = maxTV
            
            ' Format columns
            curSheet.Range("K1:K" & lastRow).NumberFormat = "0.00%"
            curSheet.Cells(2, 17).NumberFormat = "0.00%"
            curSheet.Cells(3, 17).NumberFormat = "0.00%"
        End If
    Next curSheet
End Sub


