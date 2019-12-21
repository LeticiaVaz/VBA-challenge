Attribute VB_Name = "Module1"
Sub StockHomework()
Dim Current As Worksheet
For Each ws In Worksheets
       Dim Lastrow As Long
       Dim ColumnA As String
       Dim summaryColumA As String
       Dim totalstock As Double
       Dim ClosePrice As Double
       Dim OpenPrice As Double
       Dim Variance As Double
       Dim vMin, vMax
       Dim RowNo As Long
       

       summaryColumA = 2
       ws.Range("i1").Value = "Ticker"
       ws.Range("j1").Value = "Yearly Change"
       ws.Range("l1").Value = "Total Stock Voume"
    Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    OpenPrice = ws.Cells(summaryColumA, 3).Value
    
    For i = 2 To Lastrow
                     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            ColumnA = ws.Cells(i, 1).Value
                            ClosePrice = ws.Cells(i, 6).Value
                            totalstock = totalstock + ws.Cells(i, 7).Value
                            Yearly_change = ClosePrice - OpenPrice
                            
                            
                            ' Calculate overall year percentage
                If OpenPrice <> 0 And ClosePrice <> 0 Then
                    Percentage = (ClosePrice / OpenPrice) - 1
                    ws.Range("k" & summaryColumA) = Percentage
                                     ws.Range("k" & summaryColumA).NumberFormat = "0.00%"
                Else
                    overallYearChangePercent = 0
                End If
                            
                            
                                ws.Range("j" & summaryColumA).Value = Yearly_change
                                If Yearly_change < 0 Then
                                    ws.Range("j" & summaryColumA).Interior.Color = RGB(255, 0, 0)
                                Else
                                    ws.Range("j" & summaryColumA).Interior.Color = RGB(124, 252, 0)
                                End If
                           
                            ws.Range("i" & summaryColumA).Value = ColumnA
                            ws.Range("l" & summaryColumA).Value = totalstock
    j = i + 1
    OpenPrice = ws.Cells(j, 3).Value
                            summaryColumA = summaryColumA + 1
                            totalstock = 0
                Else
                    totalstock = totalstock + ws.Cells(i, 7).Value
          End If
     Next
Next
End Sub



