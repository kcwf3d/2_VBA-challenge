Attribute VB_Name = "Module1"

Sub wkYearLoop():

    'varibales
    
    Dim Ticker As String
    Dim openVal As Double
    Dim closeVal As Double
    Dim toalSV As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    Dim curSheet As Worksheet
    Set curSheet = ThisWorkbook.Worksheets("2015")
    
    Range("K:K").NumberFormat = "0.00%"
    
    'Headers
    curSheet.Cells(1, 9).Value = "Ticker"
    curSheet.Cells(1, 10).Value = "Yearly Change"
    curSheet.Cells(1, 11).Value = "Precent Change"
    curSheet.Cells(1, 12).Value = "Total Stock Volume"

    
       'variables.2
     Dim i As Long
     Dim lastRow As Long
     
     lastRow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row
     groupNo = 1
     totalSV = 0
     openVal = curSheet.Cells(2, 3)
      
    'Loop
    For i = 2 To lastRow
        readName = curSheet.Cells(i, 1).Value
        nextName = curSheet.Cells(i + 1, 1).Value
        
        If nextName = readName Then
            totalSV = totalSV + curSheet.Cells(i, 7).Value
    
        Else
            closeVal = curSheet.Cells(i, 6).Value
            totalSV = totalSV + curSheet.Cells(i, 7).Value
            YearlyChange = closeVal - openVal
            On Error Resume Next
            PrecentChange = (closeVal - openVal) / openVal
            If Err.Number <> 0 Then
                prChange = 0
            End If
            groupNo = groupNo + 1
            curSheet.Cells(groupNo, 9).Value = readName
            curSheet.Cells(groupNo, 12).Value = totalSV
            curSheet.Cells(groupNo, 11).Value = PrecentChange
            curSheet.Cells(groupNo, 10).Value = YearlyChange
            totalSV = 0
            openVal = curSheet.Cells(i + 1, 3)
    
        End If
    Next i
    
    stLastRow = curSheet.Cells(curSheet.Rows.Count, 9).End(xlUp).Row
    
    For SummaryTable = 2 To stLastRow
        If curSheet.Cells(SummaryTable, 10) <= 0 Then
            curSheet.Cells(SummaryTable, 10).Interior.Color = RGB(255, 0, 0)
        Else
            curSheet.Cells(SummaryTable, 10).Interior.Color = RGB(0, 255, 0)
        End If
    Next SummaryTable
End Sub



