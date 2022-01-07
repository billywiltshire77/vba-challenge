Attribute VB_Name = "Module1"
Sub vba_challenge():

    Dim startprice As Double
    Dim endprice As Double
    Dim yearchange As Double
    Dim yearvolume As Long
    Dim i As Long
    Dim lastrow As Long
    Dim uniquetickers As Long
    Dim percentchange As Double
    Dim ws_count As Integer
    Dim ws As Long
    Dim uniquecount As Long
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As Long
    
    
    ws_count = ActiveWorkbook.Worksheets.Count
        
    ' Loop through worksheets
    
    For ws = 1 To ws_count
        lastrow = ActiveWorkbook.Worksheets(ws).Cells(2, 1).End(xlDown).Row
        uniquecount = 0
        ActiveWorkbook.Worksheets(ws).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(ws).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(ws).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(ws).Cells(1, 12).Value = "Total Volume"
        
    ' Loop through each row of the raw data in the given worksheet
    
       For i = 2 To lastrow
            If ActiveWorkbook.Worksheets(ws).Cells(i, 1).Value <> ActiveWorkbook.Worksheets(ws).Cells(i - 1, 1) Then
                ActiveWorkbook.Worksheets(ws).Cells(uniquecount + 2, 9) = ActiveWorkbook.Worksheets(ws).Cells(i, 1).Value
                startprice = ActiveWorkbook.Worksheets(ws).Cells(i, 3).Value
                yearvolume = (ActiveWorkbook.Worksheets(ws).Cells(i, 7).Value / 1000)
                uniquecount = uniquecount + 1
            Else
                yearchange = ActiveWorkbook.Worksheets(ws).Cells(i, 6).Value - startprice
                yearvolume = yearvolume + (ActiveWorkbook.Worksheets(ws).Cells(i, 7).Value / 1000)
            End If
            
            If startprice = 0 Then
                ActiveWorkbook.Worksheets(ws).Cells(uniquecount + 1, 11).Value = 0
            Else
                percentchange = yearchange / startprice
                ActiveWorkbook.Worksheets(ws).Cells(uniquecount + 1, 11).Value = percentchange
            End If
            
            ActiveWorkbook.Worksheets(ws).Cells(uniquecount + 1, 10).Value = yearchange
            ActiveWorkbook.Worksheets(ws).Cells(uniquecount + 1, 12).Value = yearvolume
        Next i
    
    ' Format output cells
    
        uniquetickers = ActiveWorkbook.Worksheets(ws).Cells(2, 9).End(xlDown).Row
    
        For i = 2 To uniquetickers
            ActiveWorkbook.Worksheets(ws).Cells(i, 11).NumberFormat = "0.00%"
            If ActiveWorkbook.Worksheets(ws).Cells(i, 10).Value > 0 Then
                ActiveWorkbook.Worksheets(ws).Cells(i, 10).Interior.ColorIndex = 4
            Else
                ActiveWorkbook.Worksheets(ws).Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    
    ' Bonus Code
    
        greatestincrease = 0
        greatestdecrease = 0
        greatestvolume = 0
        ActiveWorkbook.Worksheets(ws).Cells(1, 16).Value = "Ticker"
        ActiveWorkbook.Worksheets(ws).Cells(1, 17).Value = "Value"
        ActiveWorkbook.Worksheets(ws).Cells(2, 15).Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(ws).Cells(3, 15).Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(ws).Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To uniquetickers
            If ActiveWorkbook.Worksheets(ws).Cells(i, 11).Value > greatestincrease Then
                greatestincrease = ActiveWorkbook.Worksheets(ws).Cells(i, 11).Value
                ActiveWorkbook.Worksheets(ws).Cells(2, 16).Value = ActiveWorkbook.Worksheets(ws).Cells(i, 9).Value
            ElseIf ActiveWorkbook.Worksheets(ws).Cells(i, 11).Value < greatestdecrease Then
                greatestdecrease = ActiveWorkbook.Worksheets(ws).Cells(i, 11).Value
                ActiveWorkbook.Worksheets(ws).Cells(3, 16).Value = ActiveWorkbook.Worksheets(ws).Cells(i, 9).Value
            ElseIf ActiveWorkbook.Worksheets(ws).Cells(i, 12).Value > greatestvolume Then
                greatestvolume = ActiveWorkbook.Worksheets(ws).Cells(i, 12).Value
                ActiveWorkbook.Worksheets(ws).Cells(4, 16).Value = ActiveWorkbook.Worksheets(ws).Cells(i, 9).Value
            End If
        Next i
        ActiveWorkbook.Worksheets(ws).Cells(2, 17).Value = greatestincrease
        ActiveWorkbook.Worksheets(ws).Cells(3, 17).Value = greatestdecrease
        ActiveWorkbook.Worksheets(ws).Cells(4, 17).Value = greatestvolume
        ActiveWorkbook.Worksheets(ws).Cells(2, 17).NumberFormat = "0.00%"
        ActiveWorkbook.Worksheets(ws).Cells(3, 17).NumberFormat = "0.00%"
        
    Next ws
End Sub
