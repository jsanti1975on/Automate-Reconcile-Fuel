Sub fuelActualToDeclaredValues()
    Dim ws As Worksheet
    Dim lastRowD As Long
    Dim lastRowE As Long
    Dim i As Long, j As Long
    Dim matchFound As Boolean
    Dim cellD As Range
    Dim cellE As Range
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet2") 

    ' Find the last row with data in columns D and E
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Loop through each cell in column D
    For i = 2 To lastRowD ' Assuming there is a header row, start from the second row
        Set cellD = ws.Cells(i, 4) ' Column D
        matchFound = False
        
        ' Check if the value in column D is found in column E
        For j = 2 To lastRowE
            Set cellE = ws.Cells(j, 5) ' Column E
            If cellD.Value = cellE.Value Then
                matchFound = True
                Exit For
            End If
        Next j
        
        ' Highlight cell in column D if no match is found
        If Not matchFound Then
            cellD.Interior.Color = RGB(255, 0, 0) ' Red
        Else
            cellD.Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' Loop through each cell in column E to highlight unmatched values
    For j = 2 To lastRowE
        Set cellE = ws.Cells(j, 5) ' Column E
        matchFound = False
        
        ' Check if the value in column E is found in column D
        For i = 2 To lastRowD
            Set cellD = ws.Cells(i, 4) ' Column D
            If cellE.Value = cellD.Value Then
                matchFound = True
                Exit For
            End If
        Next i
        
        ' Highlight cell in column E if no match is found
        If Not matchFound Then
            cellE.Interior.Color = RGB(255, 0, 0) ' Red
        Else
            cellE.Interior.ColorIndex = xlNone
        End If
    Next j

    ' Notify the user that the comparison is complete
    MsgBox "Comparison complete. Differences are highlighted in red.", vbInformation
End Sub

Sub VerifyFuelTransactions()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fuelType As String
    Dim amount As Double
    Dim gallons As Double
    Dim expectedAmount As Double
    Dim tolerance As Double
    Dim difference As Double
    
    tolerance = 0.2 ' Define tolerance
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") 

    ' Find the last row with data in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Loop through the rows in the data 
    For i = 2 To lastRow
        ' Get fuel type, amount, and gallons
        fuelType = ws.Cells(i, 3).Value
        amount = ws.Cells(i, 5).Value
        gallons = ws.Cells(i, 6).Value

        ' Check if the fuel type is unknown
        If fuelType <> "REC" And fuelType <> "AV" Then
            ClearCellBackground ws, i, 5
            ws.Cells(i, 8).Value = "" ' Clear any existing value in column H
            GoTo ContinueLoop ' Skip this row and continue to the next one
        End If

        ' Calculate the expected amount based on the fuel type
        If fuelType = "REC" Then
            expectedAmount = 5.65 * gallons
        ElseIf fuelType = "AV" Then
            expectedAmount = 6.5 * gallons
        End If

        ' Calculate the difference
        difference = amount - expectedAmount

        ' Output the difference in column G
        ws.Cells(i, 7).Value = difference

        ' Check if the actual amount is within the tolerance (+/- 0.20)
        If Abs(difference) > tolerance Then
            ClearCellBackground ws, i, 5
            ws.Cells(i, 5).Interior.Color = RGB(255, 255, 0)
            ws.Cells(i, 8).Value = "Discrepancy"
        Else
            ws.Cells(i, 8).Value = "" ' Overwrite with nothing
        End If

ContinueLoop:
    Next i

    ' Notify the user that the verification is complete
    MsgBox "Verification complete. Discrepancies are highlighted in yellow.", vbInformation
End Sub

Sub ClearCellBackground(ws As Worksheet, row As Long, col As Long)
    ws.Cells(row, col).Interior.ColorIndex = xlNone
End Sub