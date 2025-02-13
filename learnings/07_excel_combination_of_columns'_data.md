Sub GenerateCombinationsDynamic()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim mainTable As ListObject
    Dim colCount As Long
    Dim rowCount As Long
    Dim outputRow As Long
    Dim outputRange As Range
    
    ' Set the worksheets
    Set ws1 = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your main table's sheet name
    Set ws2 = ThisWorkbook.Sheets("Sheet2") ' Change "Sheet2" to your output table's sheet name
    
    ' Set the main table by its name
    Set mainTable = ws1.ListObjects("MainTable") ' Change "MainTable" to your table's name
    
    ' Clear previous combinations on Sheet2
    ws2.Cells.ClearContents
    
    ' Set the output range starting cell on Sheet2
    Set outputRange = ws2.Range("A1")
    
    ' Initialize output row
    outputRow = 2
    
    ' Get the number of columns and rows
    colCount = mainTable.ListColumns.Count
    rowCount = mainTable.ListRows.Count
    
    ' Create the table header
    outputRange.Cells(1, 1).Value = "Combinations"
    
    ' Generate combinations
    Call GenerateCombinationsRecursive(mainTable, "", colCount, rowCount, 1, outputRow, outputRange)
    
    ' Create the output table
    Dim outputTable As ListObject
    Set outputTable = ws2.ListObjects.Add(xlSrcRange, ws2.Range("A1:A" & outputRow - 1), , xlYes)
    outputTable.Name = "OutputTable"
    
    ' Autofit the output column
    outputTable.Range.Columns.AutoFit
End Sub

Sub GenerateCombinationsRecursive(mainTable As ListObject, combination As String, colCount As Long, rowCount As Long, colIndex As Long, ByRef outputRow As Long, ByRef outputRange As Range)
    Dim i As Long
    Dim newCombination As String
    
    If colIndex > colCount Then
        ' Output the combination
        outputRange.Cells(outputRow, 1).Value = Mid(combination, 2) ' Remove the leading underscore
        outputRow = outputRow + 1
    Else
        For i = 1 To rowCount
            If Not IsEmpty(mainTable.DataBodyRange(i, colIndex)) Then
                newCombination = combination & "_" & mainTable.DataBodyRange(i, colIndex).Value
                Call GenerateCombinationsRecursive(mainTable, newCombination, colCount, rowCount, colIndex + 1, outputRow, outputRange)
            End If
        Next i
    End If
End Sub
