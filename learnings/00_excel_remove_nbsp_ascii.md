Sub CheckAndReplaceAsciiCodes()
    Dim ws As Worksheet
    Dim logSheet As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim asciiCodes As Variant
    Dim foundList As Long
    Dim charFound As String
    
    ' Define the ASCII codes to search for
    asciiCodes = Array(127, 129, 141, 143, 144, 157, 160) ' Add the ASCII codes you want to check
    
    ' Set the log sheet
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("LogSheet")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = "LogSheet"
    End If
    On Error GoTo 0
    
    ' Clear previous log
    logSheet.Cells.Clear
    logSheet.Range("A1:B1").Value = Array("Cell Address", "Found Character")
    foundList = 2 ' Start logging from the second row
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "LogSheet" Then
            Set rng = ws.UsedRange
            
            ' Loop through each cell in the used range
            For Each cell In rng
                If Not IsEmpty(cell.Value) Then
                    For i = LBound(asciiCodes) To UBound(asciiCodes)
                        charFound = Chr(asciiCodes(i))
                        
                        If InStr(1, cell.Value, charFound, vbBinaryCompare) > 0 Then
                            ' Log the cell address and found character
                            logSheet.Cells(foundList, 1).Value = ws.Name & "!" & cell.Address
                            logSheet.Cells(foundList, 2).Value = charFound
                            foundList = foundList + 1
                            
                            ' Replace the character with a space
                            cell.Value = Replace(cell.Value, charFound, " ")
                        End If
                    Next i
                End If
            Next cell
        End If
    Next ws
    
    MsgBox "Search and replace complete. Check 'LogSheet' for results.", vbInformation, "Task Completed"
End Sub
