Sub RemoveSal2()


    Dim ws1 As Worksheet
    Set ws1 = ActiveSheet  

    ' Check if the sheet is blank
    If WorksheetFunction.CountA(ws1.UsedRange) = 0 Then
        MsgBox "Page is blank", vbInformation
    Else
    
    Dim lastRow As Long

    lastRow = Columns("A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Range("A1:A" & lastRow).Select
    
    Application.DisplayAlerts = False

    Dim ws As Worksheet
    Dim cell As Range
    Dim searchTexts As Variant
    Dim replacementText As String
    
    Set ws = ThisWorkbook.Sheets("Template")
    
    searchTexts = Array("   ", "   ", "  ", ", ", ", Mr. ", ", Ms. ", ",Mr. ", ",Ms. ", ",Mr.", ",Ms.", "Mr. ", "Ms. ", "Mr.", "Ms.", "Mister.", "Miss.", "Mister ", "Miss ", ". ", ".", "  ", "   ", "   ")
    
    replacementText = " "
    
    For Each cell In ws.UsedRange
        For Each searchText In searchTexts
            If InStr(1, cell.Value, searchText) > 0 Then
                cell.Value = Replace(cell.Value, searchText, replacementText)
            End If
        Next searchText
    Next cell

    Application.DisplayAlerts = True
    
    Dim cell1 As Range

    For Each cell1 In Selection

        If Not IsEmpty(cell1.Value) And Left(cell1.Value, 1) = " " Then
        cell1.Value = Mid(cell1.Value, 2)
        End If
    Next cell1

End If
End Sub

