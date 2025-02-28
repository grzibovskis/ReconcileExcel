Option Explicit

Sub FindAndHighlightMatches()

    '======================== Declare variables
    Dim lastRow As Long
    Dim NextRow As Long
    Dim messageQualifiers As Variant
    Dim Fndr As Range
    Dim qualifier As Variant
    Dim row As Long
    Dim Fndr2 As Range
    Dim paaRef As String
    Dim cellRow As Long
    Dim matching As Boolean

    '======================== Error Handling
    On Error GoTo ErrorHandler

    '======================== Find the last non-empty row in column I of the shSekFile sheet
    lastRow = shSekFile.Cells(shSekFile.Rows.Count, "I").End(xlUp).row
    NextRow = lastRow + 1 ' Calculate the next available row

    matching = False ' Initialize the matching flag as False

    '======================== List of MESSAGES qualifiers to search for
    messageQualifiers = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15")

    '======================== Find the row in shMain containing 6 qualifier in column I
    Set Fndr = shMain.Columns(9).Find(What:="6" & "*", LookAt:=xlWhole)
    If Not Fndr Is Nothing Then
        row = Fndr.row ' Get the row where 6 is found
    Else
        MsgBox "6 reference not found in shMain!", vbExclamation
        Exit Sub
    End If

    '======================== Extract 6 reference value from the found row
    paaRef = extractValue(shMain.Cells(row, "I").value, "6")

    '======================== Find the same 6 reference in the shSekFile sheet
    Set Fndr2 = shSekFile.Columns(9).Find(What:=paaRef & "*", LookAt:=xlWhole)
    If Not Fndr2 Is Nothing Then
        cellRow = Fndr2.row ' Get the row where the 6 reference is found in shSekFile
    Else
        MsgBox "6 reference not found in shSekFile!", vbExclamation
        Exit Sub
    End If

    '======================== Loop through each message qualifier in the array
    For Each qualifier In messageQualifiers
        '======================== Search for the qualifier in column I of the shMain sheet
        Set Fndr = shMain.Columns(9).Find(What:=qualifier & "*", LookAt:=xlWhole)
        matching = False ' Reset matching flag for each qualifier

        '======================== If qualifier is found, proceed
        If Not Fndr Is Nothing Then
            row = Fndr.row ' Get the row where the qualifier is found

            '======================== Process the qualifier by comparing values from shMain and shSekFile
            Select Case qualifier
                Case "1"
                    If shSekFile.Cells(cellRow, "A").value = extractValue(shMain.Cells(row, "I").value, "1") Then matching = True
                Case "2"
                    If shSekFile.Cells(cellRow, "C").value = extractDateValue(shMain.Cells(row, "I").value, "2") Then matching = True
                Case "3"
                    If shSekFile.Cells(cellRow, "D").value = extractDateValue(shMain.Cells(row, "I").value, "3") Then matching = True
                Case "4"
                    If shSekFile.Cells(cellRow, "E").value = extractDateValue(shMain.Cells(row, "I").value, "4") Then matching = True
                Case "5"
                    If shSekFile.Cells(cellRow, "F").value = processRateValue(shMain.Cells(row, "I").value) Then matching = True
                Case "6"
                    ' Extract and match part of the message content
                    If shSekFile.Cells(cellRow, "G").value = Mid(shMain.Cells(row, "I").value, InStr(1, shMain.Cells(row, "I").value, "777") + Len("777"), 12) Then matching = True
                Case "7"
                    If shSekFile.Cells(cellRow, "H").value = extractValue(shMain.Cells(row, "I").value, "7") Then matching = True
                Case "8"
                    If shSekFile.Cells(cellRow, "J").value = extractValue(shMain.Cells(row, "I").value, "8") Then matching = True
                Case "9"
                    If shSekFile.Cells(cellRow, "K").value = cleanSafeValue(shMain.Cells(row, "I").value) Then matching = True
                Case "10"
                    If shSekFile.Cells(cellRow, "N").value = extractValue(shMain.Cells(row, "I").value, "10") Then matching = True
                Case "11"
                    If shSekFile.Cells(cellRow, "O").value = cleanUnitValue(shMain.Cells(row, "I").value) Then matching = True
                Case "12"
                    If shSekFile.Cells(cellRow, "P").value = extractValue(shMain.Cells(row, "I").value, "11") Then matching = True
                Case "12"
                    If shSekFile.Cells(cellRow, "Q").value = extractValue(shMain.Cells(row, "I").value, "12") Then matching = True
                Case "13"
                    If shSekFile.Cells(cellRow, "R").value = extractValue(shMain.Cells(row, "I").value, "13") Then matching = True
                Case "14"
                    If shSekFile.Cells(cellRow, "S").value = extractValue(shMain.Cells(row, "I").value, "14") Then matching = True
                Case "15"
                    If shSekFile.Cells(cellRow, "T").value = cleanTaxrSecValue(shMain.Cells(row, "I").value) Then matching = True
            End Select

            '======================== If matching values are found, highlight the cell in green
            If matching Then
                shMain.Cells(row, "I").Interior.Color = vbGreen
            Else
                '======================== If no match, highlight the cell in red
                shMain.Cells(row, "I").Interior.Color = vbRed
            End If
        Else
            ' Output MsgBox information if the qualifier is not found
            MsgBox "Qualifier " & qualifier & " not found.", vbInformation
        End If
    Next qualifier ' Loop to the next qualifier

    Exit Sub ' Exit before error handler

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Err.Clear
End Sub
