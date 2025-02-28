Option Explicit

Function processRateValue(value As String) As String
    Dim startPosition As Long
    Dim rateValue As Double

    ' Find the position of "SEK" and trim the string from that point
    startPosition = InStr(1, value, "SEK")
    value = Trim(Mid(value, startPosition + 3))
    
    ' Find the next space and trim the string up to that point
    startPosition = InStr(1, value, " ")
    If startPosition > 0 Then value = Left(value, startPosition - 1)
    
    ' Convert the trimmed string to a numeric value if possible
    If IsNumeric(value) Then
        rateValue = CDbl(value)
        processRateValue = CStr(rateValue)       ' Return the rate value as a string
    Else
        MsgBox "Error: Invalid rate format"
        processRateValue = vbNullString
    End If
End Function

Function extractValue(inputString As String, prefix As String) As String
    Dim startPosition As Long
    
    ' Find the start position of the prefix and extract the value after it
    startPosition = InStr(1, inputString, prefix) + Len(prefix)
    extractValue = Trim(Mid(inputString, startPosition))
End Function

Function extractDateValue(inputString As String, prefix As String) As String
    Dim startPosition As Long
    Dim dateString As String
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String
    Dim spacePosition As Long
    
    ' Find the start position of the prefix and extract the date part after it
    startPosition = InStr(1, inputString, prefix) + Len(prefix)
    dateString = Mid(inputString, startPosition)
    
    ' Ensure to extract only the date part if there's additional text after the date
    spacePosition = InStr(1, dateString, " ")
    If spacePosition > 0 Then
        dateString = Left(dateString, spacePosition - 1)
    End If
    dateString = Replace(dateString, "/", vbNullString)
    
    ' Handle both formats YYYYMMDD and YYYY/MM/DD
    If Len(dateString) = 8 Then
        ' Extract year, month, and day
        yearPart = Left(dateString, 4)
        monthPart = Mid(dateString, 5, 2)
        dayPart = Right(dateString, 2)
        
        ' Combine into the desired format YYYY-MM-DD
        extractDateValue = yearPart & "-" & monthPart & "-" & dayPart
    Else
        ' If the input is not in expected format, return an empty string or handle error
        MsgBox "Error: Invalid date format"
        extractDateValue = vbNullString
    End If
End Function

Function cleanSafeValue(inputString As String) As String
    Dim safePosition As Long
    
    ' Find the position of "TTOS" and clean the value after it
    safePosition = InStr(1, inputString, "TTOS")
    If safePosition > 0 Then
        cleanSafeValue = Mid(inputString, safePosition + Len("TTOS"))
        cleanSafeValue = Replace(cleanSafeValue, "-", vbNullString)
        cleanSafeValue = Replace(cleanSafeValue, " ", vbNullString) ' Remove spaces
        If Left(cleanSafeValue, 2) = "10" Then
            cleanSafeValue = Mid(cleanSafeValue, 3)
        End If
    Else
        cleanSafeValue = vbNullString
    End If
End Function

Function cleanUnitValue(inputString As String) As String
    Dim unitPosition As Long
    
    ' Find the position of "PPKS" and clean the value after it
    unitPosition = InStr(1, inputString, "PPKS")
    If unitPosition > 0 Then
        cleanUnitValue = Mid(inputString, unitPosition + Len("PPKS"))
        cleanUnitValue = Replace(cleanUnitValue, ",", vbNullString)
    Else
        cleanUnitValue = vbNullString
    End If
End Function

Function cleanTaxrSecValue(inputString As String) As String
    Dim taxrPosition As Long
    
    ' Find the position of "CCC" and clean the value after it
    taxrPosition = InStr(1, inputString, "CCC")
    If taxrPosition > 0 Then
        cleanTaxrSecValue = Mid(inputString, taxrPosition + Len("CCC"))
        cleanTaxrSecValue = Replace(cleanTaxrSecValue, ",", vbNullString)
    Else
        cleanTaxrSecValue = vbNullString
    End If
End Function
