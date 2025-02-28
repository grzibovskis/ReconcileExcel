'###########################################################################################################################
'###################################################### -- SEK File preparer -- ############################################
'###########################################################################################################################
Option Explicit

Sub getValues()

    '======================== Declare variables
    Dim lastRow As Long
    Dim NextRow As Long
    Dim messageQualifiers As Variant
    Dim Fndr As Range
    Dim qualifier As Variant
    Dim row As Long
    
    ' Enable error handling to catch any runtime errors
    On Error GoTo ErrorHandler

    ' Find the last non-empty row in column A of the shSekFile sheet
    lastRow = shSekFile.Cells(shSekFile.Rows.Count, "A").End(xlUp).row
    NextRow = lastRow + 1 ' Calculate the next available row

    ' List of Message qualifiers to search for
    messageQualifiers = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15")

    '======================== Loop through each qualifier
    For Each qualifier In messageQualifiers
        ' Error handling within the loop
        On Error Resume Next ' Skip to the next statement if an error occurs

        ' Find the qualifier in column A of shMain
        Set Fndr = shMain.Columns(1).Find(What:=qualifier & "*", LookAt:=xlWhole)
        
        '======================== If the qualifier is found, get the row number
        If Not Fndr Is Nothing Then
            row = Fndr.row ' Get the row where the qualifier is found

            ' Process each value based on the found row
            Select Case qualifier
                Case "1"
                    shSekFile.Cells(NextRow, "G").value = Mid(shMain.Cells(row, "A").value, InStr(1, shMain.Cells(row, "A").value, "777") + Len("777"), 12)
                Case "2"
                    shSekFile.Cells(NextRow, "J").value = extractValue(shMain.Cells(row, "A").value, "NAME1")
                Case "3"
                    shSekFile.Cells(NextRow, "H").value = extractValue(shMain.Cells(row, "A").value, "NAME2")
                Case "4"
                    shSekFile.Cells(NextRow, "A").value = extractValue(shMain.Cells(row, "A").value, "NAME3")
                Case "5"
                    shSekFile.Cells(NextRow, "K").value = cleanSafeValue(shMain.Cells(row, "A").value)
                Case "6"
                    shSekFile.Cells(NextRow, "O").value = cleanUnitValue(shMain.Cells(row, "A").value)
                Case "7"
                    shSekFile.Cells(NextRow, "C").value = extractDateValue(shMain.Cells(row, "A").value, "NAME4")
                Case "8"
                    shSekFile.Cells(NextRow, "D").value = extractDateValue(shMain.Cells(row, "A").value, "NAME5")
                Case "9"
                    shSekFile.Cells(NextRow, "P").value = extractValue(shMain.Cells(row, "A").value, "NAME6")
                Case "10"
                    shSekFile.Cells(NextRow, "Q").value = extractValue(shMain.Cells(row, "A").value, "NAME7")
                Case "11"
                    shSekFile.Cells(NextRow, "R").value = extractValue(shMain.Cells(row, "A").value, "NAME8")
                Case "12"
                    shSekFile.Cells(NextRow, "S").value = extractValue(shMain.Cells(row, "A").value, "NAME9")
                Case "13"
                    shSekFile.Cells(NextRow, "E").value = extractDateValue(shMain.Cells(row, "A").value, "NAME10")
                Case "14"
                    shSekFile.Cells(NextRow, "F").value = processRateValue(shMain.Cells(row, "A").value)
                Case "15"
                    shSekFile.Cells(NextRow, "T").value = cleanTaxrSecValue(shMain.Cells(row, "A").value)
            End Select

            ' If an error occurred inside the case block, handle it
            If Err.Number <> 0 Then
                MsgBox "An error occurred while processing qualifier: " & qualifier, vbExclamation
                Err.Clear ' Clear the error after handling
            End If

        Else
            '======================== Output if the qualifier is not found
            MsgBox "Qualifier " & qualifier & " not found.", vbInformation
        End If
    Next qualifier

    ' Exit subroutine after successful execution
    Exit Sub

ErrorHandler:
    ' Handle unexpected errors
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    Err.Clear

End Sub


Sub CompareAndPrintValues()
    '======================== Declare variables
    Dim lastRowSekA As Long
    Dim lastRowDbE As Long
    Dim lastRowDbC As Long
    Dim cellLoop As Long
    Dim cellLoopDatabase As Long
    Dim valuePrinted As Boolean
    Dim sekValueA As String
    Dim sekValueG As String

    ' Enable error handling to catch any runtime errors
    On Error GoTo ErrorHandler

    '======================== Find the last rows in SEK file column A and G
    lastRowSekA = shSekFile.Cells(shSekFile.Rows.Count, "A").End(xlUp).row

    '======================== Find the last rows in Database sheet column E and C
    lastRowDbE = shDatabase.Cells(shDatabase.Rows.Count, "E").End(xlUp).row
    lastRowDbC = shDatabase.Cells(shDatabase.Rows.Count, "C").End(xlUp).row

    '======================== Loop through SEK file sheet values in column A
    For cellLoop = 2 To lastRowSekA
        ' Get value from SEK file column A
        sekValueA = shSekFile.Cells(cellLoop, "A").value

        ' Loop through Database sheet values in column E to find a match
        For cellLoopDatabase = 2 To lastRowDbE
            ' Error handling within the loop
            On Error Resume Next
            If sekValueA = shDatabase.Cells(cellLoopDatabase, "E").value Then
                ' If there is a match, print value from Database sheet column F to SEK file column B
                shSekFile.Cells(cellLoop, "B").value = shDatabase.Cells(cellLoopDatabase, "F").value
                Exit For                         ' Exit the loop once a match is found
            End If
            ' Error handling for potential match issue
            If Err.Number <> 0 Then
                MsgBox "Error comparing values in column A and E. Row: " & cellLoop, vbExclamation
                Err.Clear
            End If
        Next cellLoopDatabase
    Next cellLoop

    '======================== Loop through SEK file sheet values in column G
    For cellLoop = 2 To lastRowSekA
        ' Reset valuePrinted flag
        valuePrinted = False

        ' Get value from SEK file column G
        sekValueG = shSekFile.Cells(cellLoop, "G").value

        ' Loop through Database sheet values in column C to find a match
        For cellLoopDatabase = 2 To lastRowDbC
            ' Error handling within the loop
            On Error Resume Next
            If sekValueG = shDatabase.Cells(cellLoopDatabase, "C").value Then
                ' If there is a match, print value from Database sheet column B to SEK file column N
                shSekFile.Cells(cellLoop, "N").value = shDatabase.Cells(cellLoopDatabase, "B").value
                ' Print value from Database sheet column A to SEK file column L
                shSekFile.Cells(cellLoop, "L").value = shDatabase.Cells(cellLoopDatabase, "A").value
                ' Set valuePrinted flag to true
                valuePrinted = True
                Exit For                         ' Exit the loop once a match is found
            End If
            ' Error handling for potential match issue
            If Err.Number <> 0 Then
                MsgBox "Error comparing values in column G and C. Row: " & cellLoop, vbExclamation
                Err.Clear
            End If
        Next cellLoopDatabase

        If valuePrinted Then
            ' Error handling for setting values
            On Error Resume Next
            shSekFile.Cells(cellLoop, "M").value = shDatabase.Cells(1, "I").value
            If Err.Number <> 0 Then
                MsgBox "Error printing value from Database to SEK file. Row: " & cellLoop, vbCritical
                Err.Clear
            End If
        End If
    Next cellLoop

    ' Exit subroutine successfully
    Exit Sub

ErrorHandler:
    ' Handle unexpected errors
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    Err.Clear
End Sub


'###########################################################################################################################
'###################################################### -- processed further -- #####################
'###########################################################################################################################

Sub createProcess()
    '======================== Declare variables
    Dim desktopPath As String
    Dim folderName As String
    Dim fileName As String
    Dim filePath As String
    Dim lastRowSekFile As Long
    Dim seFileWorkbook As Workbook
    Dim seFileWorksheet As Worksheet
    Dim answer As Integer
    Dim cellLoop As Long
    Dim targetRow As Long

    ' Enable error handling to catch any runtime errors
    On Error GoTo ErrorHandler

    '======================== Get the user's desktop path
    On Error Resume Next
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    If Err.Number <> 0 Then
        MsgBox "Error getting the desktop path.", vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrorHandler ' Resume general error handling

    '======================== Ensure that the desktop path ends with a backslash
    If Right(desktopPath, 1) <> "\" Then
        desktopPath = desktopPath & "\"
    End If

    '======================== Get the folder and file name from the specified cells
    folderName = shDatabase.Range("L1").value
    fileName = shDatabase.Range("L2").value

    '======================== Validate folder and file names
    If folderName = "" Or fileName = "" Then
        MsgBox "Folder name or file name is missing in the database sheet.", vbExclamation
        Exit Sub
    End If

    ' Construct the full file path
    filePath = desktopPath & folderName & "\" & fileName

    '======================== Check if the file exists
    If Dir(filePath) <> vbNullString Then
        ' File exists, open it
        Set seFileWorkbook = Workbooks.Open(filePath)
        Set seFileWorksheet = seFileWorkbook.Sheets(1) ' Assume we're working with the first sheet
    Else
        ' File does not exist, display a message and exit the sub
        MsgBox "The file '" & fileName & "' is not found in the folder '" & folderName & "' on your desktop.", vbExclamation
        Exit Sub
    End If

    '======================== Find the last row with data in column M of shSekFile (to start transferring data)
    lastRowSekFile = shSekFile.Cells(shSekFile.Rows.Count, "M").End(xlUp).row

    '======================== Define the starting row in the target file (assumed to start from row 9)
    targetRow = 9

    '======================== Loop through each row from row 2 to the last row with data in the SEK file
    For cellLoop = 2 To lastRowSekFile
        ' Error handling for each row processing
        On Error Resume Next
        With seFileWorksheet
            .Cells(targetRow, 1).value = "D"     ' Column A
            .Cells(targetRow, 2).value = shSekFile.Cells(cellLoop, "M").value ' Column B from SEK file column M
            .Cells(targetRow, 3).value = "D"     ' Column C
            .Cells(targetRow, 4).value = Right(shSekFile.Cells(cellLoop, "N").value, 8) ' Column D last 8 characters from SEK file column N
            .Cells(targetRow, 5).value = "SEK"   ' Column E
            .Cells(targetRow, 6).value = shSekFile.Cells(cellLoop, "R").value ' Column F from SEK file column R
            .Cells(targetRow, 7).value = Format(Date, "YYMMDD") ' Column G today's date
            .Cells(targetRow, 8).value = "MMM"   ' Column H
            .Cells(targetRow, 9).value = "777"   ' Column I
            .Cells(targetRow, 10).value = "EVENT " & shSekFile.Cells(cellLoop, "K").value & " " & shSekFile.Cells(cellLoop, "A").value ' Column J
        End With
        If Err.Number <> 0 Then
            MsgBox "Error processing row " & cellLoop & ".", vbCritical
            Err.Clear
        End If
        
        ' Move to the next row in the target file
        targetRow = targetRow + 1
    Next cellLoop

    '======================== Ask the user if they want to save the SEK file
    answer = MsgBox("Would you like to save the SEK file?", vbQuestion + vbYesNo)

    If answer = vbYes Then
        On Error Resume Next
        Call saveFile
        If Err.Number <> 0 Then
            MsgBox "Error saving the SEK file.", vbCritical
            Err.Clear
        End If
    Else
        Exit Sub
    End If

    ' Exit the subroutine successfully
    Exit Sub

ErrorHandler:
    ' Handle unexpected errors
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    Err.Clear
End Sub


'###########################################################################################################################
'###################################################### -- MESSAGES comparer to SEK file -- ###################################
'###########################################################################################################################

Sub sekMessagesChecker()
    '======================== Variable Declarations
    Dim lastRow As Long
    Dim messageQualifiers As Variant
    Dim NextRow As Long
    Dim Fndr As Range
    Dim qualifier As Variant
    Dim row As Long
    Dim Fndr2 As Range
    Dim semeRef As String
    Dim cellRow As Long
    Dim matching As Boolean

    ' Enable error handling
    On Error GoTo ErrorHandler

    '======================== Initialize Variables
    lastRow = shSekFile.Cells(shSekFile.Rows.Count, "I").End(xlUp).row
    NextRow = lastRow + 1

    shMain.Columns("I").Style = "Normal"
    matching = False

    '======================== Array of MESSAGE qualifiers to check
    messageQualifiers = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15")

    '======================== Find the MMO reference in shMain
    Set Fndr = shMain.Columns(9).Find(What:="3" & "*", LookAt:=xlWhole)
    
    If Not Fndr Is Nothing Then
        row = Fndr.row
        semeRef = extractValue(shMain.Cells(row, "I").value, "3")
    Else
        MsgBox "MMO reference not found!", vbExclamation
        Exit Sub                                 ' Exit the subroutine if MMO reference is not found
    End If
    
    '======================== Find MMO reference in shSekFile
    Set Fndr2 = shSekFile.Columns(9).Find(What:=semeRef & "*", LookAt:=xlWhole)
    
    If Not Fndr2 Is Nothing Then
        cellRow = Fndr2.row
    Else
        MsgBox "MMO reference not found in shSekFile!", vbExclamation
        Exit Sub                                 ' Exit the subroutine if MMO reference is not found in shSekFile
    End If
    
    '======================== Loop through each qualifier
    For Each qualifier In messageQualifiers
        ' Find the qualifier in column I
        Set Fndr = shMain.Columns(9).Find(What:=qualifier & "*", LookAt:=xlWhole)
        
        matching = False
        
        '======================== If the qualifier is found, get the row number
        If Not Fndr Is Nothing Then
            row = Fndr.row
            ' Process each value based on the found row
            Select Case qualifier
                Case "6"
                    If shSekFile.Cells(cellRow, "G").value = Mid(shMain.Cells(row, "I").value, InStr(1, shMain.Cells(row, "I").value, "777") + Len("777"), 12) Then
                        matching = True
                    End If
                ' Additional case checks...
            End Select
        Else
            ' If the qualifier is not found, highlight the cell in red
            shMain.Cells(row, "I").Interior.Color = RGB(255, 0, 0)
        End If
        
        '======================== Highlight the cell in green if matching, otherwise in red
        If matching Then
            shMain.Cells(row, "I").Interior.Color = RGB(0, 255, 0)
        Else
            shMain.Cells(row, "I").Interior.Color = RGB(255, 0, 0)
        End If
    Next qualifier

    ' Exit the subroutine normally
    Exit Sub

ErrorHandler:
    '======================== Error handling block
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    Err.Clear
End Sub


'###########################################################################################################################
'###################################################### -- Save File Macro -- ##############################################
'###########################################################################################################################

Sub saveFile()
    Dim wb As Workbook
    Dim savePath As String
    Dim folderDialog As FileDialog
    Dim saveFileName As String

    '======================== Error Handling
    On Error GoTo ErrorHandler

    '======================== Create a new FileDialog object to let the user select the folder
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.Title = "Select Folder to Save the File"
    folderDialog.AllowMultiSelect = False

    '======================== Show the folder picker dialog and get the selected folder
    If folderDialog.Show = -1 Then               ' If the user pressed "OK"
        savePath = folderDialog.SelectedItems(1)
    Else
        MsgBox "No folder selected. The file will not be saved.", vbInformation
        Exit Sub
    End If

    '======================== Ensure that the selected path ends with a backslash
    If Right(savePath, 1) <> "\" Then
        savePath = savePath & "\"
    End If

    '======================== Create a new workbook for saving the file
    Set wb = Workbooks.Add

    '======================== Copy all data from shSekFile to the new workbook
    shSekFile.Copy Before:=wb.Sheets(1)

    '======================== Construct the full file path with today's date
    saveFileName = "SEK_" & Format(Date, "YYYYMMDD") & ".xlsx"

    '======================== Save the file
    wb.SaveAs fileName:=savePath & saveFileName
    wb.Close SaveChanges:=False

    '======================== Notify the user that the file has been saved
    MsgBox "File saved successfully as: " & saveFileName & vbCrLf & "Location: " & savePath

    '======================== Normal exit
    Exit Sub

ErrorHandler:
    '======================== Handle errors
    MsgBox "An error occurred while saving the file: " & Err.Description, vbCritical
    Err.Clear
End Sub

'###################################################### -- Additional -- ############################################

Sub createGetSekTR()

    Call getValues
    Call CompareAndPrintValues

End Sub

Sub clearColumnA()
    shMain.Columns("A").Cells.ClearContents
    shMain.Columns("I").ClearContents
    shMain.Columns("I").Style = "Normal"
End Sub

