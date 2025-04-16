Attribute VB_Name = "Module3"
Sub RemoveEmptyLinesInCells()
Attribute RemoveEmptyLinesInCells.VB_ProcData.VB_Invoke_Func = "m\n14"
    Dim ws As Worksheet
    Dim cell As Range
    Dim cleanedText As String
    Dim lines() As String
    Dim i As Integer
    Dim finalText As String

    ' Set the worksheet
    Set ws = ActiveSheet

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Temporarily disable auto-calculation

    ' Loop through all used cells in Column B
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            ' Remove extra spaces & normalize line breaks
            cleanedText = Trim(cell.Value)
            cleanedText = Replace(cleanedText, Chr(13) & Chr(10), Chr(10)) ' Fix CRLF line breaks
            cleanedText = Replace(cleanedText, Chr(13), Chr(10)) ' Fix single CR line breaks
            cleanedText = Replace(cleanedText, Chr(10) & Chr(10), Chr(10)) ' Remove double line breaks
            
            ' Split the cell content by line breaks
            lines = Split(cleanedText, Chr(10))
            finalText = ""

            ' Loop through each line and add it only if it's not empty
            For i = LBound(lines) To UBound(lines)
                If Len(Trim(lines(i))) > 0 Then
                    If finalText <> "" Then
                        finalText = finalText & Chr(10) & Trim(lines(i))
                    Else
                        finalText = Trim(lines(i))
                    End If
                End If
            Next i
            
            ' Update the cell with cleaned text
            cell.Value = finalText
            cell.WrapText = True ' Enable text wrapping
            
            ' Simulate manual re-entry of data (fixes formula update issue)
            cell.Value = cell.Value
        End If
    Next cell

    ' Re-enable calculations and force recalculation
    Application.Calculation = xlCalculationAutomatic
    ws.Calculate ' Force worksheet formulas to update

    ' Enable screen updating
    Application.ScreenUpdating = True

    MsgBox "Empty lines removed and formulas updated!", vbInformation
End Sub

Sub CopyAndTextToColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    ' Set the worksheet (change "Sheet1" to your actual sheet name)
    Set ws = ThisWorkbook.Sheets("summry")
    ws.Activate

    ' Find the last used row in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Check if there is data in column C before proceeding
    If lastRow < 1 Then
        MsgBox "No data found in Column C!", vbExclamation
        Exit Sub
    End If

    ' Copy column C and paste as values in column D
    ws.Range("C1:C" & lastRow).Copy
    ws.Range("D1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Clear clipboard

    ' Define the range for column D explicitly
    Set rng = ws.Range("D1:D" & lastRow)

    ' Apply Text to Columns to the data in column D
    rng.TextToColumns DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|"

    ' Inform the user
    MsgBox "Column C copied as values to Column D and Text to Columns applied!", vbInformation
Call ProcessDataOnce
End Sub




