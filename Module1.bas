Attribute VB_Name = "Module1"
Sub ExtractReservedLinesWithAutoAdjustmentOptimized()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdRng As Object
    Dim para As Object
    Dim textData As String
    Dim pageNum As Integer
    Dim lineText As String
    Dim excelRow As Long
    Dim resultData() As Variant
    Dim resultIndex As Integer

    ' Initialize Word application
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject(Class:="Word.Application")
    On Error GoTo 0
    wdApp.Visible = False

    ' Ask user to select a Word document
    Dim filePath As String
    filePath = Application.GetOpenFilename("Word Files (*.doc; *.docx), *.doc; *.docx")
    If filePath = "False" Then Exit Sub

    Set wdDoc = wdApp.Documents.Open(filePath)

    ' Disable Excel screen updating for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Preallocate results array
    ReDim resultData(1 To wdDoc.ComputeStatistics(Statistic:=2), 1 To 2) ' 2 = wdStatisticPages
    resultIndex = 1

    ' Loop through the document by pages
    For pageNum = 1 To wdDoc.ComputeStatistics(Statistic:=2)
        Set wdRng = wdDoc.GoTo(What:=1, Which:=1, Count:=pageNum) ' 1 = wdGoToPage

        ' Capture the page range
        wdRng.End = wdDoc.GoTo(What:=1, Which:=1, Count:=pageNum + 1).Start - 1
        textData = wdRng.Text

        If InStr(1, textData, "Reserved", vbTextCompare) > 0 Then
            ' If "Reserved" is found, extract text up to that line
            lineText = ""
            For Each para In wdRng.Paragraphs
                lineText = lineText & para.Range.Text
                If InStr(1, para.Range.Text, "Reserved", vbTextCompare) > 0 Then Exit For
            Next para

            ' Store the result in the array
            resultData(resultIndex, 1) = pageNum
            resultData(resultIndex, 2) = lineText
            resultIndex = resultIndex + 1
        End If
    Next pageNum

    ' Write results to Excel
    With Sheet1
        .Cells(1, 1).Value = "Page Number"
        .Cells(1, 2).Value = "Extracted Text"
        .Range(.Cells(2, 1), .Cells(UBound(resultData, 1) + 1, 2)).Value = resultData
        .Columns(2).WrapText = True
        .Rows.AutoFit
    End With

    ' Clean up
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    ' Re-enable Excel screen updating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Call RemoveEmptyLinesInCells
    MsgBox "Data extraction and formatting completed!", vbInformation
End Sub
