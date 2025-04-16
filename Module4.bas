Attribute VB_Name = "Module4"
Sub ProcessDataOnce()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim rng As Range, cell As Range
    Dim firstCell As String, secondCell As String, thirdCell As String
    Dim foundCount As Integer
    Dim secondDateCell As Range
    Dim monthNames As Variant
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row ' Find last row
    monthNames = Array("January", "February", "March", "April", "May", "June", "July", _
                      "August", "September", "October", "November", "December")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Speed up macro
    
    For i = 2 To lastRow ' Loop through rows
        foundCount = 0
        Set secondDateCell = Nothing
        Set rng = ws.Range(ws.Cells(i, 4), ws.Cells(i, 16)) ' D:P range
        
        ' Find first or second date in the row
        For Each cell In rng
            If IsDate(cell.Value) Or ContainsMonth(cell.Value, monthNames) Then
                foundCount = foundCount + 1
                If foundCount = 2 Then
                    Set secondDateCell = cell
                    Exit For
                ElseIf foundCount = 1 Then
                    Set secondDateCell = cell ' Store first, but update if second is found
                End If
            End If
        Next cell
        
        ' Process and store result in Z column
        If Not secondDateCell Is Nothing Then
            ' Format date properly
            If IsDate(secondDateCell.Value) Then
                firstCell = Format(secondDateCell.Value, "DD MMMM YYYY")
            Else
                firstCell = secondDateCell.Value
            End If
            
            ' Get adjacent two values
            secondCell = secondDateCell.Offset(0, 1).Value
            thirdCell = secondDateCell.Offset(0, 2).Value
            
            ' Store result in column Z
            ws.Cells(i, 30).Value = Trim(firstCell & " " & secondCell & " " & thirdCell)
        Else
            ws.Cells(i, 30).Value = "" ' No date found
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic ' Restore calculation
    Application.ScreenUpdating = True
    
    MsgBox "Processing complete!", vbInformation
End Sub

' Function to check if a text contains a month name
Function ContainsMonth(ByVal txt As String, ByVal months As Variant) As Boolean
    Dim m As Variant
    For Each m In months
        If InStr(1, txt, m, vbTextCompare) > 0 Then
            ContainsMonth = True
            Exit Function
        End If
    Next m
    ContainsMonth = False
End Function


