Attribute VB_Name = "Module2"
Function ConcatSecondDate(rng As Range) As String
    Dim cell As Range
    Dim firstCell As String, secondCell As String, thirdCell As String
    Dim monthNames As Variant
    Dim foundCount As Integer
    Dim secondDateCell As Range
    
    monthNames = Array("January", "February", "March", "April", "May", "June", "July", _
                      "August", "September", "October", "November", "December")
    
    foundCount = 0
    Set secondDateCell = Nothing

    ' Loop through range to find occurrences of a date or text with a month
    For Each cell In rng
        If IsDate(cell.Value) Or ContainsMonth(cell.Value, monthNames) Then
            foundCount = foundCount + 1
            ' Store the first occurrence but continue searching
            If foundCount = 2 Then
                Set secondDateCell = cell
                Exit For
            ElseIf foundCount = 1 Then
                Set secondDateCell = cell ' Temporary assignment, updated if 2nd is found
            End If
        End If
    Next cell
    
    ' Process the selected date
    If Not secondDateCell Is Nothing Then
        ' Format real dates
        If IsDate(secondDateCell.Value) Then
            firstCell = Format(secondDateCell.Value, "DD MMMM YYYY")
        Else
            firstCell = secondDateCell.Value
        End If
        
        ' Get next two adjacent values
        secondCell = secondDateCell.Offset(0, 1).Value
        thirdCell = secondDateCell.Offset(0, 2).Value
        
        ' Concatenate result
        ConcatSecondDate = Trim(firstCell & " " & secondCell & " " & thirdCell)
    Else
        ConcatSecondDate = "" ' No date found
    End If
End Function

' Function to check if a cell contains a month name
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



