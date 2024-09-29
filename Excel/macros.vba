Sub DeleteEmptyColumns()
    Dim rng As Range
    Dim delRange As Range
    Dim i As Long
    
    ' Set the range to the used range of the active sheet
    Set rng = ActiveSheet.UsedRange
    
    ' Disable screen updating and calculation to speed up the macro
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each column in the range from right to left
    For i = rng.Columns.Count To 1 Step -1
        ' Check if the entire column is empty
        If Application.WorksheetFunction.CountA(rng.Columns(i)) = 0 Then
            ' Add the column to the delete range
            If delRange Is Nothing Then
                Set delRange = rng.Columns(i)
            Else
                Set delRange = Union(delRange, rng.Columns(i))
            End If
        End If
    Next i
    
    ' Delete the empty columns
    If Not delRange Is Nothing Then
        delRange.Delete
    End If
    
    ' Re-enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub DeleteEmptyRows()
    Dim rng As Range
    Dim delRange As Range
    Dim i As Long
    
    ' Set the range to the used range of the active sheet
    Set rng = ActiveSheet.UsedRange
    
    ' Disable screen updating and calculation to speed up the macro
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each row in the range from bottom to top
    For i = rng.Rows.Count To 1 Step -1
        ' Check if the entire row is empty
        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then
            ' Add the row to the delete range
            If delRange Is Nothing Then
                Set delRange = rng.Rows(i)
            Else
                Set delRange = Union(delRange, rng.Rows(i))
            End If
        End If
    Next i
    
    ' Delete the empty rows
    If Not delRange Is Nothing Then
        delRange.Delete
    End If
    
    ' Re-enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
