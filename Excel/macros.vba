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

Sub SaveAsXLSX()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ".xlsx"
    ThisWorkbook.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
End Sub


Sub ZebraRows()
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    
    ' Prompt user to select the range
    On Error Resume Next
    Set rng = Application.InputBox("Select the range for zebra striping", Type:=8)
    On Error GoTo 0
    
    If rng Is Nothing Then Exit Sub
    
    ' Determine last row and column of selected range
    lastRow = rng.Rows.Count
    lastCol = rng.Columns.Count
    
    ' Apply zebra striping
    For i = 1 To lastRow Step 2
        rng.Rows(i).Interior.Color = RGB(240, 240, 240) ' Light gray
    Next i
    
    ' Clear formatting for even rows
    For i = 2 To lastRow Step 2
        rng.Rows(i).Interior.Color = xlNone
    Next i
End Sub
