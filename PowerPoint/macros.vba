' ./powerpoint/macros.vba

Function GetRGBColor(colorName As String) As Long
    ' Function to get RGB color values by name
    ' Usage: color = GetRGBColor("blue")
    
    Dim colors As Object
    Set colors = CreateObject("Scripting.Dictionary")
    
    colors.Add "white", RGB(255, 255, 255)
    colors.Add "black", RGB(0, 0, 0)
    colors.Add "blue", RGB(0, 128, 192)
    colors.Add "red", RGB(255, 70, 50)
    colors.Add "pink", RGB(255, 150, 200)
    colors.Add "green", RGB(20, 180, 20)
    colors.Add "yellow", RGB(230, 160, 20)
    colors.Add "gray", RGB(128, 128, 128)
    colors.Add "grey", RGB(128, 128, 128)
    colors.Add "purple", RGB(200, 50, 255)
    colors.Add "light_blue", RGB(20, 200, 200)
    colors.Add "brown", RGB(128, 0, 0)
    colors.Add "navy", RGB(0, 0, 100)
    colors.Add "orange", RGB(228, 94, 50)
    
    If colors.Exists(colorName) Then
        GetRGBColor = colors(colorName)
    Else
        GetRGBColor = RGB(0, 0, 0) ' Default to black if color not found
    End If
End Function

Sub SetDefaultColors()
    ' Set default colors for the active presentation
    ' Usage: Run this macro to apply predefined color scheme
    
    With ActivePresentation.ColorScheme
        .Colors(ppBackground).RGB = GetRGBColor("white")
        .Colors(ppForeground).RGB = GetRGBColor("black")
        .Colors(ppShadow).RGB = GetRGBColor("gray")
        .Colors(ppTitle).RGB = GetRGBColor("blue")
        .Colors(ppFill).RGB = GetRGBColor("light_blue")
        .Colors(ppAccent1).RGB = GetRGBColor("red")
        .Colors(ppAccent2).RGB = GetRGBColor("green")
        .Colors(ppAccent3).RGB = GetRGBColor("yellow")
    End With
End Sub

Sub MultipleCropping()
    ' Apply cropping to multiple selected shapes
    ' Usage: Select multiple shapes, then run this macro
    ' The first selected shape's crop values will be applied to all
    
    Dim shp As Shape
    Dim firstShape As Shape
    Dim cropLeft As Single, cropTop As Single, cropRight As Single, cropBottom As Single
    
    ' Check if any shapes are selected
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select multiple shapes before running this macro.", vbExclamation
        Exit Sub
    End If
    
    ' Get the crop values from the first selected shape
    Set firstShape = ActiveWindow.Selection.ShapeRange(1)
    With firstShape.PictureFormat
        cropLeft = .cropLeft
        cropTop = .cropTop
        cropRight = .cropRight
        cropBottom = .cropBottom
    End With
    
    ' Apply the crop to all selected shapes
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoPicture Then
            With shp.PictureFormat
                .cropLeft = cropLeft
                .cropTop = cropTop
                .cropRight = cropRight
                .cropBottom = cropBottom
            End With
        End If
    Next shp
    
    MsgBox "Cropping applied to all selected shapes.", vbInformation
End Sub

Sub CropWhiteSpace()
    On Error GoTo ErrorHandler
    
    Dim shp As Shape
    Dim tempImagePath As String
    Dim pythonScriptPath As String
    Dim wslCommand As String
    Dim debugInfo As String
    Dim wshShell As Object
    Dim wshExec As Object
    Dim tempImagePathWSL As String
    Dim shpLeft As Single, shpTop As Single
    Dim wslCheckCommand As String
    
    debugInfo = "Starting CropWhiteSpace" & vbNewLine
    
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Please select exactly one picture object.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    shpLeft = shp.Left
    shpTop = shp.Top
    
    ' Paths
    tempImagePath = Environ$("TEMP") & "\temp.tif"
    pythonScriptPath = "/home/ywatanabe/.dotfiles/.bin/crop_whitespace.py"
    
    debugInfo = debugInfo & "Exporting shape to: " & tempImagePath & vbNewLine
    
    ' Export shape as image
    On Error Resume Next
    shp.Export tempImagePath, ppShapeFormatTIF
    If Err.Number <> 0 Then
        debugInfo = debugInfo & "Error exporting shape: " & Err.Description & vbNewLine
        GoTo ErrorHandler
    End If
    On Error GoTo ErrorHandler
    
    ' Check if file was created after export
    If Not FileExists(tempImagePath) Then
        debugInfo = debugInfo & "Error: Temp file not created after export" & vbNewLine
        GoTo ErrorHandler
    End If
    
    ' Convert Windows path to WSL path
    tempImagePathWSL = Replace(tempImagePath, "C:", "/mnt/c")
    tempImagePathWSL = Replace(tempImagePathWSL, "\", "/")
    tempImagePathWSL = Replace(tempImagePathWSL, " ", "\ ")
    
    debugInfo = debugInfo & "WSL path: " & tempImagePathWSL & vbNewLine
    debugInfo = debugInfo & "WSL image path: " & tempImagePathWSL & vbNewLine
    
    ' Check if file exists in WSL
    Set wshShell = CreateObject("WScript.Shell")
    wslCheckCommand = "wsl test -f '" & tempImagePathWSL & "' && echo 'File exists' || echo 'File not found'"
    Set wshExec = wshShell.Exec(wslCheckCommand)
    Do While wshExec.Status = 0
        DoEvents
    Loop
    debugInfo = debugInfo & "WSL File Check: " & wshExec.StdOut.ReadAll & vbNewLine

    If InStr(wshExec.StdOut.ReadAll, "File not found") > 0 Then
        debugInfo = debugInfo & "Error: File not found in WSL before Python script" & vbNewLine
        GoTo ErrorHandler
    End If
    
    ' Check file permissions
    wslCommand = "wsl ls -l '" & tempImagePathWSL & "' && " & wslCommand
       
    ' Print the exact file path in WSL:
    wslCommand = "wsl echo 'File path:' && wsl realpath '" & tempImagePathWSL & "' && " & wslCommand
    
    ' Ensure the Python script can access required libraries:
    wslCommand = "wsl /home/ywatanabe/.env/bin/python3 -c 'import sys; print(sys.path)' && " & wslCommand
    
   
    ' Run Python script in WSL
    wslCommand = "wsl chmod 644 '" & tempImagePathWSL & "' && /home/ywatanabe/.env/bin/python3 '" & pythonScriptPath & "' -l '" & tempImagePathWSL & "' --margin 5"
    debugInfo = debugInfo & "WSL Command: " & wslCommand & vbNewLine
    
    Set wshExec = wshShell.Exec(wslCommand)
    
    Do While wshExec.Status = 0
        DoEvents
    Loop
    
    debugInfo = debugInfo & "WSL Command executed" & vbNewLine
    debugInfo = debugInfo & "Output: " & wshExec.StdOut.ReadAll & vbNewLine
    debugInfo = debugInfo & "Error: " & wshExec.StdErr.ReadAll & vbNewLine
    

    ' Add delay
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + 1 ' 1-second delay
        DoEvents
    Loop
    debugInfo = debugInfo & "Waited for 1 second after WSL command" & vbNewLine

    ' Check if file exists in Windows after Python script
    If Not FileExists(tempImagePath) Then
        debugInfo = debugInfo & "Error: Temp file not found in Windows after Python script" & vbNewLine
        GoTo ErrorHandler
    End If
    
    ' Check if file exists in WSL after Python script
    wslCheckCommand = "wsl test -f '" & tempImagePathWSL & "' && echo 'File exists' || echo 'File not found'"
    Set wshExec = wshShell.Exec(wslCheckCommand)
    Do While wshExec.Status = 0
        DoEvents
    Loop
    If InStr(wshExec.StdOut.ReadAll, "File not found") > 0 Then
        debugInfo = debugInfo & "Error: File not found in WSL after Python script" & vbNewLine
        GoTo ErrorHandler
    End If
    
    debugInfo = debugInfo & "Temp file exists (Windows): " & FileExists(tempImagePath) & vbNewLine
    
    ' Import cropped image
    debugInfo = debugInfo & "Importing cropped image" & vbNewLine
    On Error Resume Next
    shp.Delete
    If Err.Number <> 0 Then
        debugInfo = debugInfo & "Error deleting original shape: " & Err.Description & vbNewLine
        Err.Clear
    End If
    Set shp = ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex).Shapes.AddPicture(tempImagePath, msoFalse, msoTrue, shpLeft, shpTop)
    If Err.Number <> 0 Then
        debugInfo = debugInfo & "Error adding new picture: " & Err.Description & vbNewLine
        GoTo ErrorHandler
    End If
    On Error GoTo ErrorHandler
    
    ' Delete the temporary file
    ' Kill tempImagePath
    
    debugInfo = debugInfo & "CropWhiteSpace completed successfully" & vbNewLine
    MsgBox "Operation completed. Debug info:" & vbNewLine & debugInfo, vbInformation
    Exit Sub

ErrorHandler:
    debugInfo = debugInfo & "Error: " & Err.Description & vbNewLine
    MsgBox "An error occurred. Debug info:" & vbNewLine & debugInfo, vbCritical
End Sub

Function FileExists(ByVal filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

