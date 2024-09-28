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