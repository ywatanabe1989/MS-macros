Sub ImportImagesFromFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim pptSlide As Slide
    Dim imageCount As Integer
    Dim layoutChoice As Integer
    
    ' Select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Images"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    ' Ask for layout preference
    layoutChoice = InputBox("Enter layout choice:" & vbCrLf & _
        "1 = One image per slide" & vbCrLf & _
        "2 = Tiled layout (4 per slide)" & vbCrLf & _
        "3 = Tiled layout (6 per slide)", "Layout Selection", "1")
    
    If layoutChoice = "" Then Exit Sub
    
    ' Process based on choice
    Select Case CInt(layoutChoice)
        Case 1
            ProcessFolderOnePerSlide folderPath
        Case 2
            ProcessFolderTiled folderPath, 2, 2
        Case 3
            ProcessFolderTiled folderPath, 3, 2
        Case Else
            MsgBox "Invalid choice!", vbExclamation
    End Select
End Sub

Private Sub ProcessFolderOnePerSlide(folderPath As String)
    Dim fileName As String
    Dim pptSlide As Slide
    Dim slideWidth As Single, slideHeight As Single
    
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    
    ' Get first image file
    fileName = Dir(folderPath & "*.jpg")
    
    ' Process all image files
    Do While fileName <> ""
        ' Add slide
        Set pptSlide = ActivePresentation.Slides.Add( _
            ActivePresentation.Slides.Count + 1, ppLayoutBlank)
        
        ' Insert image
        With pptSlide.Shapes.AddPicture( _
            fileName:=folderPath & fileName, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=0, Top:=0)
            
            ' Scale and center
            Dim scale As Single
            scale = Application.Min(slideWidth * 0.9 / .Width, _
                slideHeight * 0.9 / .Height)
            .Width = .Width * scale
            .Height = .Height * scale
            .Left = (slideWidth - .Width) / 2
            .Top = (slideHeight - .Height) / 2
        End With
        
        ' Get next file
        fileName = Dir()
    Loop
    
    ' Also process PNG files
    fileName = Dir(folderPath & "*.png")
    Do While fileName <> ""
        Set pptSlide = ActivePresentation.Slides.Add( _
            ActivePresentation.Slides.Count + 1, ppLayoutBlank)
        
        With pptSlide.Shapes.AddPicture( _
            fileName:=folderPath & fileName, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=0, Top:=0)
            
            Dim scale2 As Single
            scale2 = Application.Min(slideWidth * 0.9 / .Width, _
                slideHeight * 0.9 / .Height)
            .Width = .Width * scale2
            .Height = .Height * scale2
            .Left = (slideWidth - .Width) / 2
            .Top = (slideHeight - .Height) / 2
        End With
        
        fileName = Dir()
    Loop
End Sub

Private Sub ProcessFolderTiled(folderPath As String, cols As Integer, rows As Integer)
    ' Similar implementation to InsertImagesTiled but reading from folder
    ' Code omitted for brevity but follows same pattern
End Sub
