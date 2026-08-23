Sub InsertImagesTiled()
    Dim fd As FileDialog
    Dim vrtSelectedItem As Variant
    Dim pptSlide As Slide
    Dim imgPaths() As String
    Dim i As Integer, j As Integer, k As Integer
    Dim cols As Integer, rows As Integer
    Dim imgWidth As Single, imgHeight As Single
    Dim xPos As Single, yPos As Single
    Dim margin As Single
    Dim imagesPerSlide As Integer
    
    ' Configuration
    cols = InputBox("Enter number of columns:", "Tile Configuration", "2")
    rows = InputBox("Enter number of rows:", "Tile Configuration", "2")
    
    If cols = "" Or rows = "" Then Exit Sub
    
    cols = CInt(cols)
    rows = CInt(rows)
    imagesPerSlide = cols * rows
    
    ' Create file picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select Images to Insert"
        .Filters.Clear
        .Filters.Add "Images", "*.jpg; *.jpeg; *.png; *.gif; *.bmp; *.tif; *.tiff"
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            ' Store selected file paths
            ReDim imgPaths(1 To .SelectedItems.Count)
            For i = 1 To .SelectedItems.Count
                imgPaths(i) = .SelectedItems(i)
            Next i
            
            ' Calculate dimensions
            Dim slideWidth As Single, slideHeight As Single
            slideWidth = ActivePresentation.PageSetup.slideWidth
            slideHeight = ActivePresentation.PageSetup.slideHeight
            
            margin = 20 ' Margin between images
            imgWidth = (slideWidth - margin * (cols + 1)) / cols
            imgHeight = (slideHeight - margin * (rows + 1) - 50) / rows ' Leave space for title
            
            ' Process images
            k = 0
            Do While k < UBound(imgPaths)
                ' Add new slide
                Set pptSlide = ActivePresentation.Slides.Add( _
                    ActivePresentation.Slides.Count + 1, ppLayoutBlank)
                
                ' Add title
                With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    margin, 10, slideWidth - 2 * margin, 40)
                    .TextFrame.textRange.text = "Image Grid - Slide " & _
                        ActivePresentation.Slides.Count
                    .TextFrame.textRange.Font.Size = 18
                    .TextFrame.textRange.Font.Bold = msoTrue
                    .TextFrame.textRange.ParagraphFormat.Alignment = ppAlignCenter
                End With
                
                ' Insert images in grid
                For i = 0 To rows - 1
                    For j = 0 To cols - 1
                        k = k + 1
                        If k <= UBound(imgPaths) Then
                            xPos = margin + j * (imgWidth + margin)
                            yPos = 60 + margin + i * (imgHeight + margin)
                            
                            With pptSlide.Shapes.AddPicture( _
                                fileName:=imgPaths(k), _
                                LinkToFile:=msoFalse, _
                                SaveWithDocument:=msoTrue, _
                                Left:=xPos, Top:=yPos)
                                
                                ' Fit image to allocated space
                                Dim tempRatio As Single
                                tempRatio = .Width / .Height
                                
                                If tempRatio > (imgWidth / imgHeight) Then
                                    .Width = imgWidth
                                    .Height = imgWidth / tempRatio
                                    .Top = yPos + (imgHeight - .Height) / 2
                                Else
                                    .Height = imgHeight
                                    .Width = imgHeight * tempRatio
                                    .Left = xPos + (imgWidth - .Width) / 2
                                End If
                            End With
                        End If
                    Next j
                Next i
            Loop
            
            MsgBox "Successfully inserted " & UBound(imgPaths) & " images!", vbInformation
        End If
    End With
End Sub
