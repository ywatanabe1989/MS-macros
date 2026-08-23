Sub InsertImagesOnePerSlide()
    Dim fd As FileDialog
    Dim vrtSelectedItem As Variant
    Dim pptSlide As Slide
    Dim pptLayout As CustomLayout
    Dim imgPath As String
    Dim slideIndex As Integer
    
    ' Create a file picker dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select Images to Insert"
        .Filters.Clear
        .Filters.Add "Images", "*.jpg; *.jpeg; *.png; *.gif; *.bmp; *.tif; *.tiff"
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            slideIndex = ActivePresentation.Slides.Count
            
            ' Set layout to blank slide
            Set pptLayout = ActivePresentation.Slides(1).CustomLayout
            
            ' Loop through selected files
            For Each vrtSelectedItem In .SelectedItems
                slideIndex = slideIndex + 1
                
                ' Add new slide
                Set pptSlide = ActivePresentation.Slides.AddSlide(slideIndex, pptLayout)
                
                ' Insert image
                With pptSlide.Shapes.AddPicture(fileName:=vrtSelectedItem, _
                    LinkToFile:=msoFalse, _
                    SaveWithDocument:=msoTrue, _
                    Left:=0, Top:=0)
                    
                    ' Fit image to slide while maintaining aspect ratio
                    Dim slideWidth As Single, slideHeight As Single
                    slideWidth = ActivePresentation.PageSetup.slideWidth
                    slideHeight = ActivePresentation.PageSetup.slideHeight
                    
                    Dim imgRatio As Single, slideRatio As Single
                    imgRatio = .Width / .Height
                    slideRatio = slideWidth / slideHeight
                    
                    If imgRatio > slideRatio Then
                        ' Image is wider - fit to width
                        .Width = slideWidth * 0.9 ' 90% of slide width
                        .Height = .Width / imgRatio
                    Else
                        ' Image is taller - fit to height
                        .Height = slideHeight * 0.9 ' 90% of slide height
                        .Width = .Height * imgRatio
                    End If
                    
                    ' Center the image
                    .Left = (slideWidth - .Width) / 2
                    .Top = (slideHeight - .Height) / 2
                    
                    ' Add filename as caption
                    Dim fileName As String
                    fileName = Mid(vrtSelectedItem, InStrRev(vrtSelectedItem, "\") + 1)
                    
                    With pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        10, slideHeight - 40, slideWidth - 20, 30)
                        .TextFrame.textRange.text = fileName
                        .TextFrame.textRange.Font.Size = 10
                        .TextFrame.textRange.ParagraphFormat.Alignment = ppAlignCenter
                    End With
                End With
            Next vrtSelectedItem
            
            MsgBox "Successfully inserted " & .SelectedItems.Count & " images!", vbInformation
        End If
    End With
End Sub
