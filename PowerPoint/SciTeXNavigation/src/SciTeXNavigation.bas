Attribute VB_Name = "SciTeXNavigation"
Option Explicit

Private Const NAVIGATION_VERSION As String = "0.1.1"
Private Const TAG_CONFIG As String = "SCITEX_CONFIG"
Private Const TAG_COVER As String = "SCITEX_COVER"
Private Const TAG_TOC As String = "SCITEX_TOC"
Private Const TAG_SECTION_TITLE As String = "SCITEX_SECTION_TITLE"
Private Const TAG_NAV_CODE As String = "SCITEX_NAV_CODE"
Private Const TAG_CURRENT_SECTION As String = "SCITEX_CURRENT_SECTION"
Private Const TAG_TOC_SPLIT_AFTER As String = "SCITEX_TOC_SPLIT_AFTER"
Private Const TITLE_SHAPE As String = "SCITEX_TITLE"
Private Const TOC_BODY_SHAPE As String = "SCITEX_TOC_BODY"
Private Const TOC_BODY_LEFT As String = "SCITEX_TOC_BODY_LEFT"
Private Const TOC_BODY_RIGHT As String = "SCITEX_TOC_BODY_RIGHT"
Private Const STATUS_SHAPE As String = "SCITEX_STATUS"
Private Const CFG_FONT_LATIN As String = "SCITEX_CFG_FONT_LATIN"
Private Const CFG_FONT_CJK As String = "SCITEX_CFG_FONT_CJK"
Private Const CFG_FONT_MIN As String = "SCITEX_CFG_FONT_MIN"
Private Const CFG_FONT_MAX As String = "SCITEX_CFG_FONT_MAX"
Private Const CFG_HIDE_HIDDEN As String = "SCITEX_CFG_HIDE_HIDDEN"
Private Const CFG_VERSION As String = "SCITEX_CFG_VERSION"
Private Const TAG_ORIGINAL_FONT_SIZE As String = "SCITEX_ORIGINAL_FONT_SIZE"

Private mFontLatin As String
Private mFontCjk As String
Private mFontMin As Single
Private mFontMax As Single
Private mHideHiddenFromToc As Boolean
Private mExplicitNavigation As Boolean

' This is the only public macro. All implementation details stay out of Alt+F8.
Public Sub RunSciTeXNavigation()
    On Error GoTo Failed

    Dim pres As Presentation
    Dim sld As Slide
    Dim sectionNumber As Long
    Dim childNumber As Long
    Dim baseTitle As String

    Set pres = ActivePresentation
    LoadConfiguration pres
    SetDisplayedVersion pres
    mExplicitNavigation = HasExplicitNavigation(pres)
    sectionNumber = 0
    childNumber = 0

    If mExplicitNavigation Then
        sectionNumber = RenumberExplicitSlides(pres)
    Else
        For Each sld In pres.Slides
            If Not IsConfigSlide(sld) And Not IsCoverSlide(sld) Then
                If IsTocSlide(sld) Then
                    sectionNumber = sectionNumber + 1
                    childNumber = 0
                    baseTitle = SlideTag(sld, TAG_SECTION_TITLE)
                    If Len(baseTitle) = 0 Then baseTitle = StripNavigationPrefix(GetSlideTitle(sld))
                    SetSlideTitle sld, CStr(sectionNumber) & ". " & baseTitle
                ElseIf sectionNumber > 0 Then
                    childNumber = childNumber + 1
                    baseTitle = StripNavigationPrefix(GetSlideTitle(sld))
                    SetSlideTitle sld, CStr(sectionNumber) & LetterCode(childNumber) & ". " & baseTitle
                End If
            End If
        Next sld
    End If

    ' Apply font configuration before rebuilding links so explicit TOC colors win.
    ApplyTypography pres

    ' Every TOC slide shows the complete presentation outline.
    For Each sld In pres.Slides
        If IsTocSlide(sld) Then RebuildFullToc pres, sld
    Next sld

    SetStatus pres, "Navigation v" & NAVIGATION_VERSION & " updated - " & CStr(sectionNumber) & " sections"
    Exit Sub

Failed:
    SetStatus ActivePresentation, "Navigation error " & CStr(Err.Number) & ": " & Err.Description
    Err.Raise Err.Number, "RunSciTeXNavigation", Err.Description
End Sub

Private Sub RebuildFullToc(ByVal pres As Presentation, ByVal tocSlide As Slide)
    Dim body As Shape
    Dim target As Slide
    Dim targetIndex As Long
    Dim lineNumber As Long
    Dim tocText As String
    Dim titleText As String
    Dim paragraphRange As TextRange
    Dim linkRange As TextRange
    Dim linkLength As Long
    Dim currentSection As Long
    Dim targetSection As Long
    Dim leftBody As Shape
    Dim rightBody As Shape

    If mExplicitNavigation Then
        Set leftBody = FindNamedShape(tocSlide, TOC_BODY_LEFT)
        Set rightBody = FindNamedShape(tocSlide, TOC_BODY_RIGHT)
        If leftBody Is Nothing Or rightBody Is Nothing Then Err.Raise vbObjectError + 2100, , "Explicit TOC requires left and right body shapes."
        currentSection = ResolveCurrentSection(pres, tocSlide)
        If currentSection <= 0 Then Err.Raise vbObjectError + 2113, , "Cannot infer the current section for TOC slide " & CStr(tocSlide.SlideIndex) & "."
        tocSlide.Tags.Add TAG_CURRENT_SECTION, CStr(currentSection)
        RebuildExplicitTocColumn pres, tocSlide, leftBody, True, currentSection
        RebuildExplicitTocColumn pres, tocSlide, rightBody, False, currentSection
        Exit Sub
    End If

    Set body = FindNamedShape(tocSlide, TOC_BODY_SHAPE)
    If body Is Nothing Then Err.Raise vbObjectError + 2101, , "Missing " & TOC_BODY_SHAPE

    currentSection = 0
    For targetIndex = 1 To tocSlide.SlideIndex
        If IsTocSlide(pres.Slides(targetIndex)) Then currentSection = currentSection + 1
    Next targetIndex

    lineNumber = 0
    tocText = ""

    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If ShouldIncludeInToc(target) Then
            lineNumber = lineNumber + 1
            titleText = GetSlideTitle(target)
            If Len(tocText) > 0 Then tocText = tocText & vbCrLf
            tocText = tocText & titleText
        End If
    Next targetIndex

    body.TextFrame.TextRange.Text = tocText
    With body.TextFrame.Ruler.Levels(1)
        .FirstMargin = 0
        .LeftMargin = 0
    End With
    With body.TextFrame.Ruler.Levels(2)
        .FirstMargin = 24
        .LeftMargin = 24
    End With

    lineNumber = 0
    targetSection = 0
    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If IsTocSlide(target) Then targetSection = targetSection + 1
        If ShouldIncludeInToc(target) Then
            lineNumber = lineNumber + 1
            Set paragraphRange = body.TextFrame.TextRange.Paragraphs(lineNumber, 1)
            linkLength = Len(paragraphRange.Text)
            Do While linkLength > 0 And (Mid$(paragraphRange.Text, linkLength, 1) = vbCr Or Mid$(paragraphRange.Text, linkLength, 1) = vbLf)
                linkLength = linkLength - 1
            Loop
            Set linkRange = paragraphRange.Characters(1, linkLength)
            With linkRange.ActionSettings(ppMouseClick)
                .Action = ppActionHyperlink
                .Hyperlink.SubAddress = CStr(target.SlideID) & "," & CStr(target.SlideIndex) & "," & GetSlideTitle(target)
            End With
            If targetSection = currentSection Then
                linkRange.Font.Color.RGB = RGB(27, 38, 53)
                paragraphRange.Font.Color.RGB = RGB(27, 38, 53)
            Else
                linkRange.Font.Color.RGB = RGB(170, 179, 188)
                paragraphRange.Font.Color.RGB = RGB(170, 179, 188)
            End If
            If IsTocSlide(target) Then
                paragraphRange.IndentLevel = 1
                linkRange.Font.Bold = msoTrue
            Else
                paragraphRange.IndentLevel = 2
            End If
        End If
    Next targetIndex
End Sub

Private Function HasExplicitNavigation(ByVal pres As Presentation) As Boolean
    Dim sld As Slide
    For Each sld In pres.Slides
        If Len(SlideTag(sld, TAG_NAV_CODE)) > 0 Then
            HasExplicitNavigation = True
            Exit Function
        End If
    Next sld
    HasExplicitNavigation = False
End Function

Private Function RenumberExplicitSlides(ByVal pres As Presentation) As Long
    Dim sld As Slide
    Dim navigationCode As String
    Dim baseTitle As String
    Dim sectionNumber As Long
    Dim maximumSection As Long

    maximumSection = 0
    For Each sld In pres.Slides
        navigationCode = SlideTag(sld, TAG_NAV_CODE)
        If Len(navigationCode) > 0 Then
            baseTitle = StripNavigationPrefix(GetSlideTitle(sld))
            SetSlideTitle sld, navigationCode & ". " & baseTitle
            sectionNumber = CLng(Val(navigationCode))
            If sectionNumber > maximumSection Then maximumSection = sectionNumber
        End If
    Next sld
    RenumberExplicitSlides = maximumSection
End Function

Private Sub RebuildExplicitTocColumn(ByVal pres As Presentation, ByVal tocSlide As Slide, ByVal body As Shape, ByVal isLeftColumn As Boolean, ByVal currentSection As Long)
    Dim target As Slide
    Dim targetIndex As Long
    Dim lineNumber As Long
    Dim tocText As String
    Dim navigationCode As String
    Dim targetSection As Long
    Dim splitAfter As Long
    Dim paragraphRange As TextRange
    Dim linkRange As TextRange
    Dim linkLength As Long

    splitAfter = CLng(Val(SlideTag(tocSlide, TAG_TOC_SPLIT_AFTER)))
    If splitAfter <= 0 Then splitAfter = 3
    tocText = ""

    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        navigationCode = SlideTag(target, TAG_NAV_CODE)
        If ShouldIncludeInToc(target) Then
            targetSection = CLng(Val(navigationCode))
            If (isLeftColumn And targetSection <= splitAfter) Or (Not isLeftColumn And targetSection > splitAfter) Then
                If Len(tocText) > 0 Then tocText = tocText & vbCrLf
                tocText = tocText & GetSlideTitle(target)
            End If
        End If
    Next targetIndex

    body.TextFrame.TextRange.Text = tocText
    With body.TextFrame.Ruler.Levels(1)
        .FirstMargin = 0
        .LeftMargin = 0
    End With
    With body.TextFrame.Ruler.Levels(2)
        .FirstMargin = 18
        .LeftMargin = 18
    End With

    lineNumber = 0
    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        navigationCode = SlideTag(target, TAG_NAV_CODE)
        If ShouldIncludeInToc(target) Then
            targetSection = CLng(Val(navigationCode))
            If (isLeftColumn And targetSection <= splitAfter) Or (Not isLeftColumn And targetSection > splitAfter) Then
                lineNumber = lineNumber + 1
                Set paragraphRange = body.TextFrame.TextRange.Paragraphs(lineNumber, 1)
                linkLength = Len(paragraphRange.Text)
                Do While linkLength > 0 And (Mid$(paragraphRange.Text, linkLength, 1) = vbCr Or Mid$(paragraphRange.Text, linkLength, 1) = vbLf)
                    linkLength = linkLength - 1
                Loop
                Set linkRange = paragraphRange.Characters(1, linkLength)
                With linkRange.ActionSettings(ppMouseClick)
                    .Action = ppActionHyperlink
                    .Hyperlink.SubAddress = CStr(target.SlideID) & "," & CStr(target.SlideIndex) & "," & GetSlideTitle(target)
                End With
                If targetSection = currentSection Then
                    linkRange.Font.Color.RGB = RGB(27, 38, 53)
                    paragraphRange.Font.Color.RGB = RGB(27, 38, 53)
                Else
                    linkRange.Font.Color.RGB = RGB(170, 179, 188)
                    paragraphRange.Font.Color.RGB = RGB(170, 179, 188)
                End If
                If IsMajorNavigationCode(navigationCode) Then
                    paragraphRange.IndentLevel = 1
                    linkRange.Font.Bold = msoTrue
                Else
                    paragraphRange.IndentLevel = 2
                End If
            End If
        End If
    Next targetIndex
End Sub

Private Function ResolveCurrentSection(ByVal pres As Presentation, ByVal tocSlide As Slide) As Long
    Dim targetIndex As Long
    Dim navigationCode As String

    ' A copied or moved TOC keeps its old tags. Prefer the first navigated slide
    ' after its new position so the TOC automatically follows the deck order.
    For targetIndex = tocSlide.SlideIndex + 1 To pres.Slides.Count
        navigationCode = SlideTag(pres.Slides(targetIndex), TAG_NAV_CODE)
        If Len(navigationCode) > 0 Then
            ResolveCurrentSection = CLng(Val(navigationCode))
            Exit Function
        End If
    Next targetIndex

    ' A TOC placed at the end belongs to the most recent preceding section.
    For targetIndex = tocSlide.SlideIndex - 1 To 1 Step -1
        navigationCode = SlideTag(pres.Slides(targetIndex), TAG_NAV_CODE)
        If Len(navigationCode) > 0 Then
            ResolveCurrentSection = CLng(Val(navigationCode))
            Exit Function
        End If
    Next targetIndex

    ResolveCurrentSection = CLng(Val(SlideTag(tocSlide, TAG_CURRENT_SECTION)))
End Function

Private Function IsMajorNavigationCode(ByVal navigationCode As String) As Boolean
    IsMajorNavigationCode = (navigationCode = CStr(CLng(Val(navigationCode))))
End Function

Private Sub LoadConfiguration(ByVal pres As Presentation)
    Dim configSlide As Slide
    Dim value As String

    mFontLatin = "Aptos"
    mFontCjk = "Yu Gothic"
    mFontMin = 18
    mFontMax = 32
    mHideHiddenFromToc = True

    Set configSlide = FindConfigSlide(pres)
    If configSlide Is Nothing Then Exit Sub

    value = ConfigText(configSlide, CFG_FONT_LATIN, mFontLatin)
    If Len(value) > 0 Then mFontLatin = value
    value = ConfigText(configSlide, CFG_FONT_CJK, mFontCjk)
    If Len(value) > 0 Then mFontCjk = value
    mFontMin = CSng(Val(ConfigText(configSlide, CFG_FONT_MIN, CStr(mFontMin))))
    mFontMax = CSng(Val(ConfigText(configSlide, CFG_FONT_MAX, CStr(mFontMax))))
    mHideHiddenFromToc = ParseBoolean(ConfigText(configSlide, CFG_HIDE_HIDDEN, "Yes"))

    If mFontMin <= 0 Then Err.Raise vbObjectError + 2110, , "Minimum font size must be greater than zero."
    If mFontMax < mFontMin Then Err.Raise vbObjectError + 2111, , "Maximum font size must be greater than or equal to minimum font size."
End Sub

Private Function FindConfigSlide(ByVal pres As Presentation) As Slide
    Dim sld As Slide
    For Each sld In pres.Slides
        If IsConfigSlide(sld) Then
            Set FindConfigSlide = sld
            Exit Function
        End If
    Next sld
    Set FindConfigSlide = Nothing
End Function

Private Sub SetDisplayedVersion(ByVal pres As Presentation)
    Dim configSlide As Slide
    Dim versionShape As Shape

    Set configSlide = FindConfigSlide(pres)
    If configSlide Is Nothing Then Exit Sub
    Set versionShape = FindNamedShape(configSlide, CFG_VERSION)
    If Not versionShape Is Nothing Then versionShape.TextFrame.TextRange.Text = NAVIGATION_VERSION
End Sub

Private Function ConfigText(ByVal configSlide As Slide, ByVal shapeName As String, ByVal defaultValue As String) As String
    Dim valueShape As Shape
    Set valueShape = FindNamedShape(configSlide, shapeName)
    If valueShape Is Nothing Then
        ConfigText = defaultValue
    Else
        ConfigText = Trim$(valueShape.TextFrame.TextRange.Text)
        If Len(ConfigText) = 0 Then ConfigText = defaultValue
    End If
End Function

Private Function ParseBoolean(ByVal value As String) As Boolean
    Select Case LCase$(Trim$(value))
        Case "yes", "true", "1", "on"
            ParseBoolean = True
        Case "no", "false", "0", "off"
            ParseBoolean = False
        Case Else
            Err.Raise vbObjectError + 2112, , "Hide hidden slides must be Yes or No."
    End Select
End Function

Private Function ShouldIncludeInToc(ByVal sld As Slide) As Boolean
    If IsConfigSlide(sld) Or IsCoverSlide(sld) Then
        ShouldIncludeInToc = False
    ElseIf mExplicitNavigation And Len(SlideTag(sld, TAG_NAV_CODE)) = 0 Then
        ShouldIncludeInToc = False
    ElseIf mHideHiddenFromToc And sld.SlideShowTransition.Hidden <> msoFalse Then
        ShouldIncludeInToc = False
    Else
        ShouldIncludeInToc = True
    End If
End Function

Private Sub ApplyTypography(ByVal pres As Presentation)
    Dim sld As Slide
    Dim shp As Shape
    Dim textRange As TextRange
    Dim originalSize As Single
    Dim targetSize As Single

    For Each sld In pres.Slides
        If Not IsConfigSlide(sld) Then
            For Each shp In sld.Shapes
                If IsTypographyManagedShape(shp) And shp.HasTextFrame = msoTrue Then
                    If shp.TextFrame.HasText = msoTrue Then
                        Set textRange = shp.TextFrame.TextRange
                        textRange.Font.Name = mFontLatin
                        On Error Resume Next
                        textRange.Font.NameFarEast = mFontCjk
                        On Error GoTo 0
                        originalSize = CSng(Val(ShapeTag(shp, TAG_ORIGINAL_FONT_SIZE)))
                        If originalSize <= 0 Then
                            originalSize = textRange.Characters(1, 1).Font.Size
                            If originalSize <= 0 Then originalSize = mFontMin
                            shp.Tags.Add TAG_ORIGINAL_FONT_SIZE, CStr(originalSize)
                        End If
                        targetSize = originalSize
                        If targetSize < mFontMin Then targetSize = mFontMin
                        If targetSize > mFontMax Then targetSize = mFontMax
                        textRange.Font.Size = targetSize
                    End If
                End If
            Next shp
        End If
    Next sld
End Sub

Private Function ShapeTag(ByVal shp As Shape, ByVal tagName As String) As String
    On Error Resume Next
    ShapeTag = Trim$(shp.Tags(tagName))
    On Error GoTo 0
End Function

Private Function IsTypographyManagedShape(ByVal shp As Shape) As Boolean
    Select Case UCase$(shp.Name)
        Case UCase$(TITLE_SHAPE), UCase$(TOC_BODY_SHAPE), UCase$(TOC_BODY_LEFT), UCase$(TOC_BODY_RIGHT), UCase$(STATUS_SHAPE), "SCITEX_RUN_BUTTON"
            IsTypographyManagedShape = True
        Case Else
            IsTypographyManagedShape = False
    End Select
End Function

Private Function GetSlideTitle(ByVal sld As Slide) As String
    Dim titleShape As Shape
    Set titleShape = FindNamedShape(sld, TITLE_SHAPE)
    If titleShape Is Nothing Then Err.Raise vbObjectError + 2102, , "Slide " & CStr(sld.SlideIndex) & " is missing " & TITLE_SHAPE
    GetSlideTitle = Trim$(titleShape.TextFrame.TextRange.Text)
End Function

Private Sub SetSlideTitle(ByVal sld As Slide, ByVal value As String)
    Dim titleShape As Shape
    Set titleShape = FindNamedShape(sld, TITLE_SHAPE)
    If titleShape Is Nothing Then Err.Raise vbObjectError + 2103, , "Slide " & CStr(sld.SlideIndex) & " is missing " & TITLE_SHAPE
    titleShape.TextFrame.TextRange.Text = value
End Sub

Private Function FindNamedShape(ByVal sld As Slide, ByVal shapeName As String) As Shape
    Dim shp As Shape
    For Each shp In sld.Shapes
        If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
            Set FindNamedShape = shp
            Exit Function
        End If
    Next shp
    Set FindNamedShape = Nothing
End Function

Private Function SlideTag(ByVal sld As Slide, ByVal tagName As String) As String
    On Error Resume Next
    SlideTag = Trim$(sld.Tags(tagName))
    On Error GoTo 0
End Function

Private Function IsConfigSlide(ByVal sld As Slide) As Boolean
    IsConfigSlide = (StrComp(SlideTag(sld, TAG_CONFIG), "1", vbTextCompare) = 0)
End Function

Private Function IsCoverSlide(ByVal sld As Slide) As Boolean
    IsCoverSlide = (StrComp(SlideTag(sld, TAG_COVER), "1", vbTextCompare) = 0)
End Function

Private Function IsTocSlide(ByVal sld As Slide) As Boolean
    IsTocSlide = (StrComp(SlideTag(sld, TAG_TOC), "1", vbTextCompare) = 0)
End Function

Private Function StripNavigationPrefix(ByVal value As String) As String
    Dim text As String
    Dim dotPosition As Long
    Dim prefix As String

    text = Trim$(value)
    dotPosition = InStr(1, text, ".", vbBinaryCompare)
    If dotPosition > 1 Then
        prefix = Left$(text, dotPosition - 1)
        If IsNavigationPrefix(prefix) Then text = Trim$(Mid$(text, dotPosition + 1))
    End If
    StripNavigationPrefix = text
End Function

Private Function IsNavigationPrefix(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String
    Dim hasDigit As Boolean

    hasDigit = False
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If character >= "0" And character <= "9" Then
            hasDigit = True
        ElseIf (character >= "a" And character <= "z") Or (character >= "A" And character <= "Z") Then
            ' Letters are allowed after the section number.
        Else
            IsNavigationPrefix = False
            Exit Function
        End If
    Next index
    IsNavigationPrefix = hasDigit
End Function

Private Function LetterCode(ByVal number As Long) As String
    Dim value As Long
    Dim result As String

    value = number
    result = ""
    Do While value > 0
        value = value - 1
        result = Chr$(97 + (value Mod 26)) & result
        value = value \ 26
    Loop
    LetterCode = result
End Function

Private Sub SetStatus(ByVal pres As Presentation, ByVal value As String)
    On Error Resume Next
    Dim statusShape As Shape
    Set statusShape = FindNamedShape(pres.Slides(1), STATUS_SHAPE)
    If Not statusShape Is Nothing Then statusShape.TextFrame.TextRange.Text = value
    On Error GoTo 0
End Sub
