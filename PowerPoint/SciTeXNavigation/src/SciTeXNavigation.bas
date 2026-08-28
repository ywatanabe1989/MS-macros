Attribute VB_Name = "SciTeXNavigation"
Option Explicit

Private Const NAVIGATION_VERSION As String = "0.4.0"
Private Const TAG_CONFIG As String = "SCITEX_CONFIG"
Private Const TAG_COVER As String = "SCITEX_COVER"
Private Const TAG_TOC As String = "SCITEX_TOC"
Private Const TAG_SECTION_TITLE As String = "SCITEX_SECTION_TITLE"
Private Const TAG_NAV_CODE As String = "SCITEX_NAV_CODE"
Private Const TAG_CURRENT_SECTION As String = "SCITEX_CURRENT_SECTION"
Private Const TAG_TOC_SPLIT_AFTER As String = "SCITEX_TOC_SPLIT_AFTER"  ' read only to delete it
Private Const TITLE_SHAPE As String = "SCITEX_TITLE"
Private Const TOC_BODY_SHAPE As String = "SCITEX_TOC_BODY"
Private Const TOC_BODY_LEFT As String = "SCITEX_TOC_BODY_LEFT"
Private Const TOC_BODY_RIGHT As String = "SCITEX_TOC_BODY_RIGHT"
Private Const TOC_LINK_PREFIX As String = "SCITEX_TOC_LINK_"
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
Private mBackupPath As String
Private mSplitCached As Long
Private mSplitSlideIndex As Long
Private mPlannedSize As Single
Private mOverfullSlides As String
Private mHideHiddenFromToc As Boolean
Private mTwoColumnToc As Boolean

' This is the only public macro. All implementation details stay out of Alt+F8.
Public Sub RunSciTeXNavigation()
    On Error GoTo Failed

    Dim pres As Presentation
    Dim sld As Slide
    Dim sectionNumber As Long

    Set pres = ActivePresentation
    mOverfullSlides = ""
    mSplitCached = 0
    mSplitSlideIndex = 0
    BackupBeforeRun pres
    LoadConfiguration pres
    SetDisplayedVersion pres
    mTwoColumnToc = HasTwoColumnToc(pres)
    sectionNumber = RenumberTocDrivenSlides(pres)

    ' Apply font configuration before rebuilding links so TOC styling wins.
    ApplyTypography pres
    FitTocTitles pres

    ' Every TOC slide shows the complete presentation outline.
    For Each sld In pres.Slides
        If IsTocSlide(sld) Then RebuildFullToc pres, sld
    Next sld

    If Len(mOverfullSlides) > 0 Then
        ' Shrinking stopped at the configured minimum and the text still does
        ' not fit. Going smaller is not the answer -- at that point the slide
        ' is carrying more than a slide should. Say which ones, rather than
        ' leaving an overflow that looks like the bug this version fixed.
        SetStatus pres, "Navigation v" & NAVIGATION_VERSION & " updated - " & _
            CStr(sectionNumber) & " sections. Too much content at " & _
            CStr(mFontMin) & "pt on slide(s): " & mOverfullSlides & _
            ". Backup: " & BackupName()
    Else
        SetStatus pres, "Navigation v" & NAVIGATION_VERSION & " updated - " & _
            CStr(sectionNumber) & " sections. Backup: " & BackupName()
    End If
    Exit Sub

Failed:
    SetStatus ActivePresentation, "Navigation error " & CStr(Err.Number) & ": " & Err.Description
    Err.Raise Err.Number, "RunSciTeXNavigation", Err.Description
End Sub

' Take a copy before touching anything.
'
' PowerPoint's undo stack does not survive a macro: once this has renumbered
' slides and rewritten two index columns, Ctrl+Z will not put the deck back.
' The only real undo is a file that still has the old content, so make one
' first and tell the operator where it is. Requested 2026-08-27 after the
' operator asked "if it goes wrong can I get back".
'
' Refuses to run on a presentation that has never been saved, because there is
' nowhere to put the copy and nothing to go back to.
Private Sub BackupBeforeRun(ByVal pres As Presentation)
    Dim stamp As String
    Dim backupPath As String
    Dim baseName As String
    Dim dotPosition As Long

    If Len(pres.Path) = 0 Then
        Err.Raise vbObjectError + 2120, , _
            "Save the presentation first. This macro cannot be undone with Ctrl+Z, " & _
            "so it will not run on a file that has never been saved."
    End If

    baseName = pres.Name
    dotPosition = InStrRev(baseName, ".")
    If dotPosition > 1 Then baseName = Left$(baseName, dotPosition - 1)

    stamp = Format$(Now, "yyyymmdd-hhnnss")
    backupPath = pres.Path & Application.PathSeparator & _
        baseName & ".before-navigation-" & stamp & ".pptm"

    pres.SaveCopyAs backupPath, ppSaveAsOpenXMLPresentationMacroEnabled
    mBackupPath = backupPath
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

    DeleteTocLinkOverlays tocSlide

    If mTwoColumnToc Then
        Set leftBody = FindNamedShape(tocSlide, TOC_BODY_LEFT)
        Set rightBody = FindNamedShape(tocSlide, TOC_BODY_RIGHT)
        If leftBody Is Nothing Or rightBody Is Nothing Then Err.Raise vbObjectError + 2100, , "Two-column TOC requires left and right body shapes."
        currentSection = CLng(Val(SlideTag(tocSlide, TAG_NAV_CODE)))
        ' Same band for both columns before anything is measured, or the
        ' measurement answers a question about box sizes instead of text.
        NormaliseTocBoxes pres, leftBody, rightBody
        RebuildTocColumn pres, tocSlide, leftBody, True, currentSection
        RebuildTocColumn pres, tocSlide, rightBody, False, currentSection
        ' Both columns are laid out at the planned size, so this is a backstop
        ' rather than a correction: if FitTocBody had to shrink one of them,
        ' two sizes of the same list reads as a mistake even when both fit.
        MatchColumnSizes leftBody, rightBody
        Exit Sub
    End If

    Set body = FindNamedShape(tocSlide, TOC_BODY_SHAPE)
    If body Is Nothing Then Err.Raise vbObjectError + 2101, , "Missing " & TOC_BODY_SHAPE

    currentSection = CLng(Val(SlideTag(tocSlide, TAG_NAV_CODE)))

    lineNumber = 0
    tocText = ""

    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If ShouldIncludeInToc(target) Then
            lineNumber = lineNumber + 1
            titleText = NavigationEntryText(target)
            If Len(tocText) > 0 Then tocText = tocText & vbCrLf
            tocText = tocText & titleText
        End If
    Next targetIndex

    body.TextFrame.TextRange.Text = tocText
    body.TextFrame.TextRange.Font.Bold = msoFalse
    body.TextFrame.TextRange.Font.Underline = msoFalse
    With body.TextFrame.Ruler.Levels(1)
        .FirstMargin = 0
        .LeftMargin = 0
    End With
    With body.TextFrame.Ruler.Levels(2)
        .FirstMargin = 24
        .LeftMargin = 24
    End With

    lineNumber = 0
    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If ShouldIncludeInToc(target) Then
            targetSection = CLng(Val(SlideTag(target, TAG_NAV_CODE)))
            lineNumber = lineNumber + 1
            Set paragraphRange = body.TextFrame.TextRange.Paragraphs(lineNumber, 1)
            linkLength = Len(paragraphRange.Text)
            Do While linkLength > 0 And (Mid$(paragraphRange.Text, linkLength, 1) = vbCr Or Mid$(paragraphRange.Text, linkLength, 1) = vbLf)
                linkLength = linkLength - 1
            Loop
            Set linkRange = paragraphRange.Characters(1, linkLength)
            ClearTextHyperlink linkRange
            If targetSection = currentSection Then
                linkRange.Font.Color.RGB = RGB(27, 38, 53)
                paragraphRange.Font.Color.RGB = RGB(27, 38, 53)
            Else
                linkRange.Font.Color.RGB = RGB(170, 179, 188)
                paragraphRange.Font.Color.RGB = RGB(170, 179, 188)
            End If
            If IsTocSlide(target) Then
                paragraphRange.IndentLevel = 1
                paragraphRange.Font.Bold = msoFalse
                paragraphRange.Font.Underline = msoFalse
                linkRange.Font.Bold = msoTrue
                linkRange.Font.Underline = msoTrue
            Else
                paragraphRange.IndentLevel = 2
                linkRange.Font.Bold = msoFalse
                paragraphRange.Font.Bold = msoFalse
                linkRange.Font.Underline = msoFalse
                paragraphRange.Font.Underline = msoFalse
            End If
            AddTocLinkOverlay tocSlide, paragraphRange, target, "B", lineNumber
        End If
    Next targetIndex
End Sub

Private Function HasTwoColumnToc(ByVal pres As Presentation) As Boolean
    Dim sld As Slide
    Dim leftBody As Shape
    Dim rightBody As Shape

    For Each sld In pres.Slides
        If IsTocSlide(sld) Then
            Set leftBody = FindNamedShape(sld, TOC_BODY_LEFT)
            Set rightBody = FindNamedShape(sld, TOC_BODY_RIGHT)
            If Not leftBody Is Nothing And Not rightBody Is Nothing Then
                HasTwoColumnToc = True
                Exit Function
            End If
        End If
    Next sld
    HasTwoColumnToc = False
End Function

Private Function RenumberTocDrivenSlides(ByVal pres As Presentation) As Long
    Dim sld As Slide
    Dim sectionNumber As Long
    Dim childNumber As Long
    Dim sectionTitle As String
    Dim baseTitle As String
    Dim navigationCode As String
    Dim splitAfter As Long

    sectionNumber = 0
    childNumber = 0
    For Each sld In pres.Slides
        If Not IsConfigSlide(sld) And Not IsCoverSlide(sld) Then
            If IsTocSlide(sld) Then
                sectionNumber = sectionNumber + 1
                childNumber = 0
                sectionTitle = ResolveTocSectionTitle(pres, sld)
                If Len(sectionTitle) = 0 Then Err.Raise vbObjectError + 2113, , "Cannot determine the section title for TOC slide " & CStr(sld.SlideIndex) & "."
                sld.Tags.Add TAG_SECTION_TITLE, sectionTitle
                sld.Tags.Add TAG_NAV_CODE, CStr(sectionNumber)
                sld.Tags.Add TAG_CURRENT_SECTION, CStr(sectionNumber)
                SetTocVisibleTitle sld, sectionTitle
            ElseIf sectionNumber > 0 Then
                childNumber = childNumber + 1
                baseTitle = StripNavigationPrefix(GetSlideTitle(sld))
                navigationCode = CStr(sectionNumber) & LetterCode(childNumber)
                SetSlideTitle sld, navigationCode & ". " & baseTitle
                sld.Tags.Add TAG_NAV_CODE, navigationCode
            Else
                DeleteSlideTag sld, TAG_NAV_CODE
            End If
        End If
    Next sld

    splitAfter = (sectionNumber + 1) \ 2
    For Each sld In pres.Slides
        If IsTocSlide(sld) Then sld.Tags.Add TAG_TOC_SPLIT_AFTER, CStr(splitAfter)
    Next sld
    RenumberTocDrivenSlides = sectionNumber
End Function

Private Function ResolveTocSectionTitle(ByVal pres As Presentation, ByVal tocSlide As Slide) As String
    Dim visibleTitle As String
    Dim parsedTitle As String
    Dim taggedTitle As String

    visibleTitle = StripNavigationPrefix(GetSlideTitle(tocSlide))
    parsedTitle = SectionTitleAfterDelimiter(visibleTitle)
    If Len(parsedTitle) > 0 Then
        ResolveTocSectionTitle = parsedTitle
        Exit Function
    End If

    If Not IsGenericTocLabel(visibleTitle) Then
        ResolveTocSectionTitle = visibleTitle
        Exit Function
    End If

    taggedTitle = SlideTag(tocSlide, TAG_SECTION_TITLE)
    If Len(taggedTitle) > 0 Then
        ResolveTocSectionTitle = taggedTitle
    Else
        ResolveTocSectionTitle = FirstContentTitleAfterToc(pres, tocSlide)
    End If
End Function

Private Function SectionTitleAfterDelimiter(ByVal value As String) As String
    Dim delimiterPosition As Long

    delimiterPosition = InStrRev(value, ChrW(&HFF1A), -1, vbBinaryCompare)
    If delimiterPosition = 0 Then delimiterPosition = InStrRev(value, ":", -1, vbBinaryCompare)
    If delimiterPosition > 0 Then SectionTitleAfterDelimiter = Trim$(Mid$(value, delimiterPosition + 1))
End Function

Private Function IsGenericTocLabel(ByVal value As String) As Boolean
    Dim normalized As String
    Dim japaneseToc As String

    normalized = LCase$(Trim$(value))
    japaneseToc = ChrW(&H76EE) & ChrW(&H6B21)
    IsGenericTocLabel = (normalized = "contents" Or normalized = "table of contents" Or normalized = "toc" Or value = japaneseToc)
End Function

Private Function FirstContentTitleAfterToc(ByVal pres As Presentation, ByVal tocSlide As Slide) As String
    Dim targetIndex As Long
    Dim target As Slide

    For targetIndex = tocSlide.SlideIndex + 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If IsTocSlide(target) Then Exit For
        If Not IsConfigSlide(target) And Not IsCoverSlide(target) Then
            FirstContentTitleAfterToc = StripNavigationPrefix(GetSlideTitle(target))
            Exit Function
        End If
    Next targetIndex
End Function

Private Sub SetTocVisibleTitle(ByVal tocSlide As Slide, ByVal sectionTitle As String)
    Dim visibleTitle As String
    Dim delimiterPosition As Long
    Dim prefix As String
    Dim separatorSpace As String
    Dim japaneseToc As String

    visibleTitle = StripNavigationPrefix(GetSlideTitle(tocSlide))
    delimiterPosition = InStrRev(visibleTitle, ChrW(&HFF1A), -1, vbBinaryCompare)
    If delimiterPosition = 0 Then delimiterPosition = InStrRev(visibleTitle, ":", -1, vbBinaryCompare)
    If delimiterPosition > 0 Then
        prefix = Left$(visibleTitle, delimiterPosition)
        If Mid$(visibleTitle, delimiterPosition + 1, 1) = " " Then separatorSpace = " "
        SetSlideTitle tocSlide, prefix & separatorSpace & sectionTitle
    ElseIf IsGenericTocLabel(visibleTitle) Then
        japaneseToc = ChrW(&H76EE) & ChrW(&H6B21)
        If visibleTitle = japaneseToc Then
            SetSlideTitle tocSlide, visibleTitle & ChrW(&HFF1A) & sectionTitle
        Else
            SetSlideTitle tocSlide, visibleTitle & ": " & sectionTitle
        End If
    Else
        SetSlideTitle tocSlide, sectionTitle
    End If

End Sub

' Titles: one size for the whole deck, and the room measured rather than assumed.
'
' Two numbers used to be written in here. Both are on the slide and can be
' counted, so neither should have been a number:
'
'   LOGO_CLEARANCE = 72   how much room the master's logo takes on the right.
'                         Measured now: the leftmost shape that actually sits
'                         beside the title. A deck with no logo got 72pt taken
'                         away from it for nothing; a deck with a wider one
'                         had its titles run underneath.
'
'   the size itself       v18 carried 24pt on ten slides and 28pt on forty --
'                         the operator's "24pt vs 28pt". Titles disagree
'                         because each was fitted alone. One size that fits
'                         every title cannot disagree with itself.
'
' The floor (SCITEX_CFG_FONT_MIN) stays a setting. It is not derivable: it says
' how small type may get before it stops being readable, and that is a
' judgement about the reader, not a fact about the slide.
Private Sub FitTocTitles(ByVal pres As Presentation)
    Dim sld As Slide
    Dim titleShape As Shape
    Dim titleRange As TextRange
    Dim availableWidth As Single
    Dim targetSize As Single

    targetSize = UniformTitleSize(pres)

    For Each sld In pres.Slides
        If Not IsConfigSlide(sld) Then
            Set titleShape = FindNamedShape(sld, TITLE_SHAPE)
            If Not titleShape Is Nothing Then
                If titleShape.TextFrame.HasText <> msoTrue Then GoTo ContinueSlide
                availableWidth = TitleRoom(pres, sld, titleShape)
                If availableWidth <= 0 Then Err.Raise vbObjectError + 2114, , _
                    "Title has no usable width on slide " & CStr(sld.SlideIndex) & "."

                ' Keep the title inside the header instead of allowing PowerPoint
                ' to grow the text box over whatever sits beside it.
                titleShape.TextFrame2.AutoSize = msoAutoSizeNone
                titleShape.TextFrame.WordWrap = msoFalse
                titleShape.Width = availableWidth
                Set titleRange = titleShape.TextFrame.TextRange
                titleRange.Font.Size = targetSize

                ' The common size is chosen to fit every title, so this loop
                ' should not run. It stays because a font substitution on
                ' another machine can change the measurement after the fact,
                ' and a title that overflows silently is the bug we started from.
                Do While titleRange.BoundWidth > availableWidth And titleRange.Font.Size > mFontMin
                    titleRange.Font.Size = titleRange.Font.Size - 0.5
                Loop
                If titleRange.BoundWidth > availableWidth Then NoteOverfull sld
            End If
        End If
ContinueSlide:
    Next sld
End Sub

' The largest size at which EVERY title fits its own slide.
'
' Counted, not chosen: try the ceiling, measure all titles, step down until the
' widest one fits. Whatever comes out is the size every title gets, so they
' cannot disagree.
Private Function UniformTitleSize(ByVal pres As Presentation) As Single
    Dim candidate As Single
    Dim sld As Slide
    Dim titleShape As Shape
    Dim titleRange As TextRange
    Dim room As Single
    Dim fitsAll As Boolean

    candidate = mFontMax
    Do
        fitsAll = True
        For Each sld In pres.Slides
            If Not IsConfigSlide(sld) Then
                Set titleShape = FindNamedShape(sld, TITLE_SHAPE)
                If Not titleShape Is Nothing Then
                    If titleShape.TextFrame.HasText = msoTrue Then
                        room = TitleRoom(pres, sld, titleShape)
                        If room > 0 Then
                            titleShape.TextFrame2.AutoSize = msoAutoSizeNone
                            titleShape.TextFrame.WordWrap = msoFalse
                            Set titleRange = titleShape.TextFrame.TextRange
                            titleRange.Font.Size = candidate
                            If titleRange.BoundWidth > room Then fitsAll = False
                        End If
                    End If
                End If
            End If
        Next sld
        If fitsAll Then Exit Do
        candidate = candidate - 0.5
    Loop While candidate > mFontMin

    If candidate < mFontMin Then candidate = mFontMin
    UniformTitleSize = candidate
End Function

' How much horizontal room this title actually has.
'
' The old code subtracted a written-in 72pt for "the logo". What is really
' there is whatever shape sits beside the title, and it is on the slide: take
' the leftmost shape that starts to the right of the title and overlaps it
' vertically. No such shape means the room runs to the slide edge.
Private Function TitleRoom(ByVal pres As Presentation, ByVal sld As Slide, _
                           ByVal titleShape As Shape) As Single
    Dim shp As Shape
    Dim limit As Single
    Dim titleTop As Single
    Dim titleBottom As Single

    limit = pres.PageSetup.SlideWidth
    titleTop = titleShape.Top
    titleBottom = titleShape.Top + titleShape.Height

    For Each shp In sld.Shapes
        If shp.Name <> titleShape.Name Then
            If shp.Left > titleShape.Left Then
                ' vertical overlap, otherwise it is not in the way
                If shp.Top < titleBottom And (shp.Top + shp.Height) > titleTop Then
                    If shp.Left < limit Then limit = shp.Left
                End If
            End If
        End If
    Next shp

    TitleRoom = limit - titleShape.Left - SpaceWidth(titleShape.TextFrame.TextRange)
End Function

' The prefix and the title are separated by a TAB, not by spaces.
'
' The configured face is proportional, so "1a." and "3i." are not the same
' width and no amount of padding lines the titles up -- the operator sees a
' ragged left edge down the whole index. A tab stop is absolute: every title
' starts at the same x regardless of what the prefix measures, and that x is
' the widest prefix in the column -- measured, not written in.
Private Function NavigationEntryText(ByVal sld As Slide) As String
    If IsTocSlide(sld) Then
        NavigationEntryText = SlideTag(sld, TAG_NAV_CODE) & "." & vbTab & SlideTag(sld, TAG_SECTION_TITLE)
    Else
        NavigationEntryText = GetSlideTitle(sld)
    End If
End Function

' Put both index columns on one size -- the smaller, so neither overflows.
'
' Fitting each column separately is correct per column and wrong per slide: a
' 23-entry column shrinks and a 9-entry one does not, and the reader sees two
' type sizes in one list. Counting settles it; there is nothing to choose.
Private Sub MatchColumnSizes(ByVal leftBody As Shape, ByVal rightBody As Shape)
    Dim leftSize As Single
    Dim rightSize As Single
    Dim smaller As Single

    If leftBody.TextFrame.HasText <> msoTrue Then Exit Sub
    If rightBody.TextFrame.HasText <> msoTrue Then Exit Sub

    leftSize = leftBody.TextFrame.TextRange.Characters(1, 1).Font.Size
    rightSize = rightBody.TextFrame.TextRange.Characters(1, 1).Font.Size
    If leftSize <= 0 Or rightSize <= 0 Then Exit Sub

    smaller = leftSize
    If rightSize < smaller Then smaller = rightSize

    leftBody.TextFrame.TextRange.Font.Size = smaller
    rightBody.TextFrame.TextRange.Font.Size = smaller

    ' The tab stop was measured at the old size, so it has to be re-measured
    ' at the new one. Skipping this is how a "fix" leaves the column ragged.
    SetPrefixTabStop leftBody
    SetPrefixTabStop rightBody
End Sub

Private Sub RebuildTocColumn(ByVal pres As Presentation, ByVal tocSlide As Slide, ByVal body As Shape, ByVal isLeftColumn As Boolean, ByVal currentSection As Long)
    Dim target As Slide
    Dim targetIndex As Long
    Dim lineNumber As Long
    Dim targetSection As Long
    Dim splitAfter As Long
    Dim paragraphRange As TextRange
    Dim linkRange As TextRange

    ' The cut is measured, never stored. A written-in split is the defect this
    ' version removes, so a leftover tag is cleared rather than obeyed.
    DeleteSlideTag tocSlide, TAG_TOC_SPLIT_AFTER
    splitAfter = ColumnSplit(pres, tocSlide)

    If LayOutColumn(pres, body, isLeftColumn, splitAfter, mPlannedSize) = 0 Then Exit Sub

    lineNumber = 0
    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If ShouldIncludeInToc(target) Then
            targetSection = CLng(Val(SlideTag(target, TAG_NAV_CODE)))
            If BelongsInColumn(targetSection, isLeftColumn, splitAfter) Then
                lineNumber = lineNumber + 1
                Set paragraphRange = body.TextFrame.TextRange.Paragraphs(lineNumber, 1)
                Set linkRange = EntryRange(paragraphRange)
                ClearTextHyperlink linkRange
                If targetSection = currentSection Then
                    linkRange.Font.Color.RGB = RGB(27, 38, 53)
                    paragraphRange.Font.Color.RGB = RGB(27, 38, 53)
                Else
                    linkRange.Font.Color.RGB = RGB(170, 179, 188)
                    paragraphRange.Font.Color.RGB = RGB(170, 179, 188)
                End If
                AddTocLinkOverlay tocSlide, paragraphRange, target, IIf(isLeftColumn, "L", "R"), lineNumber
            End If
        End If
    Next targetIndex

    ' A last check that the finished column really is inside its box. The
    ' planner measured this exact layout, so this should never shrink anything.
    ' If it does, the measurement and the render have drifted apart, and the
    ' status line says which slide.
    FitTocBody body
End Sub

' One tab stop, placed where the widest prefix actually ends.
'
' There is nothing to choose here. Every prefix is on the slide and can be
' measured, so the stop that clears all of them is the widest one plus a gap.
' A written-in number (34pt was the old one) is a guess that goes stale the
' first time a prefix reaches "3i." from "3a.", or the font changes, and
' nothing tells you it has.
'
' Ruler.TabStops accumulates: rebuilding the index without clearing leaves the
' previous run's stops behind, and the second stop is the one the tab lands on.
Private Sub SetPrefixTabStop(ByVal body As Shape)
    Dim bodyRuler As Ruler
    Dim index As Long
    Dim stopAt As Single

    stopAt = MeasuredPrefixTab(body)
    Set bodyRuler = body.TextFrame.Ruler
    For index = bodyRuler.TabStops.Count To 1 Step -1
        bodyRuler.TabStops(index).Clear
    Next index
    bodyRuler.TabStops.Add ppTabStopLeft, stopAt
    With body.TextFrame.Ruler.Levels(1)
        .FirstMargin = 0
        .LeftMargin = stopAt
    End With
    With body.TextFrame.Ruler.Levels(2)
        .FirstMargin = stopAt
        .LeftMargin = stopAt
    End With
End Sub

' The widest prefix in this column, measured, plus one space.
'
' Called AFTER the text is written, because that is the only moment the widths
' exist. The gap is a measured space in the same run rather than a chosen
' number, so it scales with the font instead of drifting from it.
Private Function MeasuredPrefixTab(ByVal body As Shape) As Single
    Dim para As TextRange
    Dim index As Long
    Dim tabPosition As Long
    Dim widest As Single
    Dim w As Single

    If body.TextFrame.HasText <> msoTrue Then
        MeasuredPrefixTab = 0
        Exit Function
    End If

    For index = 1 To body.TextFrame.TextRange.Paragraphs.Count
        Set para = body.TextFrame.TextRange.Paragraphs(index, 1)
        tabPosition = InStr(1, para.text, vbTab, vbBinaryCompare)
        If tabPosition > 1 Then
            w = para.Characters(1, tabPosition - 1).BoundWidth
            If w > widest Then widest = w
        End If
    Next index

    If widest <= 0 Then
        MeasuredPrefixTab = 0
    Else
        MeasuredPrefixTab = widest + SpaceWidth(body.TextFrame.TextRange)
    End If
End Function

' One space, in the text's own font. Used as the gap after a prefix and after
' a title, so spacing follows the type rather than a constant.
Private Function SpaceWidth(ByVal sample As TextRange) As Single
    Dim probe As Single
    On Error Resume Next
    probe = sample.Characters(1, 1).BoundWidth * 0.45
    On Error GoTo 0
    If probe <= 0 Then probe = 4
    SpaceWidth = probe
End Function

' Shrink the index text until the BOX fits the slide.
'
' The bodies are authored with spAutoFit, which does not clip -- it grows the
' shape. So the failure is never "text outside its box", it is the box itself
' walking off the bottom of the slide while every stored rectangle still looks
' correct. Measured 2026-08-27 on AICHI v11: 15 entries at 18pt leave about
' 154pt of headroom, seven lines, so a few wrapped titles on a machine with
' different font metrics are enough to push it over. Turning autosize off and
' shrinking to fit makes that impossible rather than unlikely.
Private Sub FitTocBody(ByVal body As Shape)
    Dim bodyRange As TextRange
    Dim targetSize As Single
    Dim startSize As Single
    Dim available As Single

    If body.TextFrame.HasText <> msoTrue Then Exit Sub

    body.TextFrame2.AutoSize = msoAutoSizeNone
    body.TextFrame.WordWrap = msoTrue
    available = UsableHeight(body)

    Set bodyRange = body.TextFrame.TextRange
    targetSize = mPlannedSize
    If targetSize <= 0 Or targetSize > mFontMax Then targetSize = mFontMax
    startSize = targetSize
    bodyRange.Font.Size = targetSize

    Do While bodyRange.BoundHeight > available And targetSize > mFontMin
        targetSize = targetSize - 0.5
        If targetSize < mFontMin Then targetSize = mFontMin
        bodyRange.Font.Size = targetSize
    Loop

    ' The tab stop was measured at the planned size. If this had to shrink, the
    ' stop is stale and the prefixes go ragged -- the original complaint.
    If targetSize < startSize Then SetPrefixTabStop body

    If bodyRange.BoundHeight > available Then NoteOverfull body.Parent
End Sub

Private Function BackupName() As String
    Dim separatorPosition As Long
    separatorPosition = InStrRev(mBackupPath, Application.PathSeparator)
    If separatorPosition > 0 Then
        BackupName = Mid$(mBackupPath, separatorPosition + 1)
    Else
        BackupName = mBackupPath
    End If
End Function

Private Sub NoteOverfull(ByVal sld As Slide)
    Dim label As String
    label = CStr(sld.SlideIndex)
    If InStr(1, "," & mOverfullSlides & ",", "," & label & ",") > 0 Then Exit Sub
    If Len(mOverfullSlides) > 0 Then mOverfullSlides = mOverfullSlides & ","
    mOverfullSlides = mOverfullSlides & label
End Sub

' Where to cut the index, and at what size.
'
' The rule the deck has to satisfy is one thing: every entry is on the slide.
' Not "the two columns look even" -- that was this routine's objective until
' 2026-08-28 and it was the wrong one. Even-looking columns with the last entry
' hanging below the slide edge is a failure; a lopsided pair with everything
' visible is not.
'
' So fill from the top of the left column, keep going while it still fits, and
' send the remainder right. Shrink the type only when no cut fits at the current
' size, because small type is a cost to accept last, not a knob to turn first.
'
' Largest size first, and within a size the largest left column, is a search
' over (size, cut). Two facts make it cheap:
'
'   - the left column only grows as the cut moves down, so at a fixed size the
'     largest workable cut is the first fit walking down from the top;
'   - moving the cut back up to help the right column makes the right column
'     BIGGER, not smaller. Once the left fits and the right does not, no other
'     cut at this size will do either -- drop a size instead.
'
' That turns a (sizes x cuts) grid into one walk per size, exiting at the first
' size that works.
Private Function ColumnSplit(ByVal pres As Presentation, ByVal tocSlide As Slide) As Long
    Dim leftBody As Shape
    Dim rightBody As Shape

    If mSplitSlideIndex = tocSlide.SlideIndex And mSplitCached > 0 Then
        ColumnSplit = mSplitCached
        Exit Function
    End If

    Set leftBody = FindNamedShape(tocSlide, TOC_BODY_LEFT)
    Set rightBody = FindNamedShape(tocSlide, TOC_BODY_RIGHT)
    ' No fallback cut here. A guessed number that renders without complaint is
    ' exactly the failure this rewrite removes.
    If leftBody Is Nothing Or rightBody Is Nothing Then
        Err.Raise vbObjectError + 2117, , _
            "Index slide " & tocSlide.SlideIndex & " is missing a column body shape."
    End If

    mSplitCached = PlanColumns(pres, tocSlide, leftBody, rightBody)
    mSplitSlideIndex = tocSlide.SlideIndex
    ColumnSplit = mSplitCached
End Function

' Settle the size and the cut together, since neither is decidable alone.
' Records the size in mPlannedSize for the columns to use; returns the cut.
Private Function PlanColumns(ByVal pres As Presentation, ByVal tocSlide As Slide, _
                             ByVal leftBody As Shape, ByVal rightBody As Shape) As Long
    Dim highest As Long
    Dim trySize As Single
    Dim cut As Long
    Dim fittingCut As Long
    Dim leftRoom As Single
    Dim rightRoom As Single

    highest = HighestSection(pres)
    If highest <= 1 Then
        mPlannedSize = mFontMax
        PlanColumns = 1
        Exit Function
    End If

    leftRoom = UsableHeight(leftBody)
    rightRoom = UsableHeight(rightBody)

    trySize = mFontMax
    Do While trySize >= mFontMin
        fittingCut = LargestFittingCut(pres, tocSlide, leftBody, highest, trySize, leftRoom)
        If fittingCut > 0 Then
            If ColumnHeightAt(pres, tocSlide, rightBody, False, fittingCut, trySize) <= rightRoom Then
                mPlannedSize = trySize
                PlanColumns = fittingCut
                Exit Function
            End If
        End If
        trySize = trySize - 0.5
    Loop

    ' Nothing fits even at the smallest size the config allows. Name the slide
    ' on the status line rather than cropping quietly, and give the left column
    ' as much as it can hold so the overflow is at one end instead of both.
    mPlannedSize = mFontMin
    NoteOverfull tocSlide
    fittingCut = LargestFittingCut(pres, tocSlide, leftBody, highest, mFontMin, leftRoom)
    If fittingCut <= 0 Then fittingCut = 1
    PlanColumns = fittingCut
End Function

' The lowest cut whose left column still fits, or 0 if none does.
Private Function LargestFittingCut(ByVal pres As Presentation, ByVal tocSlide As Slide, _
                                   ByVal leftBody As Shape, ByVal highest As Long, _
                                   ByVal fontSize As Single, ByVal room As Single) As Long
    Dim cut As Long

    For cut = highest - 1 To 1 Step -1
        If ColumnHeightAt(pres, tocSlide, leftBody, True, cut, fontSize) <= room Then
            LargestFittingCut = cut
            Exit Function
        End If
    Next cut
    LargestFittingCut = 0
End Function

Private Function HighestSection(ByVal pres As Presentation) As Long
    Dim sld As Slide
    Dim section As Long

    For Each sld In pres.Slides
        If ShouldIncludeInToc(sld) Then
            section = CLng(Val(SlideTag(sld, TAG_NAV_CODE)))
            If section > HighestSection Then HighestSection = section
        End If
    Next sld
End Function

' The height the text may occupy, which is not the height of the box.
Private Function UsableHeight(ByVal body As Shape) As Single
    UsableHeight = body.Height - body.TextFrame.MarginTop - body.TextFrame.MarginBottom
    If UsableHeight <= 0 Then Err.Raise vbObjectError + 2115, , _
        "Index body " & body.Name & " has no usable height."
End Function

' Fill one column for a trial cut at a trial size and report how tall it came out.
'
' This goes through LayOutColumn, the same routine that builds the real column,
' because a trial that skips the indent, the weight or the tab stop measures a
' column that will never be rendered -- and reports a fit for text that then
' hangs off the slide. The trial text is thrown away; RebuildTocColumn writes
' the final one with colours and links straight after.
Private Function ColumnHeightAt(ByVal pres As Presentation, ByVal tocSlide As Slide, _
                                ByVal body As Shape, ByVal isLeftColumn As Boolean, _
                                ByVal splitAfter As Long, ByVal fontSize As Single) As Single
    If LayOutColumn(pres, body, isLeftColumn, splitAfter, fontSize) = 0 Then
        ColumnHeightAt = 0
        Exit Function
    End If
    ColumnHeightAt = body.TextFrame.TextRange.BoundHeight
End Function

' Write one column's entries and give them the shape they will have on the
' slide: size, weight, indent level, tab stop. Returns the entry count.
'
' Trial fills and the final build both come through here. Keeping one routine
' is what makes the measurement above trustworthy -- when this changes, what
' was measured changes with it.
'
' Order matters. The weight is applied BEFORE the tab stop is measured, because
' a section line's prefix is bold and a stop measured on the regular weight
' lands short of it.
Private Function LayOutColumn(ByVal pres As Presentation, ByVal body As Shape, _
                              ByVal isLeftColumn As Boolean, ByVal splitAfter As Long, _
                              ByVal fontSize As Single) As Long
    Dim target As Slide
    Dim targetIndex As Long
    Dim targetSection As Long
    Dim tocText As String
    Dim lineNumber As Long
    Dim paragraphRange As TextRange

    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If ShouldIncludeInToc(target) Then
            targetSection = CLng(Val(SlideTag(target, TAG_NAV_CODE)))
            If BelongsInColumn(targetSection, isLeftColumn, splitAfter) Then
                If Len(tocText) > 0 Then tocText = tocText & vbCrLf
                tocText = tocText & NavigationEntryText(target)
                LayOutColumn = LayOutColumn + 1
            End If
        End If
    Next targetIndex

    body.TextFrame2.AutoSize = msoAutoSizeNone
    body.TextFrame.WordWrap = msoTrue
    body.TextFrame.TextRange.text = tocText
    If LayOutColumn = 0 Then Exit Function

    body.TextFrame.TextRange.Font.Size = fontSize
    body.TextFrame.TextRange.Font.Bold = msoFalse
    body.TextFrame.TextRange.Font.Underline = msoFalse

    lineNumber = 0
    For targetIndex = 1 To pres.Slides.Count
        Set target = pres.Slides(targetIndex)
        If ShouldIncludeInToc(target) Then
            targetSection = CLng(Val(SlideTag(target, TAG_NAV_CODE)))
            If BelongsInColumn(targetSection, isLeftColumn, splitAfter) Then
                lineNumber = lineNumber + 1
                Set paragraphRange = body.TextFrame.TextRange.Paragraphs(lineNumber, 1)
                If IsTocSlide(target) Then
                    paragraphRange.IndentLevel = 1
                    EntryRange(paragraphRange).Font.Bold = msoTrue
                    EntryRange(paragraphRange).Font.Underline = msoTrue
                Else
                    paragraphRange.IndentLevel = 2
                End If
            End If
        End If
    Next targetIndex

    ' The hanging indent and the tab stop are the same measurement, so one
    ' routine owns both. Setting them apart is how they drift.
    SetPrefixTabStop body
End Function

Private Function BelongsInColumn(ByVal section As Long, ByVal isLeftColumn As Boolean, _
                                 ByVal splitAfter As Long) As Boolean
    If isLeftColumn Then
        BelongsInColumn = (section <= splitAfter)
    Else
        BelongsInColumn = (section > splitAfter)
    End If
End Function

' A paragraph without its trailing newline. Styling the newline is what leaves
' an underline running past the end of a line.
Private Function EntryRange(ByVal paragraphRange As TextRange) As TextRange
    Dim entryLength As Long
    Dim lastCharacter As String

    entryLength = Len(paragraphRange.text)
    Do While entryLength > 0
        lastCharacter = Mid$(paragraphRange.text, entryLength, 1)
        If lastCharacter <> vbCr And lastCharacter <> vbLf Then Exit Do
        entryLength = entryLength - 1
    Loop
    Set EntryRange = paragraphRange.Characters(1, entryLength)
End Function

' Give both columns the same vertical band before anything is measured.
'
' In v18 the left box is 826pt tall and the right 265pt, on every index slide.
' Neither number was chosen: spAutoFit grew each box to its own content, so the
' column with more entries ended up with more room, which is backwards. Both
' get the same top and the same height -- the band from the top of the boxes to
' the bottom of the slide -- and then the fit is a fair question.
'
' Widths are left alone. They differ too (363.6 vs 291.6), but that is a layout
' decision someone made, not an artefact, and the measurement above accounts
' for it.
Private Sub NormaliseTocBoxes(ByVal pres As Presentation, _
                              ByVal leftBody As Shape, ByVal rightBody As Shape)
    Dim topEdge As Single
    Dim band As Single

    topEdge = leftBody.Top
    If rightBody.Top < topEdge Then topEdge = rightBody.Top

    band = pres.PageSetup.SlideHeight - topEdge - BottomMargin(pres)
    If band <= 0 Then Err.Raise vbObjectError + 2116, , _
        "No vertical room for the index columns."

    leftBody.TextFrame2.AutoSize = msoAutoSizeNone
    rightBody.TextFrame2.AutoSize = msoAutoSizeNone
    leftBody.Top = topEdge
    rightBody.Top = topEdge
    leftBody.Height = band
    rightBody.Height = band
End Sub

' How much to keep clear below the index. Derived from the slide, not chosen:
' the same proportion the top margin already uses.
Private Function BottomMargin(ByVal pres As Presentation) As Single
    BottomMargin = pres.PageSetup.SlideHeight * 0.02
End Function

Private Sub ClearTextHyperlink(ByVal textRange As TextRange)
    On Error Resume Next
    textRange.ActionSettings(ppMouseClick).Hyperlink.Delete
    textRange.ActionSettings(ppMouseClick).Action = ppActionNone
    On Error GoTo 0
End Sub

Private Sub DeleteTocLinkOverlays(ByVal tocSlide As Slide)
    Dim shapeIndex As Long
    Dim shapeName As String

    For shapeIndex = tocSlide.Shapes.Count To 1 Step -1
        shapeName = tocSlide.Shapes(shapeIndex).Name
        If StrComp(Left$(shapeName, Len(TOC_LINK_PREFIX)), TOC_LINK_PREFIX, vbTextCompare) = 0 Then
            tocSlide.Shapes(shapeIndex).Delete
        ElseIf tocSlide.Shapes(shapeIndex).Type = msoLine Then
            ' Copied TOC slides can retain old manually drawn rules. The TOC
            ' generator owns this slide content, so remove those stale lines.
            tocSlide.Shapes(shapeIndex).Delete
        End If
    Next shapeIndex
End Sub

Private Sub AddTocLinkOverlay(ByVal tocSlide As Slide, ByVal paragraphRange As TextRange, ByVal target As Slide, ByVal columnKey As String, ByVal lineNumber As Long)
    Dim overlay As Shape
    Dim overlayName As String

    overlayName = TOC_LINK_PREFIX & columnKey & "_" & CStr(lineNumber)
    Set overlay = tocSlide.Shapes.AddShape(msoShapeRectangle, paragraphRange.BoundLeft, paragraphRange.BoundTop, paragraphRange.BoundWidth, paragraphRange.BoundHeight)
    overlay.Name = overlayName
    overlay.Fill.Visible = msoTrue
    overlay.Fill.Solid
    overlay.Fill.Transparency = 1
    overlay.Line.Visible = msoFalse
    overlay.AlternativeText = "Link to " & NavigationEntryText(target)
    With overlay.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.SubAddress = CStr(target.SlideID) & "," & CStr(target.SlideIndex) & "," & NavigationEntryText(target)
    End With
    overlay.ZOrder msoBringToFront
End Sub

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
    ElseIf Len(SlideTag(sld, TAG_NAV_CODE)) = 0 Then
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
                        If IsTitleShape(shp) Then
                            ' Bind to the THEME fonts rather than stamping a
                            ' literal face. Writing "Arial" here is what made
                            ' the titles disagree with the master: the master
                            ' asks for the theme major font and this line
                            ' overwrote it on every run.
                            textRange.Font.Name = "+mj-lt"
                            On Error Resume Next
                            textRange.Font.NameFarEast = "+mj-ea"
                            On Error GoTo 0
                        Else
                            textRange.Font.Name = mFontLatin
                            On Error Resume Next
                            textRange.Font.NameFarEast = mFontCjk
                            On Error GoTo 0
                        End If
                        originalSize = CSng(Val(ShapeTag(shp, TAG_ORIGINAL_FONT_SIZE)))
                        If originalSize <= 0 Then
                            originalSize = textRange.Characters(1, 1).Font.Size
                            If originalSize <= 0 Then originalSize = mFontMin
                            shp.Tags.Add TAG_ORIGINAL_FONT_SIZE, CStr(originalSize)
                        End If
                        ' Titles are sized by FitTocTitles, which picks one
                        ' size that fits every title in the deck. Setting a
                        ' size here too meant two routines deciding the same
                        ' thing, and the second one silently winning.
                        If Not IsTitleShape(shp) Then
                            targetSize = originalSize
                            If targetSize < mFontMin Then targetSize = mFontMin
                            If targetSize > mFontMax Then targetSize = mFontMax
                            textRange.Font.Size = targetSize
                        End If
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

Private Function IsTitleShape(ByVal shp As Shape) As Boolean
    IsTitleShape = (StrComp(shp.Name, TITLE_SHAPE, vbTextCompare) = 0)
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

Private Sub DeleteSlideTag(ByVal sld As Slide, ByVal tagName As String)
    On Error Resume Next
    sld.Tags.Delete tagName
    On Error GoTo 0
End Sub

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
