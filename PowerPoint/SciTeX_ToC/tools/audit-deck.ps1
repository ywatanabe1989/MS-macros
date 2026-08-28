param(
    [Parameter(Mandatory = $true)][string]$Deck
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.WindowState = 2
    $presentation = $powerPoint.Presentations.Open($Deck, $true, $false, $true)

    Write-Output ("DECK slides={0} masters={1} layouts={2} width={3} height={4}" -f $presentation.Slides.Count, $presentation.SlideMasters.Count, $presentation.SlideMaster.CustomLayouts.Count, $presentation.PageSetup.SlideWidth, $presentation.PageSetup.SlideHeight)
    foreach ($slide in $presentation.Slides) {
        $textShapes = @()
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = ($shape.TextFrame.TextRange.Text -replace "`r|`n", " / ").Trim()
                if ($text.Length -gt 90) {
                    $text = $text.Substring(0, 90) + "..."
                }
                $textShapes += [pscustomobject]@{
                    Name = $shape.Name
                    Top = [math]::Round($shape.Top, 1)
                    Left = [math]::Round($shape.Left, 1)
                    Size = $shape.TextFrame.TextRange.Font.Size
                    Text = $text
                }
            }
        }
        $candidateText = ($textShapes | Sort-Object Top, Left | Select-Object -First 5 | ForEach-Object { "[{0}|top={1}|size={2}] {3}" -f $_.Name, $_.Top, $_.Size, $_.Text }) -join " || "
        Write-Output ("SLIDE {0:D2} id={1} hidden={2} shapes={3} text_shapes={4} tags={5} :: {6}" -f $slide.SlideIndex, $slide.SlideID, ($slide.SlideShowTransition.Hidden -ne 0), $slide.Shapes.Count, $textShapes.Count, $slide.Tags.Count, $candidateText)
    }
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
