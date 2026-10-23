param([Parameter(Mandatory = $true)][string]$JobPath)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$job = Get-Content -LiteralPath $JobPath -Raw -Encoding UTF8 | ConvertFrom-Json
$application = $null
$presentation = $null
$initialCount = 0

function Get-TextMetric($shape, [int]$shapeId, $row = $null, $column = $null) {
    if ($shape.HasTextFrame -eq 0 -or $shape.TextFrame2.HasText -eq 0) { return $null }
    $frame = $shape.TextFrame2
    $range = $frame.TextRange
    $width = [double]$shape.Width - [double]$frame.MarginLeft - [double]$frame.MarginRight
    $height = [double]$shape.Height - [double]$frame.MarginTop - [double]$frame.MarginBottom
    $font = [double]$range.Characters(1, 1).Font.Size
    return @{
        shape_id = $shapeId; row = $row; column = $column
        bound_width_pt = [double]$range.BoundWidth; bound_height_pt = [double]$range.BoundHeight
        available_width_pt = $width; available_height_pt = $height
        base_font_pt = $font
        overflow = ([double]$range.BoundHeight -gt ($height + 1.0) -or [double]$range.BoundWidth -gt ($width + 1.0))
    }
}

function Get-ShapeMetrics($shapes) {
    $metrics = [System.Collections.Generic.List[object]]::new()
    for ($index = 1; $index -le $shapes.Count; $index++) {
        $shape = $shapes.Item($index)
        if ($shape.Type -eq 6) {
            foreach ($metric in (Get-ShapeMetrics $shape.GroupItems)) { $metrics.Add($metric) }
        } elseif ($shape.HasTable -ne 0) {
            for ($row = 1; $row -le $shape.Table.Rows.Count; $row++) {
                for ($column = 1; $column -le $shape.Table.Columns.Count; $column++) {
                    $metric = Get-TextMetric $shape.Table.Cell($row, $column).Shape $shape.Id $row $column
                    if ($null -ne $metric) { $metrics.Add($metric) }
                }
            }
        } else {
            $metric = Get-TextMetric $shape $shape.Id
            if ($null -ne $metric) { $metrics.Add($metric) }
        }
    }
    return $metrics.ToArray()
}

try {
    $application = New-Object -ComObject PowerPoint.Application
    $initialCount = $application.Presentations.Count
    $presentation = $application.Presentations.Open($job.source, -1, 0, 0)
    $height = [int][Math]::Round($job.width * $presentation.PageSetup.SlideHeight / $presentation.PageSetup.SlideWidth)
    $slides = [System.Collections.Generic.List[object]]::new()
    for ($page = 1; $page -le $presentation.Slides.Count; $page++) {
        if ($job.pages.Count -gt 0 -and $job.pages -notcontains $page) { continue }
        $slide = $presentation.Slides.Item($page)
        $filename = 'slide_{0:D3}.png' -f $page
        $slide.Export((Join-Path $job.output $filename), 'PNG', [int]$job.width, $height)
        $slides.Add(@{ page = $page; image = $filename; text_metrics = @(Get-ShapeMetrics $slide.Shapes) })
    }
    @{ slides = $slides.ToArray() } | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $job.output 'render.json') -Encoding UTF8
} finally {
    if ($null -ne $presentation) { $presentation.Close() }
    if ($null -ne $application -and $initialCount -eq 0) { $application.Quit() }
    if ($null -ne $presentation) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($null -ne $application) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
