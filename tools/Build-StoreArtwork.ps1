param([Parameter(Mandatory = $true)][string]$ListingDirectory)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$menu = Join-Path (Split-Path -Parent $PSScriptRoot) 'ModernMenu'
[void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($ListingDirectory))

# Original tile artwork: two folded ribbons converge on a single square.
# Draw from the same 100-unit grid at each target resolution.
$shapes = @(
    @{ Color = '#FAF5E6'; XY = @(16,28, 37,28, 51,42, 60,42, 60,49, 45,49, 31,36, 16,36) }
    @{ Color = '#6DE4C1'; XY = @(16,64, 31,64, 45,51, 60,51, 60,58, 51,58, 37,72, 16,72) }
    @{ Color = '#FFD385'; XY = @(64,42, 80,42, 80,58, 64,58) }
)
foreach ($size in 44,50,150,300) {
    $image = [Drawing.Bitmap]::new($size, $size)
    $g = [Drawing.Graphics]::FromImage($image)
    try {
        $g.SmoothingMode = 'AntiAlias'
        $g.Clear([Drawing.ColorTranslator]::FromHtml('#18334B'))
        $g.ScaleTransform(($size / 100.0), ($size / 100.0))
        foreach ($shape in $shapes) {
            $points = for ($i = 0; $i -lt $shape.XY.Count; $i += 2) {
                [Drawing.PointF]::new($shape.XY[$i], $shape.XY[$i + 1])
            }
            $brush = [Drawing.SolidBrush]::new([Drawing.ColorTranslator]::FromHtml($shape.Color))
            try { $g.FillPolygon($brush, [Drawing.PointF[]]$points) } finally { $brush.Dispose() }
        }
        $destination = if ($size -eq 300) { Join-Path $ListingDirectory 'Store-icon-300.png' }
                       else { Join-Path $menu "Logo$size.png" }
        $image.Save($destination, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $g.Dispose(); $image.Dispose() }
}
