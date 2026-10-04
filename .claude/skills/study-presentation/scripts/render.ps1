# 발표자료 시각 검수용 렌더링
#   powershell -ExecutionPolicy Bypass -File render.ps1 -Pptx <deck.pptx> -OutDir <dir>
# PowerPoint(COM)로 슬라이드를 PNG로 내보내고, 한눈에 보는 contact sheet(sheet-N.png)를 만든다.
# PowerPoint가 파일을 여는 단계에서 실패하면 pptx 구조가 깨진 것이다.
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [int]$Width = 1600
)
$ErrorActionPreference = "Stop"
$Pptx = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
Get-ChildItem $OutDir -Filter "slide*.png" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem $OutDir -Filter "sheet*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$height = [int]($Width * 9 / 16)
$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($Pptx, $true, $false, $false)
  $i = 1
  foreach ($s in $pres.Slides) {
    $s.Export((Join-Path $OutDir ("slide{0:D2}.png" -f $i)), "PNG", $Width, $height)
    $i++
  }
  $pres.Close()
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}

# contact sheet: 2열 x 4행, 8장씩
Add-Type -AssemblyName System.Drawing
$files = Get-ChildItem $OutDir -Filter "slide*.png" | Sort-Object Name
$tw = 800; $th = 450; $pad = 12
$per = 8
for ($p = 0; $p -lt [math]::Ceiling($files.Count / $per); $p++) {
  $chunk = $files | Select-Object -Skip ($p * $per) -First $per
  $rows = [math]::Ceiling($chunk.Count / 2)
  $bmp = New-Object System.Drawing.Bitmap ((2 * $tw) + (3 * $pad)), (($rows * $th) + (($rows + 1) * $pad))
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::FromArgb(200, 200, 200))
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $font = New-Object System.Drawing.Font "Arial", 14, ([System.Drawing.FontStyle]::Bold)
  $k = 0
  foreach ($f in $chunk) {
    $img = [System.Drawing.Image]::FromFile($f.FullName)
    $x = $pad + ($k % 2) * ($tw + $pad); $y = $pad + [math]::Floor($k / 2) * ($th + $pad)
    $g.DrawImage($img, $x, $y, $tw, $th)
    $g.DrawString($f.BaseName, $font, [System.Drawing.Brushes]::Red, ($x + 6), ($y + 4))
    $img.Dispose(); $k++
  }
  $bmp.Save((Join-Path $OutDir ("sheet-{0}.png" -f ($p + 1))), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}
Get-ChildItem $OutDir -Filter "*.png" | Sort-Object Name | ForEach-Object { $_.FullName }
