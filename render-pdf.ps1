# render-pdf.ps1 — 用 Windows 內建 PDF 引擎把 PDF 每頁轉成 PNG（in-process，不需額外安裝）
# 用法:
#   powershell -File render-pdf.ps1 -Pdf 簡報.pdf [-OutDir .preview] [-Width 1600]
# 產出: <OutDir>\page01.png ... pageNN.png

param(
  [Parameter(Mandatory = $true)][string]$Pdf,
  [string]$OutDir = ".preview",
  [int]$Width = 1600
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapEncoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]

$asTaskOp = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1
$asTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq "IAsyncAction"
} | Select-Object -First 1

function Await-Op([object]$WinRtTask, [type]$ResultType) {
  $asTask = $asTaskOp.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  return $netTask.Result
}

function Await-Action([object]$WinRtTask) {
  $netTask = $asTaskAction.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
}

$pdfPath = (Resolve-Path $Pdf).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$file = Await-Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath)) ([Windows.Storage.StorageFile])
$doc = Await-Op ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
Write-Output "PDF pages: $($doc.PageCount)"

for ($i = 0; $i -lt $doc.PageCount; $i++) {
  $page = $doc.GetPage($i)
  $stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
  $opts = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
  $opts.DestinationWidth = [uint32]$Width
  Await-Action ($page.RenderToStreamAsync($stream, $opts))

  $decoder = Await-Op ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bitmap = Await-Op ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

  $outStream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
  $encoder = Await-Op ([Windows.Graphics.Imaging.BitmapEncoder]::CreateAsync([Windows.Graphics.Imaging.BitmapEncoder]::PngEncoderId, $outStream)) ([Windows.Graphics.Imaging.BitmapEncoder])
  $encoder.SetSoftwareBitmap($bitmap)
  Await-Action ($encoder.FlushAsync())

  $outPath = Join-Path $OutDir ("page{0:D2}.png" -f ($i + 1))
  $fs = [System.IO.File]::Create($outPath)
  try {
    $outStream.Seek(0)
    $readStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($outStream)
    $readStream.CopyTo($fs)
  } finally {
    $fs.Dispose()
  }
  $page.Dispose()
  Write-Output ("page{0:D2}.png done" -f ($i + 1))
}
Write-Output "DONE: $($doc.PageCount) pages -> $((Resolve-Path $OutDir).Path)"
