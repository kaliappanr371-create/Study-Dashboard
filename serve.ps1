# Local server using only built-in Windows PowerShell (no Python needed). Right-click > Run with PowerShell.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$l = New-Object System.Net.HttpListener
$l.Prefixes.Add('http://localhost:8765/')
$l.Start()
Write-Host 'Study Dashboard is running at http://localhost:8765  (close this window to stop)'
Start-Process 'http://localhost:8765'
$mime = @{'.html'='text/html';'.js'='text/javascript';'.css'='text/css';'.json'='application/json';'.webmanifest'='application/manifest+json';'.png'='image/png';'.ico'='image/x-icon'}
while ($l.IsListening) {
  $c = $l.GetContext()
  $p = [Uri]::UnescapeDataString($c.Request.Url.AbsolutePath).TrimStart('/')
  if ($p -eq '') { $p = 'index.html' }
  $f = Join-Path $root $p
  if ((Test-Path $f -PathType Leaf) -and ($f.StartsWith($root))) {
    $b = [IO.File]::ReadAllBytes($f); $e = [IO.Path]::GetExtension($f).ToLower()
    $c.Response.ContentType = $(if ($mime[$e]) { $mime[$e] } else { 'application/octet-stream' })
    $c.Response.Headers.Add('Cache-Control','no-cache')
    $c.Response.OutputStream.Write($b, 0, $b.Length)
  } else { $c.Response.StatusCode = 404 }
  $c.Response.Close()
}
