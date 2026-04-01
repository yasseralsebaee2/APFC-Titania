$html = Get-Content "c:\Users\yasser.alsebaee\Downloads\apfc_dashboard.html" -Raw

$styleStart = $html.IndexOf("<style>")
$styleEnd = $html.IndexOf("</style>")
if ($styleStart -ne -1 -and $styleEnd -ne -1) {
    $styleStart += 7
    $styleContent = $html.Substring($styleStart, $styleEnd - $styleStart).Trim()
    Set-Content -Path "css\style.css" -Value $styleContent -Encoding UTF8
    Write-Host "CSS Extracted"
}

$scriptStart = $html.IndexOf("<script>")
$scriptEnd = $html.LastIndexOf("</script>")
if ($scriptStart -ne -1 -and $scriptEnd -ne -1) {
    $scriptStart += 8
    $scriptContent = $html.Substring($scriptStart, $scriptEnd - $scriptStart).Trim()
    Set-Content -Path "js\app.js" -Value $scriptContent -Encoding UTF8
    Write-Host "JS Extracted"
}

$bodyStart = $html.IndexOf("<body>")
$bodyEnd = $html.IndexOf("</body>")
if ($bodyStart -ne -1 -and $bodyEnd -ne -1) {
    $bodyHtml = $html.Substring($bodyStart, ($bodyEnd + 7) - $bodyStart)
    
    # We find the script block inside the body
    $innerScriptStart = $bodyHtml.IndexOf("<script>")
    $innerScriptEnd = $bodyHtml.LastIndexOf("</script>")
    if ($innerScriptStart -ne -1 -and $innerScriptEnd -ne -1) {
        $innerScriptBlock = $bodyHtml.Substring($innerScriptStart, ($innerScriptEnd + 9) - $innerScriptStart)
        $bodyHtml = $bodyHtml.Replace($innerScriptBlock, "")
    }
    
    $indexHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>APFC Dashboard | Project Dashboard</title>
  
  <meta name="theme-color" content="#06080b">
  <link rel="manifest" href="manifest.json">
  <link rel="apple-touch-icon" href="assets/APFC_Logo.png">
  
  <link rel="stylesheet" href="css/style.css" />
</head>
$bodyHtml
  <script src="js/app.js"></script>
  <script>
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
          .then(reg => console.log('Service Worker registered', reg))
          .catch(err => console.error('SW registration failed:', err));
      });
    }
  </script>
</html>
"@
    Set-Content -Path "index.html" -Value $indexHtml -Encoding UTF8
    Write-Host "HTML Created"
}
