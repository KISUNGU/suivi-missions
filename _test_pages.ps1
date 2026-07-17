foreach ($p in @('index.html','saisie.html','admin.html','superadmin.html')) {
  try {
    $resp = Invoke-WebRequest -Uri ("http://localhost:3000/" + $p) -UseBasicParsing
    Write-Host ($p + " -> " + $resp.StatusCode + " (" + $resp.Content.Length + " bytes)")
  } catch {
    Write-Host ($p + " -> ERROR " + $_.Exception.Message)
  }
}
