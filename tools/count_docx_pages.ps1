$ErrorActionPreference = "Stop"

function Invoke-ComRetry {
  param([scriptblock]$Block)

  $lastError = $null
  for ($i = 0; $i -lt 60; $i++) {
    try {
      return & $Block
    } catch [System.Runtime.InteropServices.COMException] {
      $lastError = $_
      Start-Sleep -Milliseconds 500
    }
  }

  throw $lastError
}

$path = (Resolve-Path $args[0]).Path
$word = $null
$doc = $null
$wordProcessId = $null
$existingWordProcessIds = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $currentWordProcessIds = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $wordProcessId = $currentWordProcessIds | Where-Object { $existingWordProcessIds -notcontains $_ } | Select-Object -First 1

  $doc = Invoke-ComRetry { $word.Documents.Open($path, $false, $true) }
  Invoke-ComRetry { $doc.Repaginate() } | Out-Null
  Start-Sleep -Seconds 1
  $pages = Invoke-ComRetry { $doc.ComputeStatistics(2) }
  Write-Output $pages
} finally {
  if ($doc -ne $null) {
    try {
      Invoke-ComRetry { $doc.Close($false) } | Out-Null
    } catch {
      Write-Warning "No se pudo cerrar el documento Word: $($_.Exception.Message)"
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
  }

  if ($word -ne $null) {
    try {
      Invoke-ComRetry { $word.Quit() } | Out-Null
    } catch {
      Write-Warning "No se pudo cerrar Word: $($_.Exception.Message)"
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }

  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()

  if ($wordProcessId -ne $null) {
    Stop-Process -Id $wordProcessId -Force -ErrorAction SilentlyContinue
  }
}
