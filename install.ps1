# Clause Craft one-step setup for Windows.
# Downloads Clause Craft to Documents\Clause Craft, adds it to desktop Word for this user
# (no admin rights needed) and opens Word with the Clause Craft panel showing.
#
# Run it from the Windows Run box (Windows key + R), then press Enter:
#   powershell -c "irm https://chadnasir.github.io/clause-craft-cre/install.ps1|iex"

$ErrorActionPreference = 'Stop'
try {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
  $docs = [Environment]::GetFolderPath('MyDocuments')
  $folder = Join-Path $docs 'Clause Craft'
  $work = Join-Path $env:TEMP ('ClauseCraft-' + [Guid]::NewGuid().ToString('N'))
  $zip = Join-Path $work 'Clause-Craft-for-Windows.zip'
  New-Item -ItemType Directory -Path $work | Out-Null

  Write-Host ''
  Write-Host '  Downloading Clause Craft...'
  # The query string only bypasses the web cache, so a just-published update is picked up.
  Invoke-WebRequest -Uri ('https://chadnasir.github.io/clause-craft-cre/download/Clause-Craft-for-Windows.zip?t=' + [DateTime]::UtcNow.Ticks) -OutFile $zip -UseBasicParsing
  Expand-Archive -Path $zip -DestinationPath $work

  # Copy into Documents\Clause Craft. A file that is open in Word (the demo lease) is left as it is.
  New-Item -ItemType Directory -Force -Path $folder | Out-Null
  $unzipped = Join-Path $work 'Clause Craft'
  foreach ($file in Get-ChildItem $unzipped -Recurse -File) {
    $target = Join-Path $folder $file.FullName.Substring($unzipped.Length + 1)
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    try { Copy-Item $file.FullName $target -Force } catch { Write-Host ('  Kept your open copy of ' + $file.Name) }
  }
  Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
  Write-Host "  Saved in $folder"

  $installer = Join-Path $folder 'Install in Word.cmd'
  if (-not (Test-Path $installer)) { throw "The download didn't include Install in Word.cmd." }
  Start-Process -FilePath $installer -NoNewWindow -Wait
}
catch {
  Write-Host ''
  Write-Host ('  Clause Craft could not be set up: ' + $_.Exception.Message)
  Write-Host '  Check the internet connection and try again, or use the download page:'
  Write-Host '  https://chadnasir.github.io/clause-craft-cre/download/'
  Write-Host ''
  Read-Host '  Press Enter to close'
}
