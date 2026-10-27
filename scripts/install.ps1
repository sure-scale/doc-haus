# doc.haus installer for Windows:
#   irm https://doc.haus/install.ps1 | iex
# Installs Git when it is missing, downloads doc.haus to %USERPROFILE%\doc-haus
# (or DOCHAUS_DIR) and starts it with start.cmd. Running it again starts the
# existing copy. It runs inside the caller's PowerShell, so it returns rather
# than exits and keeps its settings inside the function.
function Install-DocHaus {
  $ErrorActionPreference = 'Stop'
  $repository = 'https://github.com/sure-scale/doc-haus.git'
  $dir = if ($env:DOCHAUS_DIR) { $env:DOCHAUS_DIR } else { Join-Path $HOME 'doc-haus' }

  if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
      Write-Host 'doc.haus needs Git. Install it from https://git-scm.com/downloads/win, then run this command again.'
      return
    }
    Write-Host 'doc.haus: installing Git. Windows may ask for permission.'
    winget install --id Git.Git -e --source winget --accept-source-agreements --accept-package-agreements
    # The Git installer updates PATH for new windows; read it again for this one.
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
      Write-Host 'doc.haus could not install Git. Install it from https://git-scm.com/downloads/win, then run this command again.'
      return
    }
  }

  if (Test-Path $dir) {
    $origin = git -C $dir remote get-url origin 2>$null
    if ($origin -notmatch 'sure-scale/doc-haus') {
      Write-Host "doc.haus: $dir already exists and is not a copy of doc.haus. Set DOCHAUS_DIR to another folder, then run this command again."
      return
    }
    Write-Host "doc.haus: starting the copy in $dir"
  } else {
    git clone $repository $dir
    if ($LASTEXITCODE -ne 0) { Write-Host 'doc.haus: the download failed. Check your connection, then run this command again.'; return }
  }

  # A Start menu entry to start doc.haus again later. cmd /k keeps the window,
  # and any message in it, open after doc.haus stops. It shows the logo from
  # scripts\icons, which copies made before it was added do not have.
  $dir = (Resolve-Path $dir).Path
  $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path ([Environment]::GetFolderPath('Programs')) 'doc.haus.lnk'))
  $shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\cmd.exe'
  $shortcut.Arguments = '/k "' + (Join-Path $dir 'start.cmd') + '"'
  $shortcut.WorkingDirectory = $dir
  $shortcut.Description = 'Start doc.haus'
  $icon = Join-Path $dir 'scripts\icons\doc-haus.ico'
  if (Test-Path $icon) { $shortcut.IconLocation = "$icon,0" }
  $shortcut.Save()

  $previousWordSetup = $env:DOCHAUS_SETUP_WORD_ADDIN
  try {
    $env:DOCHAUS_SETUP_WORD_ADDIN = '1'
    & (Join-Path $dir 'start.cmd')
  } finally { $env:DOCHAUS_SETUP_WORD_ADDIN = $previousWordSetup }
}
Install-DocHaus
