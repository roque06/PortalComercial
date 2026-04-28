$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try {
  # Debe ejecutarse antes de crear cualquier control WinForms.
  [System.Windows.Forms.Application]::SetUnhandledExceptionMode(
    [System.Windows.Forms.UnhandledExceptionMode]::CatchException
  )
} catch {
  # Si falla por contexto de host, continuamos; los handlers de abajo siguen activos.
}

$script:scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:projectRoot = Split-Path -Parent $script:scriptDir
$script:currentProcess = $null
$script:outputDataHandler = $null
$script:errorDataHandler = $null
$script:currentMode = ""
$script:currentLogFile = ""
$script:cmbBizagi = $null
$script:bizagiUrls = @(
  "https://dev-bscrd-santacruz.bizagi.com/",
  "https://srvdevbizapp01/BPM_BSC/"
)
$script:logQueue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
$script:logTailPosition = 0L
$script:logTailRemainder = ""

function Append-Log {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return
  }

  if (-not [string]::IsNullOrWhiteSpace($script:currentLogFile)) {
    try {
      [System.IO.File]::AppendAllText(
        $script:currentLogFile,
        $Text + [Environment]::NewLine,
        [System.Text.Encoding]::UTF8
      )
    } catch {
      # Si no se puede escribir al archivo en este instante, no bloqueamos la UI.
    }
  }

  Append-LogUIOnly -Text $Text
}

function Append-LogUIOnly {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return
  }

  try {
    if ($null -ne $script:txtLog -and -not $script:txtLog.IsDisposed -and $script:txtLog.IsHandleCreated -and -not $script:txtLog.InvokeRequired) {
      $script:txtLog.AppendText($Text + [Environment]::NewLine)
      $script:txtLog.SelectionStart = $script:txtLog.TextLength
      $script:txtLog.ScrollToCaret()
      return
    }
    [void]$script:logQueue.Enqueue($Text)
  } catch {
    # Nunca propagamos excepciones de logging a la UI.
  }
}

function Flush-LogQueueToUI {
  try {
    if ($null -eq $script:txtLog -or $script:txtLog.IsDisposed) {
      return
    }
    if (-not $script:txtLog.IsHandleCreated) {
      return
    }

    $hadLines = $false
    while ($true) {
      $line = $null
      $hasItem = $false

      [System.Threading.Monitor]::Enter($script:logQueue.SyncRoot)
      try {
        if ($script:logQueue.Count -gt 0) {
          $line = [string]$script:logQueue.Dequeue()
          $hasItem = $true
        }
      } finally {
        [System.Threading.Monitor]::Exit($script:logQueue.SyncRoot)
      }

      if (-not $hasItem) {
        break
      }

      if (-not [string]::IsNullOrEmpty($line)) {
        $script:txtLog.AppendText($line + [Environment]::NewLine)
        $hadLines = $true
      }
    }

    if ($hadLines) {
      $script:txtLog.SelectionStart = $script:txtLog.TextLength
      $script:txtLog.ScrollToCaret()
    }
  } catch {
    # Nunca propagamos excepciones de logging a la UI.
  }
}

function Pump-LogFileToQueue {
  if ([string]::IsNullOrWhiteSpace($script:currentLogFile)) {
    return
  }
  if (-not (Test-Path $script:currentLogFile)) {
    return
  }

  try {
    $fs = [System.IO.File]::Open(
      $script:currentLogFile,
      [System.IO.FileMode]::Open,
      [System.IO.FileAccess]::Read,
      [System.IO.FileShare]::ReadWrite
    )
    try {
      if ($script:logTailPosition -gt $fs.Length) {
        $script:logTailPosition = 0L
        $script:logTailRemainder = ""
      }

      [void]$fs.Seek($script:logTailPosition, [System.IO.SeekOrigin]::Begin)
      $sr = New-Object System.IO.StreamReader(
        $fs,
        [System.Text.Encoding]::UTF8,
        $true,
        4096,
        $true
      )

      $chunk = $sr.ReadToEnd()
      $script:logTailPosition = $fs.Position

      if (-not [string]::IsNullOrEmpty($chunk)) {
        $text = ($script:logTailRemainder + $chunk) -replace "`r`n", "`n" -replace "`r", "`n"
        $parts = $text -split "`n", -1

        if ($text.EndsWith("`n")) {
          $script:logTailRemainder = ""
        } else {
          $script:logTailRemainder = $parts[-1]
          if ($parts.Length -gt 1) {
            $parts = $parts[0..($parts.Length - 2)]
          } else {
            $parts = @()
          }
        }

        foreach ($line in $parts) {
          if (-not [string]::IsNullOrWhiteSpace($line)) {
            [void]$script:logQueue.Enqueue($line)
          }
        }
      }
    } finally {
      $fs.Dispose()
    }
  } catch {
    # Nunca propagamos excepciones de tail de log a la UI.
  }
}

function Safe-UIAction {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock]$Action,
    [string]$Context = "UI"
  )

  try {
    & $Action
  } catch {
    try {
      Append-Log ("[{0}] Error en {1}: {2}" -f (Get-Date -Format "HH:mm:ss"), $Context, $_.Exception.Message)
    } catch {
      # Evitar cascadas de error al reportar errores de UI.
    }
  }
}

function Try-SetBackColor {
  param(
    [Parameter(Mandatory = $true)]
    $Control,
    [AllowNull()]
    [System.Drawing.Color]$Color,
    [System.Drawing.Color]$Fallback = [System.Drawing.Color]::White
  )

  try {
    if ($null -eq $Color -or $Color.IsEmpty) {
      $Control.BackColor = $Fallback
    } else {
      $Control.BackColor = $Color
    }
  } catch {
    # No dejamos que un tema/estado visual cierre la app.
  }
}

function Stop-PlaywrightProcesses {
  param(
    [string]$SpecPath = "",
    [bool]$KillAllEdge = $false
  )

  $projectRoot = $script:projectRoot
  $specToken = [string]$SpecPath
  $processes = Get-CimInstance Win32_Process | Where-Object {
    $cmd = $_.CommandLine
    if ([string]::IsNullOrWhiteSpace($cmd)) { return $false }
    ($cmd -match "playwright\\s+test") -and ($cmd -like "*$projectRoot*") -and (
      [string]::IsNullOrWhiteSpace($specToken) -or ($cmd -like "*$specToken*")
    )
  }

  foreach ($proc in $processes) {
    try {
      Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
      Append-Log ("Proceso detenido (Playwright): {0} (PID={1})" -f $proc.Name, $proc.ProcessId)
    } catch {
      Append-Log ("No se pudo detener proceso (Playwright) PID={0}: {1}" -f $proc.ProcessId, $_.Exception.Message)
    }
  }

  $edgeProcesses = Get-CimInstance Win32_Process | Where-Object {
    $cmd = $_.CommandLine
    if ([string]::IsNullOrWhiteSpace($cmd)) { return $false }
    ($_.Name -match "^msedge(\\.exe)?$") -and (
      $cmd -like "*$projectRoot\\profile-edge-robot*" -or
      $cmd -like "*$projectRoot\\profile-edge-robot-deposito*" -or
      $cmd -like "*$projectRoot\\artifacts\\profiles\\profile-edge-robot*" -or
      $cmd -like "*$projectRoot\\artifacts\\profiles\\profile-edge-robot-deposito*"
    )
  }

  foreach ($proc in $edgeProcesses) {
    try {
      Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
      Append-Log ("Proceso detenido (Edge PW): {0} (PID={1})" -f $proc.Name, $proc.ProcessId)
    } catch {
      Append-Log ("No se pudo detener Edge PW PID={0}: {1}" -f $proc.ProcessId, $_.Exception.Message)
    }
  }

  if ($KillAllEdge) {
    $allEdge = Get-CimInstance Win32_Process | Where-Object {
      $_.Name -match "^msedge(\\.exe)?$"
    }
    foreach ($proc in $allEdge) {
      try {
        Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
        Append-Log ("Proceso detenido (Edge): {0} (PID={1})" -f $proc.Name, $proc.ProcessId)
      } catch {
        Append-Log ("No se pudo detener Edge PID={0}: {1}" -f $proc.ProcessId, $_.Exception.Message)
      }
    }
  }
}

function Clear-ProfileLocks {
  param(
    [string[]]$ProfileRoots
  )

  $lockNames = @("SingletonLock", "SingletonCookie", "SingletonSocket", "Lockfile", "LOCK")
  foreach ($root in $ProfileRoots) {
    if ([string]::IsNullOrWhiteSpace($root)) { continue }
    if (-not (Test-Path $root)) { continue }

    foreach ($name in $lockNames) {
      $paths = @(
        (Join-Path $root $name),
        (Join-Path (Join-Path $root "Default") $name)
      )
      foreach ($p in $paths) {
        if (Test-Path $p) {
          try {
            Remove-Item -LiteralPath $p -Force -ErrorAction Stop
            Append-Log ("Lock eliminado: {0}" -f $p)
          } catch {
            Append-Log ("No se pudo eliminar lock {0}: {1}" -f $p, $_.Exception.Message)
          }
        }
      }
    }
  }
}

function Get-ModeLabel {
  param([string]$Mode)

  switch ($Mode) {
    "existente" { return "Cuenta Efectivo - Cliente existente" }
    "nuevo" { return "Cuenta Efectivo - Cliente nuevo" }
    "ambos" { return "Cuenta Efectivo - Ambos" }
    "deposito-plazos" { return "Deposito a plazos - existente" }
    default { return $Mode }
  }
}

function Get-RunnerCommandText {
  param(
    [Parameter(Mandatory = $true)][string]$Mode,
    [Parameter(Mandatory = $true)][string]$LogPath,
    [string]$BizagiUrl = ""
  )

  $projectEscaped = $script:projectRoot.Replace("'", "''")
  $logEscaped = $LogPath.Replace("'", "''")
  $bizagiUrlEscaped = $BizagiUrl.Replace("'", "''")
  $specExistente = "tests/e2e/ce-e/ce-ex.spec.ts"
  $specNuevo = "tests/e2e/ce-e/ce-new.spec.ts"
  $specDeposito = "tests/e2e/ce-e/Certificado-ex.spec.ts"

  return @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'
`$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

function Write-RunnerLine {
  param([string]`$Text)
  [System.IO.File]::AppendAllText(
    '$logEscaped',
    (`$Text + [Environment]::NewLine),
    [System.Text.Encoding]::UTF8
  )
}

function Invoke-PlaywrightSpec {
  param([Parameter(Mandatory = `$true)][string]`$SpecPath)
  Write-RunnerLine ""
  Write-RunnerLine ("Ejecutando: {0}" -f `$SpecPath)
  Write-RunnerLine ""
  & npx.cmd playwright test `$SpecPath --reporter=list *>> '$logEscaped'
  return `$LASTEXITCODE
}

if ('$bizagiUrlEscaped' -ne '') { `$env:PW_BIZAGI_URL = '$bizagiUrlEscaped' }
Set-Location -LiteralPath '$projectEscaped'
if (-not (Test-Path 'package.json')) {
  throw "No se encontro package.json en '$projectEscaped'."
}
if (-not (Test-Path 'node_modules')) {
  Write-RunnerLine "No se encontro node_modules. Ejecutando npm install..."
  & npm.cmd install *>> '$logEscaped'
  if (`$LASTEXITCODE -ne 0) {
    throw "npm install fallo."
  }
}

switch ('$Mode') {
  'existente' {
    `$code = Invoke-PlaywrightSpec -SpecPath '$specExistente'
    exit `$code
  }
  'nuevo' {
    `$code = Invoke-PlaywrightSpec -SpecPath '$specNuevo'
    exit `$code
  }
  'ambos' {
    `$code1 = Invoke-PlaywrightSpec -SpecPath '$specExistente'
    if (`$code1 -ne 0) { exit `$code1 }
    `$code2 = Invoke-PlaywrightSpec -SpecPath '$specNuevo'
    exit `$code2
  }
  'deposito-plazos' {
    `$code = Invoke-PlaywrightSpec -SpecPath '$specDeposito'
    exit `$code
  }
  default {
    throw "Modo no soportado: $Mode"
  }
}
"@
}

function Set-RunState {
  param(
    [bool]$IsRunning,
    [string]$Mode = ""
  )

  $action = {
    param([bool]$running, [string]$mode)
    $script:btnExistente.Enabled = -not $running
    $script:btnNuevo.Enabled = -not $running
    $script:btnAmbos.Enabled = -not $running
    $script:btnDeposito.Enabled = -not $running
    $script:btnStop.Enabled = $running

    if ($running) {
      $script:lblStatus.Text = "Estado: Ejecutando (" + (Get-ModeLabel -Mode $mode) + ")"
      $script:lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    } else {
      $script:lblStatus.Text = "Estado: Listo"
      $script:lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    }
  }

  if ($script:form.InvokeRequired) {
    $null = $script:form.BeginInvoke($action, @($IsRunning, $Mode))
  } else {
    & $action $IsRunning $Mode
  }
}

function Start-Run {
  param([string]$Mode)

  if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
    Append-Log "Ya hay una ejecucion en progreso."
    return
  }

  $logsDir = Join-Path $script:projectRoot "artifacts\outputs\runner-logs"
  if (-not (Test-Path $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
  }

  $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $script:currentLogFile = Join-Path $logsDir ("runner_{0}_{1}.log" -f $Mode, $stamp)
  $script:currentMode = $Mode
  [System.IO.File]::WriteAllText($script:currentLogFile, "", [System.Text.Encoding]::UTF8)
  $script:logTailPosition = 0L
  $script:logTailRemainder = ""

  $selectedBizagiUrl = $script:bizagiUrls[0]
  if ($null -ne $script:cmbBizagi -and $script:cmbBizagi.SelectedIndex -ge 0 -and $script:cmbBizagi.SelectedIndex -lt $script:bizagiUrls.Count) {
    $selectedBizagiUrl = $script:bizagiUrls[$script:cmbBizagi.SelectedIndex]
  }

  Append-Log ""
  Append-Log ("[{0}] Inicio de ejecucion: {1}" -f (Get-Date -Format "HH:mm:ss"), (Get-ModeLabel -Mode $Mode))
  Append-Log ("Bizagi URL: {0}" -f $selectedBizagiUrl)
  Append-Log ("Log: {0}" -f $script:currentLogFile)

  $specPath = ""
  $profileRoots = @()
  switch ($Mode) {
    "existente" { $specPath = "tests/e2e/ce-e/ce-ex.spec.ts" }
    "nuevo" { $specPath = "tests/e2e/ce-e/ce-new.spec.ts" }
    "ambos" { $specPath = "" }
    "deposito-plazos" { $specPath = "tests/e2e/ce-e/Certificado-ex.spec.ts" }
  }
  Append-Log "Limpiando procesos previos de Playwright..."
  Stop-PlaywrightProcesses -SpecPath $specPath -KillAllEdge $true
  Start-Sleep -Milliseconds 800
  $profileRoots = @(
    (Join-Path $script:projectRoot "artifacts\profiles\profile-edge-robot"),
    (Join-Path $script:projectRoot "artifacts\profiles\profile-edge-robot-deposito"),
    (Join-Path $script:projectRoot "profile-edge-robot"),
    (Join-Path $script:projectRoot "profile-edge-robot-deposito")
  )
  Append-Log "Limpiando locks de perfil..."
  Clear-ProfileLocks -ProfileRoots $profileRoots
  try {
    $script:logTailPosition = ([System.IO.FileInfo]$script:currentLogFile).Length
  } catch {
    $script:logTailPosition = 0L
  }

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path $psExe)) {
      $psExe = "powershell.exe"
    }
    $psi.FileName = $psExe
    $commandText = Get-RunnerCommandText -Mode $Mode -LogPath $script:currentLogFile -BizagiUrl $selectedBizagiUrl
    $commandBytes = [System.Text.Encoding]::Unicode.GetBytes($commandText)
    $encodedCommand = [Convert]::ToBase64String($commandBytes)
    $psi.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
    $psi.WorkingDirectory = $script:projectRoot
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true

    $started = $proc.Start()
    if (-not $started) {
      throw "No se pudo iniciar powershell.exe para el runner."
    }

    $script:outputDataHandler = $null
    $script:errorDataHandler = $null
    $script:currentProcess = $proc
  } catch {
    Append-Log ("No se pudo iniciar el proceso: {0}" -f $_.Exception.Message)
    $script:currentProcess = $null
    $script:outputDataHandler = $null
    $script:errorDataHandler = $null
    return
  }

  Set-RunState -IsRunning $true -Mode $Mode
  $script:pollTimer.Start()
}

function Stop-Run {
  if (-not $script:currentProcess) {
    return
  }

  if (-not $script:currentProcess.HasExited) {
    try {
      Stop-Process -Id $script:currentProcess.Id -Force -ErrorAction Stop
      Append-Log "Ejecucion detenida manualmente."
    } catch {
      Append-Log ("No se pudo detener el proceso: {0}" -f $_.Exception.Message)
    }
  }

  try { $script:currentProcess.CancelOutputRead() } catch { }
  try { $script:currentProcess.CancelErrorRead() } catch { }
  if ($script:outputDataHandler) {
    try { $script:currentProcess.remove_OutputDataReceived($script:outputDataHandler) } catch { }
  }
  if ($script:errorDataHandler) {
    try { $script:currentProcess.remove_ErrorDataReceived($script:errorDataHandler) } catch { }
  }
  $script:outputDataHandler = $null
  $script:errorDataHandler = $null
  $script:currentProcess = $null
  Set-RunState -IsRunning $false
}

$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "PortalComercial - Runner"
$script:form.StartPosition = "CenterScreen"
$script:form.Size = New-Object System.Drawing.Size(1120, 760)
$script:form.MinimumSize = New-Object System.Drawing.Size(1024, 680)
$script:form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F6FA")

$colorPortalBlue = [System.Drawing.ColorTranslator]::FromHtml("#2D8ED4")
$colorPortalBlueDark = [System.Drawing.ColorTranslator]::FromHtml("#1F6FA8")
$colorPortalBlueSoft = [System.Drawing.ColorTranslator]::FromHtml("#DCEFFD")
$colorPortalGreen = [System.Drawing.ColorTranslator]::FromHtml("#14A86B")
$colorTextMuted = [System.Drawing.ColorTranslator]::FromHtml("#4D5B69")
$colorTextStrong = [System.Drawing.ColorTranslator]::FromHtml("#183B56")
$colorCard = [System.Drawing.Color]::White

function Set-PrimaryButtonStyle {
  param(
    [System.Windows.Forms.Button]$Button,
    [System.Drawing.Color]$BaseColor,
    [System.Drawing.Color]$HoverColor
  )

  $base = $BaseColor
  $hover = $HoverColor

  $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Button.FlatAppearance.BorderSize = 0
  $Button.BackColor = $base
  $Button.ForeColor = [System.Drawing.Color]::White
  $Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
  $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
  $Button.Add_MouseEnter(({ param($s, $e) Try-SetBackColor -Control $s -Color $hover -Fallback $base }.GetNewClosure()))
  $Button.Add_MouseLeave(({ param($s, $e) Try-SetBackColor -Control $s -Color $base }.GetNewClosure()))
}

function Set-SecondaryButtonStyle {
  param([System.Windows.Forms.Button]$Button)

  $secondaryBase = [System.Drawing.Color]::White
  $secondaryHover = [System.Drawing.ColorTranslator]::FromHtml("#F5FAFF")
  $secondaryBorder = [System.Drawing.ColorTranslator]::FromHtml("#C8D7E5")

  $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Button.FlatAppearance.BorderSize = 1
  $Button.FlatAppearance.BorderColor = $secondaryBorder
  $Button.BackColor = $secondaryBase
  $Button.ForeColor = $colorTextStrong
  $Button.Font = New-Object System.Drawing.Font("Segoe UI", 9)
  $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
  $Button.Add_MouseEnter(({ param($s, $e) Try-SetBackColor -Control $s -Color $secondaryHover -Fallback $secondaryBase }.GetNewClosure()))
  $Button.Add_MouseLeave(({ param($s, $e) Try-SetBackColor -Control $s -Color $secondaryBase }.GetNewClosure()))
}

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.Height = 484
$topPanel.BackColor = $script:form.BackColor
$script:form.Controls.Add($topPanel)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(2600, 76)
$headerPanel.BackColor = $colorPortalBlue
$topPanel.Controls.Add($headerPanel)

$title = New-Object System.Windows.Forms.Label
$title.Text = "Motor de Ejecucion - PortalComercial"
$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 17, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.ForeColor = [System.Drawing.Color]::White
$title.Location = New-Object System.Drawing.Point(20, 21)
$headerPanel.Controls.Add($title)

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Location = New-Object System.Drawing.Point(18, 86)
$statusPanel.Size = New-Object System.Drawing.Size(380, 28)
$statusPanel.BackColor = $colorPortalBlueSoft
$topPanel.Controls.Add($statusPanel)

$statusDot = New-Object System.Windows.Forms.Panel
$statusDot.Size = New-Object System.Drawing.Size(10, 10)
$statusDot.Location = New-Object System.Drawing.Point(10, 9)
$statusDot.BackColor = $colorPortalGreen
$statusPanel.Controls.Add($statusDot)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Text = "Estado: Listo"
$script:lblStatus.AutoSize = $true
$script:lblStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#0D6E47")
$script:lblStatus.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5, [System.Drawing.FontStyle]::Bold)
$script:lblStatus.Location = New-Object System.Drawing.Point(26, 6)
$statusPanel.Controls.Add($script:lblStatus)

$groupCuenta = New-Object System.Windows.Forms.GroupBox
$groupCuenta.Text = "Cuenta Efectivo"
$groupCuenta.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5, [System.Drawing.FontStyle]::Bold)
$groupCuenta.ForeColor = $colorTextStrong
$groupCuenta.BackColor = $colorCard
$groupCuenta.Location = New-Object System.Drawing.Point(18, 122)
$groupCuenta.Size = New-Object System.Drawing.Size(1068, 106)
$topPanel.Controls.Add($groupCuenta)

$script:btnExistente = New-Object System.Windows.Forms.Button
$script:btnExistente.Text = "Cuenta Efectivo - Cliente existente"
$script:btnExistente.Size = New-Object System.Drawing.Size(338, 44)
$script:btnExistente.Location = New-Object System.Drawing.Point(14, 34)
Set-PrimaryButtonStyle -Button $script:btnExistente -BaseColor $colorPortalBlue -HoverColor $colorPortalBlueDark
$groupCuenta.Controls.Add($script:btnExistente)

$script:btnNuevo = New-Object System.Windows.Forms.Button
$script:btnNuevo.Text = "Cuenta Efectivo - Cliente nuevo"
$script:btnNuevo.Size = New-Object System.Drawing.Size(338, 44)
$script:btnNuevo.Location = New-Object System.Drawing.Point(364, 34)
Set-PrimaryButtonStyle -Button $script:btnNuevo -BaseColor $colorPortalBlue -HoverColor $colorPortalBlueDark
$groupCuenta.Controls.Add($script:btnNuevo)

$script:btnAmbos = New-Object System.Windows.Forms.Button
$script:btnAmbos.Text = "Cuenta Efectivo - Ambos"
$script:btnAmbos.Size = New-Object System.Drawing.Size(338, 44)
$script:btnAmbos.Location = New-Object System.Drawing.Point(714, 34)
Set-PrimaryButtonStyle -Button $script:btnAmbos -BaseColor $colorPortalBlue -HoverColor $colorPortalBlueDark
$groupCuenta.Controls.Add($script:btnAmbos)

$groupDeposito = New-Object System.Windows.Forms.GroupBox
$groupDeposito.Text = "Deposito a plazos"
$groupDeposito.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5, [System.Drawing.FontStyle]::Bold)
$groupDeposito.ForeColor = $colorTextStrong
$groupDeposito.BackColor = $colorCard
$groupDeposito.Location = New-Object System.Drawing.Point(18, 234)
$groupDeposito.Size = New-Object System.Drawing.Size(1068, 106)
$topPanel.Controls.Add($groupDeposito)

$script:btnDeposito = New-Object System.Windows.Forms.Button
$script:btnDeposito.Text = "Deposito a plazos - existente"
$script:btnDeposito.Size = New-Object System.Drawing.Size(338, 44)
$script:btnDeposito.Location = New-Object System.Drawing.Point(14, 34)
Set-PrimaryButtonStyle -Button $script:btnDeposito -BaseColor $colorPortalBlue -HoverColor $colorPortalBlueDark
$groupDeposito.Controls.Add($script:btnDeposito)

$groupBizagi = New-Object System.Windows.Forms.GroupBox
$groupBizagi.Text = "Configuracion Bizagi"
$groupBizagi.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5, [System.Drawing.FontStyle]::Bold)
$groupBizagi.ForeColor = $colorTextStrong
$groupBizagi.BackColor = $colorCard
$groupBizagi.Location = New-Object System.Drawing.Point(18, 346)
$groupBizagi.Size = New-Object System.Drawing.Size(1068, 62)
$topPanel.Controls.Add($groupBizagi)

$lblBizagi = New-Object System.Windows.Forms.Label
$lblBizagi.Text = "URL Bizagi:"
$lblBizagi.AutoSize = $true
$lblBizagi.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$lblBizagi.ForeColor = $colorTextStrong
$lblBizagi.Location = New-Object System.Drawing.Point(14, 24)
$groupBizagi.Controls.Add($lblBizagi)

$script:cmbBizagi = New-Object System.Windows.Forms.ComboBox
$script:cmbBizagi.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:cmbBizagi.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$script:cmbBizagi.Location = New-Object System.Drawing.Point(110, 20)
$script:cmbBizagi.Size = New-Object System.Drawing.Size(720, 26)
[void]$script:cmbBizagi.Items.Add("Dev Cloud  -  https://dev-bscrd-santacruz.bizagi.com/")
[void]$script:cmbBizagi.Items.Add("Dev Local  -  https://srvdevbizapp01/BPM_BSC/")
$script:cmbBizagi.SelectedIndex = 0
$groupBizagi.Controls.Add($script:cmbBizagi)

$groupActions = New-Object System.Windows.Forms.GroupBox
$groupActions.Text = "Acciones de control"
$groupActions.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$groupActions.ForeColor = $colorTextStrong
$groupActions.BackColor = $colorCard
$groupActions.Location = New-Object System.Drawing.Point(18, 416)
$groupActions.Size = New-Object System.Drawing.Size(560, 56)
$topPanel.Controls.Add($groupActions)

$script:btnStop = New-Object System.Windows.Forms.Button
$script:btnStop.Text = "Detener"
$script:btnStop.Size = New-Object System.Drawing.Size(110, 30)
$script:btnStop.Location = New-Object System.Drawing.Point(12, 20)
Set-SecondaryButtonStyle -Button $script:btnStop
$script:btnStop.Enabled = $false
$groupActions.Controls.Add($script:btnStop)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Limpiar log"
$btnClear.Size = New-Object System.Drawing.Size(120, 30)
$btnClear.Location = New-Object System.Drawing.Point(132, 20)
Set-SecondaryButtonStyle -Button $btnClear
$groupActions.Controls.Add($btnClear)

$btnOpenLog = New-Object System.Windows.Forms.Button
$btnOpenLog.Text = "Abrir ultimo log"
$btnOpenLog.Size = New-Object System.Drawing.Size(140, 30)
$btnOpenLog.Location = New-Object System.Drawing.Point(262, 20)
Set-SecondaryButtonStyle -Button $btnOpenLog
$groupActions.Controls.Add($btnOpenLog)

$logContainer = New-Object System.Windows.Forms.Panel
$logContainer.Dock = "Fill"
$logContainer.Padding = New-Object System.Windows.Forms.Padding(16, 10, 16, 16)
$logContainer.BackColor = $script:form.BackColor
$script:form.Controls.Add($logContainer)
# Para Dock, el orden importa: Fill debe quedar por detras de Top para ocupar solo el espacio restante.
$script:form.Controls.SetChildIndex($logContainer, 0)

$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Consola de ejecucion"
$logGroup.Dock = "Fill"
$logGroup.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5, [System.Drawing.FontStyle]::Bold)
$logGroup.ForeColor = $colorTextStrong
$logGroup.BackColor = $colorCard
$logContainer.Controls.Add($logGroup)

$script:txtLog = New-Object System.Windows.Forms.TextBox
$script:txtLog.Dock = "Fill"
$script:txtLog.Multiline = $true
$script:txtLog.ReadOnly = $true
$script:txtLog.ScrollBars = "Both"
$script:txtLog.WordWrap = $false
$script:txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$script:txtLog.BackColor = [System.Drawing.Color]::White
$script:txtLog.ForeColor = [System.Drawing.Color]::Black
$script:txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logGroup.Controls.Add($script:txtLog)

$script:pollTimer = New-Object System.Windows.Forms.Timer
$script:pollTimer.Interval = 120
$script:pollTimer.Add_Tick({
  Safe-UIAction -Context "pollTimer.Tick" -Action {
    Pump-LogFileToQueue
    Flush-LogQueueToUI

    if (-not $script:currentProcess) {
      return
    }

    if ($script:currentProcess.HasExited) {
      try { [void]$script:currentProcess.WaitForExit(2000) } catch { }
      try { $script:currentProcess.CancelOutputRead() } catch { }
      try { $script:currentProcess.CancelErrorRead() } catch { }
      if ($script:outputDataHandler) {
        try { $script:currentProcess.remove_OutputDataReceived($script:outputDataHandler) } catch { }
      }
      if ($script:errorDataHandler) {
        try { $script:currentProcess.remove_ErrorDataReceived($script:errorDataHandler) } catch { }
      }
      Pump-LogFileToQueue
      Flush-LogQueueToUI
      $script:outputDataHandler = $null
      $script:errorDataHandler = $null

      $exitCode = $script:currentProcess.ExitCode
      if ($exitCode -eq 0) {
        Append-Log ("[{0}] Ejecucion finalizada OK." -f (Get-Date -Format "HH:mm:ss"))
      } else {
        Append-Log ("[{0}] Ejecucion finalizo con errores (exit={1})." -f (Get-Date -Format "HH:mm:ss"), $exitCode)
      }
      $script:currentProcess = $null
      Set-RunState -IsRunning $false
    }
  }
})

$script:btnExistente.Add_Click({ Safe-UIAction -Context "btnExistente.Click" -Action { Start-Run -Mode "existente" } })
$script:btnNuevo.Add_Click({ Safe-UIAction -Context "btnNuevo.Click" -Action { Start-Run -Mode "nuevo" } })
$script:btnAmbos.Add_Click({ Safe-UIAction -Context "btnAmbos.Click" -Action { Start-Run -Mode "ambos" } })
$script:btnDeposito.Add_Click({ Safe-UIAction -Context "btnDeposito.Click" -Action { Start-Run -Mode "deposito-plazos" } })
$script:btnStop.Add_Click({ Safe-UIAction -Context "btnStop.Click" -Action { Stop-Run } })

$btnClear.Add_Click({
  Safe-UIAction -Context "btnClear.Click" -Action {
    $script:txtLog.Clear()
    Append-Log "Log limpiado."
  }
})

$btnOpenLog.Add_Click({
  Safe-UIAction -Context "btnOpenLog.Click" -Action {
    if (-not [string]::IsNullOrWhiteSpace($script:currentLogFile) -and (Test-Path $script:currentLogFile)) {
      Start-Process notepad.exe $script:currentLogFile | Out-Null
    } else {
      [System.Windows.Forms.MessageBox]::Show(
        "Aun no existe un log generado.",
        "PortalComercial Runner",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
      ) | Out-Null
    }
  }
})

$script:form.Add_FormClosing({
  Safe-UIAction -Context "form.Closing" -Action { Stop-Run }
})
$script:form.Add_Shown({
  Safe-UIAction -Context "form.Shown" -Action { Flush-LogQueueToUI }
})

[System.Windows.Forms.Application]::add_ThreadException({
  param($sender, $eventArgs)
  Safe-UIAction -Context "ThreadException" -Action {
    $msg = if ($eventArgs -and $eventArgs.Exception) { $eventArgs.Exception.Message } else { "Excepcion no controlada en UI." }
    Append-Log ("[{0}] ERROR UI: {1}" -f (Get-Date -Format "HH:mm:ss"), $msg)
  }
})
[System.AppDomain]::CurrentDomain.add_UnhandledException({
  param($sender, $eventArgs)
  Safe-UIAction -Context "UnhandledException" -Action {
    $msg = if ($eventArgs -and $eventArgs.ExceptionObject) { [string]$eventArgs.ExceptionObject } else { "Excepcion no controlada en AppDomain." }
    Append-Log ("[{0}] ERROR FATAL: {1}" -f (Get-Date -Format "HH:mm:ss"), $msg)
  }
})

Append-Log ("Proyecto: {0}" -f $script:projectRoot)
Append-Log "Listo para ejecutar."

try {
  $script:pollTimer.Start()
  [System.Windows.Forms.Application]::Run($script:form)
} finally {
  $script:pollTimer.Stop()
  Stop-Run
}
