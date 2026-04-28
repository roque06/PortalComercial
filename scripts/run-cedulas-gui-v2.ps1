$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try {
  [System.Windows.Forms.Application]::SetUnhandledExceptionMode(
    [System.Windows.Forms.UnhandledExceptionMode]::CatchException
  )
} catch {
}

$script:scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:projectRoot = Split-Path -Parent $script:scriptDir
$script:dbConnectionsPath = "C:\CierreBancarioQA\src\config\dbConnections.json"
$script:sqlplusPath = "C:\instantclient_21_13\sqlplus.exe"
$script:currentProcess = $null
$script:currentLogFile = ""
$script:logQueue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
$script:logTailPosition = 0L
$script:logTailRemainder = ""

function Add-Log {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return }
  $script:txtLog.AppendText($Text + [Environment]::NewLine)
  $script:txtLog.SelectionStart = $script:txtLog.TextLength
  $script:txtLog.ScrollToCaret()
}

function Enqueue-Log {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return }
  [void]$script:logQueue.Enqueue($Text)
}

function Flush-LogQueue {
  if ($null -eq $script:txtLog -or $script:txtLog.IsDisposed) {
    return
  }
  while ($script:logQueue.Count -gt 0) {
    $line = $script:logQueue.Dequeue()
    if (-not [string]::IsNullOrWhiteSpace($line)) {
      Add-Log $line
    }
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
      $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 4096, $true)
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
            Enqueue-Log $line
          }
        }
      }
    } finally {
      $fs.Dispose()
    }
  } catch {
  }
}

function Invoke-SafeUi {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock]$Action,
    [string]$Context = "UI"
  )

  try {
    & $Action
  } catch {
    try {
      Add-Log ("ERROR UI [{0}]: {1}" -f $Context, $_.Exception.Message)
    } catch {
    }
  }
}

function Set-UiRunningState {
  param([bool]$IsRunning)
  $script:btnBrowseExcel.Enabled = -not $IsRunning
  $script:btnBrowseOutput.Enabled = -not $IsRunning
  $script:btnRun.Enabled = -not $IsRunning
  $script:btnStop.Enabled = $IsRunning
  $script:cmbInstance.Enabled = -not $IsRunning
  $script:txtExcel.Enabled = -not $IsRunning
  $script:txtOutput.Enabled = -not $IsRunning
  $script:lblStatus.Text = if ($IsRunning) { "Estado: Ejecutando" } else { "Estado: Listo" }
  $script:lblStatus.ForeColor = if ($IsRunning) {
    [System.Drawing.ColorTranslator]::FromHtml("#9A6700")
  } else {
    [System.Drawing.ColorTranslator]::FromHtml("#0D6E47")
  }
}

function Load-Instances {
  if (-not (Test-Path $script:dbConnectionsPath)) {
    throw "No existe dbConnections.json en $($script:dbConnectionsPath)"
  }

  $connections = Get-Content $script:dbConnectionsPath -Raw | ConvertFrom-Json
  $instances = @($connections | ForEach-Object { $_.name } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  return $instances | Sort-Object
}

function Build-DefaultOutputPath {
  param([string]$ExcelPath, [string]$Instance)
  if ([string]::IsNullOrWhiteSpace($ExcelPath)) { return "" }
  try {
    $directory = [System.IO.Path]::GetDirectoryName($ExcelPath)
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath)
    $extension = [System.IO.Path]::GetExtension($ExcelPath)
    if ([string]::IsNullOrWhiteSpace($extension)) { $extension = ".xlsx" }
    return [System.IO.Path]::Combine($directory, ("{0}_no_encontradas_{1}{2}" -f $fileName, $Instance.ToUpper(), $extension))
  } catch {
    return ""
  }
}

function Start-Validation {
  try {
    $excelPath = $script:txtExcel.Text.Trim()
    $outputPath = $script:txtOutput.Text.Trim()
    $instance = [string]$script:cmbInstance.SelectedItem

    if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path $excelPath)) {
      [System.Windows.Forms.MessageBox]::Show(
        "Selecciona un archivo Excel valido.",
        "Validador de Cedulas",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      ) | Out-Null
      return
    }

    if ([string]::IsNullOrWhiteSpace($instance)) {
      [System.Windows.Forms.MessageBox]::Show(
        "Selecciona una instancia.",
        "Validador de Cedulas",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      ) | Out-Null
      return
    }

    if ([string]::IsNullOrWhiteSpace($outputPath)) {
      $outputPath = Build-DefaultOutputPath -ExcelPath $excelPath -Instance $instance
      $script:txtOutput.Text = $outputPath
    }

    $logsDir = Join-Path $script:projectRoot "artifacts\outputs\cedulas-logs"
    if (-not (Test-Path $logsDir)) {
      New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:currentLogFile = Join-Path $logsDir ("cedulas_{0}_{1}.log" -f $instance.ToUpper(), $stamp)
    [System.IO.File]::WriteAllText($script:currentLogFile, "", [System.Text.Encoding]::UTF8)
    $script:logTailPosition = 0L
    $script:logTailRemainder = ""

    $scriptPath = Join-Path $script:projectRoot "scripts\check-cedulas-db.mjs"
    $projectEscaped = $script:projectRoot.Replace("'", "''")
    $scriptEscaped = $scriptPath.Replace("'", "''")
    $excelEscaped = $excelPath.Replace("'", "''")
    $instanceEscaped = $instance.Replace("'", "''")
    $outputEscaped = $outputPath.Replace("'", "''")
    $dbEscaped = $script:dbConnectionsPath.Replace("'", "''")
    $sqlplusEscaped = $script:sqlplusPath.Replace("'", "''")
    $logEscaped = $script:currentLogFile.Replace("'", "''")

    $commandText = @"
`$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath '$projectEscaped'
& node.exe '$scriptEscaped' --input '$excelEscaped' --instance '$instanceEscaped' --output '$outputEscaped' --db-connections '$dbEscaped' --sqlplus '$sqlplusEscaped' *>> '$logEscaped'
exit `$LASTEXITCODE
"@

    $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($commandText))
    $psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path $psExe)) {
      $psExe = "powershell.exe"
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $psExe
    $psi.WorkingDirectory = $script:projectRoot
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true

    $started = $proc.Start()
    if (-not $started) {
      throw "No se pudo iniciar el proceso de validacion."
    }

    $script:currentProcess = $proc
    Add-Log ""
    Add-Log ("Archivo origen: " + $excelPath)
    Add-Log ("Instancia: " + $instance.ToUpper())
    Add-Log ("Archivo salida: " + $outputPath)
    Add-Log ("Log: " + $script:currentLogFile)
    Set-UiRunningState -IsRunning $true
    $script:pollTimer.Start()
  } catch {
    Add-Log ("ERROR al iniciar la validacion: " + $_.Exception.Message)
    Set-UiRunningState -IsRunning $false
  }
}

function Stop-Validation {
  if ($null -eq $script:currentProcess) { return }
  if (-not $script:currentProcess.HasExited) {
    try {
      Stop-Process -Id $script:currentProcess.Id -Force -ErrorAction Stop
      Add-Log "Proceso detenido manualmente."
    } catch {
      Add-Log ("No se pudo detener el proceso: " + $_.Exception.Message)
    }
  }
  $script:currentProcess = $null
  $script:pollTimer.Stop()
  Set-UiRunningState -IsRunning $false
}

$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "Validador de Cedulas B2000"
$script:form.StartPosition = "CenterScreen"
$script:form.Size = New-Object System.Drawing.Size(980, 700)
$script:form.MinimumSize = New-Object System.Drawing.Size(920, 620)
$script:form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F5F7FB")

$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 78
$header.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1261A0")
$script:form.Controls.Add($header)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Validador de Cedulas por Instancia Oracle"
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20, 22)
$header.Controls.Add($lblTitle)

$card = New-Object System.Windows.Forms.Panel
$card.Location = New-Object System.Drawing.Point(18, 96)
$card.Size = New-Object System.Drawing.Size(928, 220)
$card.BackColor = [System.Drawing.Color]::White
$script:form.Controls.Add($card)

$lblStatusTag = New-Object System.Windows.Forms.Label
$lblStatusTag.Text = "Estado actual"
$lblStatusTag.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblStatusTag.AutoSize = $true
$lblStatusTag.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#52606D")
$lblStatusTag.Location = New-Object System.Drawing.Point(22, 18)
$card.Controls.Add($lblStatusTag)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Text = "Estado: Listo"
$script:lblStatus.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$script:lblStatus.AutoSize = $true
$script:lblStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#0D6E47")
$script:lblStatus.Location = New-Object System.Drawing.Point(20, 40)
$card.Controls.Add($script:lblStatus)

$lblExcel = New-Object System.Windows.Forms.Label
$lblExcel.Text = "Excel de entrada"
$lblExcel.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$lblExcel.AutoSize = $true
$lblExcel.Location = New-Object System.Drawing.Point(22, 78)
$card.Controls.Add($lblExcel)

$script:txtExcel = New-Object System.Windows.Forms.TextBox
$script:txtExcel.Location = New-Object System.Drawing.Point(24, 100)
$script:txtExcel.Size = New-Object System.Drawing.Size(720, 25)
$script:txtExcel.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$card.Controls.Add($script:txtExcel)

$script:btnBrowseExcel = New-Object System.Windows.Forms.Button
$script:btnBrowseExcel.Text = "Buscar..."
$script:btnBrowseExcel.Location = New-Object System.Drawing.Point(758, 98)
$script:btnBrowseExcel.Size = New-Object System.Drawing.Size(130, 30)
$card.Controls.Add($script:btnBrowseExcel)

$lblInstance = New-Object System.Windows.Forms.Label
$lblInstance.Text = "Instancia"
$lblInstance.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$lblInstance.AutoSize = $true
$lblInstance.Location = New-Object System.Drawing.Point(22, 138)
$card.Controls.Add($lblInstance)

$script:cmbInstance = New-Object System.Windows.Forms.ComboBox
$script:cmbInstance.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:cmbInstance.Location = New-Object System.Drawing.Point(24, 160)
$script:cmbInstance.Size = New-Object System.Drawing.Size(220, 26)
$script:cmbInstance.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$card.Controls.Add($script:cmbInstance)

$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Excel de salida"
$lblOutput.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$lblOutput.AutoSize = $true
$lblOutput.Location = New-Object System.Drawing.Point(264, 138)
$card.Controls.Add($lblOutput)

$script:txtOutput = New-Object System.Windows.Forms.TextBox
$script:txtOutput.Location = New-Object System.Drawing.Point(266, 160)
$script:txtOutput.Size = New-Object System.Drawing.Size(478, 25)
$script:txtOutput.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$card.Controls.Add($script:txtOutput)

$script:btnBrowseOutput = New-Object System.Windows.Forms.Button
$script:btnBrowseOutput.Text = "Guardar como..."
$script:btnBrowseOutput.Location = New-Object System.Drawing.Point(758, 158)
$script:btnBrowseOutput.Size = New-Object System.Drawing.Size(130, 30)
$card.Controls.Add($script:btnBrowseOutput)

$script:btnRun = New-Object System.Windows.Forms.Button
$script:btnRun.Text = "Validar y Exportar"
$script:btnRun.Location = New-Object System.Drawing.Point(24, 194)
$script:btnRun.Size = New-Object System.Drawing.Size(180, 32)
$script:btnRun.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1261A0")
$script:btnRun.ForeColor = [System.Drawing.Color]::White
$script:btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$script:btnRun.FlatAppearance.BorderSize = 0
$card.Controls.Add($script:btnRun)

$script:btnStop = New-Object System.Windows.Forms.Button
$script:btnStop.Text = "Detener"
$script:btnStop.Location = New-Object System.Drawing.Point(216, 194)
$script:btnStop.Size = New-Object System.Drawing.Size(110, 32)
$script:btnStop.Enabled = $false
$card.Controls.Add($script:btnStop)

$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Proceso"
$logGroup.Location = New-Object System.Drawing.Point(18, 330)
$logGroup.Size = New-Object System.Drawing.Size(928, 312)
$logGroup.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$logGroup.BackColor = [System.Drawing.Color]::White
$script:form.Controls.Add($logGroup)

$script:txtLog = New-Object System.Windows.Forms.TextBox
$script:txtLog.Multiline = $true
$script:txtLog.ScrollBars = "Vertical"
$script:txtLog.ReadOnly = $true
$script:txtLog.WordWrap = $true
$script:txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$script:txtLog.Location = New-Object System.Drawing.Point(16, 28)
$script:txtLog.Size = New-Object System.Drawing.Size(892, 268)
$logGroup.Controls.Add($script:txtLog)

$script:pollTimer = New-Object System.Windows.Forms.Timer
$script:pollTimer.Interval = 150
$script:pollTimer.Add_Tick({
  Invoke-SafeUi -Context "pollTimer.Tick" -Action {
    Pump-LogFileToQueue
    Flush-LogQueue

    if ($null -eq $script:currentProcess) {
      return
    }

    if ($script:currentProcess.HasExited) {
      try { [void]$script:currentProcess.WaitForExit(2000) } catch { }
      Pump-LogFileToQueue
      Flush-LogQueue
      $exitCode = $script:currentProcess.ExitCode
      if ($exitCode -eq 0) {
        Add-Log "Proceso finalizado correctamente."
      } else {
        Add-Log ("Proceso finalizado con error. Exit code: " + $exitCode)
      }
      $script:currentProcess = $null
      $script:pollTimer.Stop()
      Set-UiRunningState -IsRunning $false
    }
  }
})

$script:btnBrowseExcel.Add_Click({
  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|Todos (*.*)|*.*"
  if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $script:txtExcel.Text = $dialog.FileName
    if ([string]::IsNullOrWhiteSpace($script:txtOutput.Text) -and $script:cmbInstance.SelectedItem) {
      $script:txtOutput.Text = Build-DefaultOutputPath -ExcelPath $dialog.FileName -Instance ([string]$script:cmbInstance.SelectedItem)
    }
  }
})

$script:btnBrowseOutput.Add_Click({
  $dialog = New-Object System.Windows.Forms.SaveFileDialog
  $dialog.Filter = "Excel (*.xlsx)|*.xlsx"
  $dialog.DefaultExt = "xlsx"
  $dialog.FileName = if (-not [string]::IsNullOrWhiteSpace($script:txtOutput.Text)) {
    [System.IO.Path]::GetFileName($script:txtOutput.Text)
  } else {
    "resultado_no_encontradas.xlsx"
  }
  if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $script:txtOutput.Text = $dialog.FileName
  }
})

$script:cmbInstance.Add_SelectedIndexChanged({
  if (-not [string]::IsNullOrWhiteSpace($script:txtExcel.Text)) {
    $script:txtOutput.Text = Build-DefaultOutputPath -ExcelPath $script:txtExcel.Text -Instance ([string]$script:cmbInstance.SelectedItem)
  }
})

$script:btnRun.Add_Click({ Invoke-SafeUi -Context "btnRun.Click" -Action { Start-Validation } })
$script:btnStop.Add_Click({ Invoke-SafeUi -Context "btnStop.Click" -Action { Stop-Validation } })
$script:form.Add_FormClosing({ Invoke-SafeUi -Context "form.Closing" -Action { Stop-Validation } })

[System.Windows.Forms.Application]::add_ThreadException({
  param($sender, $eventArgs)
  Invoke-SafeUi -Context "ThreadException" -Action {
    $msg = if ($eventArgs -and $eventArgs.Exception) { $eventArgs.Exception.Message } else { "Excepcion no controlada en UI." }
    Add-Log ("ERROR UI: " + $msg)
  }
})

[System.AppDomain]::CurrentDomain.add_UnhandledException({
  param($sender, $eventArgs)
  Invoke-SafeUi -Context "UnhandledException" -Action {
    $msg = if ($eventArgs -and $eventArgs.ExceptionObject) { [string]$eventArgs.ExceptionObject } else { "Excepcion no controlada en AppDomain." }
    Add-Log ("ERROR FATAL: " + $msg)
  }
})

try {
  $instances = Load-Instances
  foreach ($instance in $instances) {
    [void]$script:cmbInstance.Items.Add($instance)
  }
  $defaultIndex = $script:cmbInstance.Items.IndexOf("QA1")
  if ($defaultIndex -ge 0) {
    $script:cmbInstance.SelectedIndex = $defaultIndex
  } elseif ($script:cmbInstance.Items.Count -gt 0) {
    $script:cmbInstance.SelectedIndex = 0
  }
  Add-Log ("dbConnections.json: " + $script:dbConnectionsPath)
  Add-Log ("sqlplus: " + $script:sqlplusPath)
  Add-Log "Listo para procesar archivos."
} catch {
  [System.Windows.Forms.MessageBox]::Show(
    $_.Exception.Message,
    "Validador de Cedulas",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  ) | Out-Null
}

[System.Windows.Forms.Application]::Run($script:form)
