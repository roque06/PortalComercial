param(
  [string]$SourceUserData = "$env:USERPROFILE\AppData\Local\Microsoft\Edge\User Data",
  [string]$SourceProfile = "Default",
  [string]$TargetUserData = ".\artifacts\profiles\profile-edge-robot"
)

$ErrorActionPreference = "Stop"

$sourceUserDataResolved = [System.IO.Path]::GetFullPath($SourceUserData)
$targetUserDataResolved = [System.IO.Path]::GetFullPath($TargetUserData)
$sourceProfilePath = Join-Path $sourceUserDataResolved $SourceProfile
$targetProfilePath = Join-Path $targetUserDataResolved $SourceProfile

if (-not (Test-Path $sourceUserDataResolved)) {
  throw "No existe SourceUserData: $sourceUserDataResolved"
}

if (-not (Test-Path $sourceProfilePath)) {
  throw "No existe el perfil de Edge '$SourceProfile' en: $sourceUserDataResolved"
}

$edgeRunning = Get-Process msedge -ErrorAction SilentlyContinue
if ($edgeRunning) {
  throw "Edge esta abierto. Cierra Edge antes de clonar el perfil."
}

New-Item -ItemType Directory -Path $targetUserDataResolved -Force | Out-Null

# Archivos de raiz utiles para descifrar cookies y mantener metadata del navegador.
foreach ($rootFile in @("Local State", "First Run", "Last Version", "Last Browser")) {
  $src = Join-Path $sourceUserDataResolved $rootFile
  if (Test-Path $src) {
    Copy-Item -Path $src -Destination $targetUserDataResolved -Force
  }
}

# Limpieza previa para evitar basura de corridas anteriores.
if (Test-Path $targetProfilePath) {
  Remove-Item -Path $targetProfilePath -Recurse -Force
}
New-Item -ItemType Directory -Path $targetProfilePath -Force | Out-Null

$excludeDirs = @(
  "Cache",
  "Code Cache",
  "GPUCache",
  "ShaderCache",
  "GrShaderCache",
  "DawnCache",
  "DawnGraphiteCache",
  "DawnWebGPUCache",
  "Crashpad",
  "Service Worker\CacheStorage"
)

$excludeFiles = @(
  "LOCK",
  "SingletonLock",
  "SingletonCookie",
  "SingletonSocket"
)

$robocopyArgs = @(
  $sourceProfilePath,
  $targetProfilePath,
  "/MIR",
  "/R:1",
  "/W:1",
  "/NFL",
  "/NDL",
  "/NJH",
  "/NJS",
  "/NP"
)

if ($excludeDirs.Count -gt 0) {
  $robocopyArgs += "/XD"
  $robocopyArgs += $excludeDirs
}

if ($excludeFiles.Count -gt 0) {
  $robocopyArgs += "/XF"
  $robocopyArgs += $excludeFiles
}

& robocopy @robocopyArgs | Out-Null
$rc = $LASTEXITCODE

# Robocopy considera exito codigos 0..7.
if ($rc -gt 7) {
  throw "Robocopy fallo con codigo $rc."
}

Write-Host "Perfil clonado correctamente."
Write-Host "Origen: $sourceProfilePath"
Write-Host "Destino: $targetProfilePath"
Write-Host "Ya puedes ejecutar Playwright usando artifacts\\profiles\\profile-edge-robot."
