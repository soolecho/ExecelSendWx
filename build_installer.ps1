param(
    [string]$CertificateThumbprint = "191C64E4EC07377CA032878878D0A45F554C8146",
    [string]$TimestampServer = "http://timestamp.digicert.com"
)
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$installerScript = Join-Path $projectRoot "installer.iss"
$outputDir = Join-Path $projectRoot "installer_output"
$sourceDir = Join-Path $projectRoot "dist_nuitka_config\main.dist"
$iscc = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $iscc)) { $iscc = Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe" }
if (-not (Test-Path $iscc)) { throw "ISCC not found." }
if (-not (Test-Path $installerScript)) { throw "Installer script not found: $installerScript" }

$LogDir = Join-Path $projectRoot "logs"
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$logPath = Join-Path $LogDir ("build_installer_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

function Write-Tee($msg) {
    $line = ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg)
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

Write-Tee "Starting build_installer"
Write-Tee ("SourceDir: {0}" -f $sourceDir)
Write-Tee ("OutputDir: {0}" -f $outputDir)
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$running = Get-CimInstance Win32_Process | Where-Object {
    $_.ExecutablePath -and
    $_.ExecutablePath.StartsWith($sourceDir, [System.StringComparison]::OrdinalIgnoreCase)
} | Select-Object -First 1
if ($running) { throw "Please close packaged app before building: $($running.ExecutablePath)" }

$certificates = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object {
    $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(30)
}
$certificate = if ($CertificateThumbprint) {
    $certificates | Where-Object Thumbprint -eq $CertificateThumbprint | Select-Object -First 1
} else {
    $certificates | Sort-Object NotAfter -Descending | Select-Object -First 1
}
if (-not $certificate) { throw "Valid code-signing certificate not found." }
Write-Tee ("Signer: {0} Thumbprint={1}" -f $certificate.Subject, $certificate.Thumbprint)

function Set-CodeSignature {
    param([Parameter(Mandatory=$true)][string]$FilePath)
    $result = Set-AuthenticodeSignature -LiteralPath $FilePath -Certificate $certificate -HashAlgorithm SHA256 -IncludeChain All -TimestampServer $TimestampServer
    if ($result.Status -ne "Valid") { throw "Signing failed: $FilePath ; $($result.StatusMessage)" }
    Write-Tee ("Signed: {0}" -f [IO.Path]::GetFileName($FilePath))
}

function Invoke-InnoCompiler {
    param([string]$Desc)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $output = & $iscc /Qp $installerScript 2>&1
    $code = $LASTEXITCODE
    $ErrorActionPreference = $prev
    foreach ($line in $output) { $line = "$line"; Write-Tee ("ISCC[{0}]: {1}" -f $Desc, $line) }
    return [int]$code
}

# Stage 1: 首次编译；SignedUninstaller=yes 时 ISCC 会以 exit!=0 提示签名 uninst*.e32
Write-Tee "Stage 1: compile to generate unsigned uninstaller"
$before = @(Get-ChildItem $outputDir -File -Filter "uninst*.e32" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
$code = Invoke-InnoCompiler -Desc "S1"
$afterAll = @(Get-ChildItem $outputDir -File -Filter "uninst*.e32" -ErrorAction SilentlyContinue)
$newUninstallers = $afterAll | Where-Object { $before -notcontains $_.FullName }
if ($newUninstallers.Count -eq 0) {
    $newUninstallers = @($afterAll | Where-Object { (Get-AuthenticodeSignature $_.FullName).Status -ne "Valid" })
}
Write-Tee ("Stage 1 exit={0} new_uninstallers={1}" -f $code, $newUninstallers.Count)

if ($newUninstallers.Count -gt 0) {
    foreach ($u in $newUninstallers) {
        Write-Tee ("Signing uninstaller: {0}" -f $u.Name)
        Set-CodeSignature -FilePath $u.FullName
    }
}

# Stage 2: 重试编译，最多 4 次
$code2 = -1
for ($i = 1; $i -le 4; $i++) {
    Start-Sleep -Seconds (2 * $i)
    Write-Tee ("Stage 2: recompile retry {0}/4" -f $i)
    $code2 = Invoke-InnoCompiler -Desc ("S2-R$i")
    Write-Tee ("Stage 2 retry {0} exit={1}" -f $i, $code2)
    if ($code2 -eq 0) { break }
}
if ($code2 -ne 0) { throw "ISCC failed with exit code $code2 after retries." }

$installer = Get-ChildItem $outputDir -File -Filter "*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $installer) { throw "Installer output not found." }
Write-Tee ("Signing installer: {0}" -f $installer.Name)
Set-CodeSignature -FilePath $installer.FullName

$publicCert = Join-Path $outputDir "code_signing_public.cer"
Export-Certificate -Cert $certificate -FilePath $publicCert -Type CERT -Force | Out-Null

$hash = Get-FileHash $installer.FullName -Algorithm SHA256
$sig = Get-AuthenticodeSignature $installer.FullName
Write-Tee ""
Write-Tee "Installer build completed."
Write-Tee ("Path: {0}" -f $installer.FullName)
Write-Tee ("Size: {0:N0} bytes" -f $installer.Length)
Write-Tee ("Signed: {0}" -f $sig.Status)
Write-Tee ("Signer: {0}" -f $sig.SignerCertificate.Subject)
Write-Tee ("SHA256: {0}" -f $hash.Hash)
exit 0
