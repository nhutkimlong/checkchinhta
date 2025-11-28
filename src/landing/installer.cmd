<# :
@echo off
:: Batch section
title Cai dat AI Check Chinh Ta
echo Dang khoi dong trinh cai dat...
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ([System.IO.File]::ReadAllText('%~f0'))"
pause
goto :eof
#>

# PowerShell section starts here
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$AddinName = "AICheckChinhTa"
$ManifestUrl = "https://checkchinhta.netlify.app/manifest.xml"

# Lay duong dan Documents chuan cua he thong (bat ke ngon ngu Win)
$DocumentsPath = [Environment]::GetFolderPath("MyDocuments")
$TargetDir = Join-Path $DocumentsPath "AddIns\$AddinName"

Write-Host "Dang cai dat $AddinName..." -ForegroundColor Cyan

# 1. Tao thu muc
if (!(Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    Write-Host "Da tao thu muc: $TargetDir" -ForegroundColor Gray
}

# 2. Chuyen doi sang UNC path (\\localhost\...)
# Day la buoc quan trong de Word nhan dien la Network Share
$Drive = $TargetDir.Substring(0, 1)
$PathWithoutDrive = $TargetDir.Substring(2)
$UncPath = "\\localhost\$Drive`$$PathWithoutDrive"

# Fallback neu UNC khong truy cap duoc
if (!(Test-Path $UncPath)) {
    Write-Warning "Khong truy cap duoc UNC path. Su dung path cuc bo."
    $UncPath = $TargetDir
}

# 3. Tai manifest
$ManifestPath = Join-Path $TargetDir "manifest.xml"
Write-Host "Dang tai manifest..." -ForegroundColor Gray
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $ManifestUrl -OutFile $ManifestPath -UseBasicParsing
} catch {
    Write-Error "Loi tai file: $_"
    Write-Host "Vui long kiem tra ket noi Internet." -ForegroundColor Red
    Exit
}

# 4. Dang ky Registry
$TrustCenterKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if (!(Test-Path $TrustCenterKey)) { New-Item -Path $TrustCenterKey -Force | Out-Null }

# Xoa key cu
Get-ChildItem $TrustCenterKey | ForEach-Object {
    $val = (Get-ItemProperty -Path $_.PSPath -Name "Url" -ErrorAction SilentlyContinue).Url
    if ($val -like "*$AddinName*" -or $val -like "*ProtonXCorrector*") {
        Remove-Item -Path $_.PSPath -Force
    }
}

# Tao key moi
$CatalogId = "{" + [guid]::NewGuid().ToString() + "}"
$NewKey = "$TrustCenterKey\$CatalogId"
New-Item -Path $NewKey -Force | Out-Null
Set-ItemProperty -Path $NewKey -Name "Id" -Value $CatalogId
Set-ItemProperty -Path $NewKey -Name "Url" -Value $UncPath
Set-ItemProperty -Path $NewKey -Name "Flags" -Value 1 -Type DWord

Write-Host "------------------------------------------------"
Write-Host "CAI DAT THANH CONG!" -ForegroundColor Green
Write-Host "------------------------------------------------"
Write-Host "HUONG DAN:" -ForegroundColor Yellow
Write-Host "1. Khoi dong lai Word."
Write-Host "2. Vao Insert > Add-ins > Shared Folder."
Write-Host "3. Chon '$AddinName' va nhan Add."
Write-Host "------------------------------------------------"
