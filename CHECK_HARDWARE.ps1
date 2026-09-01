$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host 'GMP Offline OCR - Hardware Check' -ForegroundColor Cyan
Write-Host '================================' -ForegroundColor Cyan

$gpus = Get-CimInstance Win32_VideoController

Write-Host ''
Write-Host 'Detected display adapters:' -ForegroundColor Yellow
foreach ($gpu in $gpus) {
    $vramGb = if ($gpu.AdapterRAM) { [math]::Round(([uint64]$gpu.AdapterRAM / 1GB), 1) } else { 'unknown' }
    Write-Host "- $($gpu.Name) | VRAM: $vramGb GB | Driver: $($gpu.DriverVersion)"
}

$nvidiaAdapters = @($gpus | Where-Object { $_.Name -match 'NVIDIA' })
if ($nvidiaAdapters.Count -eq 0) {
    Write-Host ''
    Write-Host 'Result: no NVIDIA GPU detected.' -ForegroundColor Yellow
    Write-Host 'The local OCR deployment will use CPU mode. It is compatible, but significantly slower than a supported CUDA GPU.' -ForegroundColor Yellow
    exit 0
}

$nvidiaSmi = Get-Command nvidia-smi -ErrorAction SilentlyContinue
if (-not $nvidiaSmi) {
    Write-Host ''
    Write-Host 'Result: NVIDIA GPU detected, but nvidia-smi is unavailable.' -ForegroundColor Yellow
    Write-Host 'The local OCR deployment will use CPU mode. Install or update the NVIDIA driver only if CUDA acceleration is added later.' -ForegroundColor Yellow
    exit 0
}

Write-Host ''
Write-Host 'NVIDIA CUDA devices:' -ForegroundColor Yellow
$deviceInfo = & nvidia-smi --query-gpu=name,memory.total,driver_version --format=csv,noheader 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host $deviceInfo -ForegroundColor Red
    exit 1
}
$deviceInfo | ForEach-Object { Write-Host "- $_" }

$memoryValues = & nvidia-smi --query-gpu=memory.total --format=csv,noheader,nounits
$maxMemoryMb = ($memoryValues | ForEach-Object { [int]$_.Trim() } | Measure-Object -Maximum).Maximum

Write-Host ''
if ($maxMemoryMb -ge 16000) {
    Write-Host "Result: $maxMemoryMb MB NVIDIA VRAM detected." -ForegroundColor Green
    Write-Host 'The current deployment still uses CPU mode for consistent GTX 1050-compatible behavior.'
    exit 0
}

if ($maxMemoryMb -ge 12000) {
    Write-Host "Result: NVIDIA GPU detected with $maxMemoryMb MB VRAM." -ForegroundColor Yellow
    Write-Host 'The current deployment uses CPU mode; this VRAM amount does not affect it.'
    exit 0
}

Write-Host "Result: $maxMemoryMb MB VRAM detected." -ForegroundColor Yellow
Write-Host 'The current deployment uses CPU mode, so it remains usable despite limited VRAM.'
exit 0
