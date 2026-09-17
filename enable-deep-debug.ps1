$ErrorActionPreference = "Stop"

Write-Host "Enabling Legion RGB deep debug mode for this PowerShell session..." -ForegroundColor Cyan

# Deep diagnostics from driver/app
$env:LEGION_RGB_REVERSE_MODE = "1"
$env:RUST_BACKTRACE = "full"

# Ensure LOQ uses LampArray writes (do not force vendor-only)
$env:LEGION_RGB_LOQ_USE_LAMPARRAY = "1"
Remove-Item Env:LEGION_RGB_FORCE_VENDOR_ONLY -ErrorAction SilentlyContinue

# scrap/nokhwa build scripts require VCPKG_ROOT on Windows.
$localVcpkgRoot = Join-Path $PSScriptRoot "target\vcpkg_repo"
if (Test-Path $localVcpkgRoot) {
    $env:VCPKG_ROOT = $localVcpkgRoot
} elseif (-not [string]::IsNullOrWhiteSpace($env:VCPKG_ROOT) -and (Test-Path $env:VCPKG_ROOT)) {
    # Keep caller-provided VCPKG_ROOT when valid.
} else {
    Write-Warning "VCPKG_ROOT was not found automatically. Build may fail until vcpkg is installed."
}

Write-Host "" 
Write-Host "Environment variables now active:" -ForegroundColor Green
Write-Host "LEGION_RGB_REVERSE_MODE=$env:LEGION_RGB_REVERSE_MODE"
Write-Host "RUST_BACKTRACE=$env:RUST_BACKTRACE"
Write-Host "LEGION_RGB_LOQ_USE_LAMPARRAY=$env:LEGION_RGB_LOQ_USE_LAMPARRAY"
if (-not [string]::IsNullOrWhiteSpace($env:VCPKG_ROOT)) {
    Write-Host "VCPKG_ROOT=$env:VCPKG_ROOT"
} else {
    Write-Host "VCPKG_ROOT=<unset>"
}
if (Test-Path Env:LEGION_RGB_FORCE_VENDOR_ONLY) {
    Write-Host "LEGION_RGB_FORCE_VENDOR_ONLY=$env:LEGION_RGB_FORCE_VENDOR_ONLY"
} else {
    Write-Host "LEGION_RGB_FORCE_VENDOR_ONLY=<unset>"
}

Write-Host ""
Write-Host "Next commands to run in this SAME PowerShell window:" -ForegroundColor Yellow
Write-Host "  if (-not (Test-Path \"$env:VCPKG_ROOT\")) { .\\target\\vcpkg_repo\\bootstrap-vcpkg.bat }"
Write-Host "  cargo build --release"
Write-Host "  cargo run --release"
Write-Host ""
Write-Host "Debug file will be created next to the running executable as legion_rgb_debug.log" -ForegroundColor Yellow
