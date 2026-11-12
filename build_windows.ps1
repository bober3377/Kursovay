$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

py -3 -m pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) {
    throw "Не удалось установить зависимости сборки."
}

py -3 -m PyInstaller `
    --clean `
    --noconfirm `
    --windowed `
    --onedir `
    --name ElectronicTextbook `
    --add-data "schema.sql;." `
    --collect-all psycopg `
    main.py
if ($LASTEXITCODE -ne 0) {
    throw "Сборка приложения завершилась с ошибкой."
}

$packageDirectory = Join-Path $PSScriptRoot "dist\ElectronicTextbook"
Copy-Item `
    -LiteralPath (Join-Path $PSScriptRoot "config.example.ini") `
    -Destination (Join-Path $packageDirectory "config.ini") `
    -Force

$archivePath = Join-Path $PSScriptRoot "ElectronicTextbook-Windows.zip"
$archiveCreated = $false
for ($attempt = 1; $attempt -le 3; $attempt++) {
    try {
        if (Test-Path -LiteralPath $archivePath) {
            Remove-Item -LiteralPath $archivePath
        }
        Start-Sleep -Seconds 2
        Compress-Archive `
            -Path $packageDirectory `
            -DestinationPath $archivePath `
            -CompressionLevel Optimal
        $archiveCreated = $true
        break
    } catch {
        if ($attempt -eq 3) {
            throw
        }
        Start-Sleep -Seconds 2
    }
}
if (-not $archiveCreated) {
    throw "Не удалось создать ZIP-архив приложения."
}

Write-Host "Сборка готова:"
Write-Host "  Приложение: $packageDirectory\ElectronicTextbook.exe"
Write-Host "  Архив для распространения: $archivePath"
Write-Host "Перед запуском откройте config.ini и укажите подключение к PostgreSQL."
