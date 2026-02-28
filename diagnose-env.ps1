# ============================================================
# Diagnostic script for .env file
#
# Autor: Alejandro Ariaz (@Ariaz39)
# Licencia: MIT License
# Repositorio: https://github.com/Ariaz39/sharepoint-from-onedrive-migration-tool
# ============================================================

$basePath = $PSScriptRoot
$envPath = Join-Path $basePath ".env"

Write-Host "============================================================" -ForegroundColor Magenta
Write-Host "Diagnosing .env file" -ForegroundColor Magenta
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host ""

if (-not (Test-Path $envPath)) {
    Write-Host "ERROR: .env file not found at: $envPath" -ForegroundColor Red
    exit 1
}

Write-Host "File path: $envPath" -ForegroundColor Cyan
$fileInfo = Get-Item $envPath
Write-Host "File size: $($fileInfo.Length) bytes" -ForegroundColor White
Write-Host "Last modified: $($fileInfo.LastWriteTime)" -ForegroundColor White
Write-Host ""

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Raw file content (first 50 lines):" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$lineNumber = 1
Get-Content $envPath -Encoding UTF8 | Select-Object -First 50 | ForEach-Object {
    $line = $_
    Write-Host "$($lineNumber.ToString().PadLeft(3)): $line" -ForegroundColor White

    # Show line details if it contains USERS
    if ($line -match "^USERS=") {
        Write-Host "     ^-- USERS LINE FOUND" -ForegroundColor Yellow
        Write-Host "     Length: $($line.Length) chars" -ForegroundColor Yellow
        Write-Host "     Bytes: $([System.Text.Encoding]::UTF8.GetBytes($line).Length)" -ForegroundColor Yellow

        $parts = $line -split "=", 2
        if ($parts.Count -eq 2) {
            $value = $parts[1].Trim()
            Write-Host "     Raw value: '$value'" -ForegroundColor Yellow
            Write-Host "     Value length: $($value.Length) chars" -ForegroundColor Yellow
        }
    }

    $lineNumber++
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Parsing USERS variable:" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$usersLine = Get-Content $envPath -Encoding UTF8 | Where-Object { $_ -match "^USERS=" -and -not $_.Trim().StartsWith("#") } | Select-Object -First 1

if ($usersLine) {
    Write-Host "Found USERS line: $usersLine" -ForegroundColor Green

    $parts = $usersLine -split "=", 2
    if ($parts.Count -eq 2) {
        $rawValue = $parts[1]
        Write-Host "Raw value (before trim): '$rawValue'" -ForegroundColor White

        $trimmedValue = $rawValue.Trim()
        Write-Host "After trim: '$trimmedValue'" -ForegroundColor White

        # Remove quotes if present
        $finalValue = $trimmedValue
        if ($trimmedValue.StartsWith('"') -and $trimmedValue.EndsWith('"')) {
            $finalValue = $trimmedValue.Substring(1, $trimmedValue.Length - 2)
            Write-Host "After removing quotes: '$finalValue'" -ForegroundColor White
        }

        Write-Host ""
        Write-Host "Splitting by comma (OLD METHOD - may fail):" -ForegroundColor Yellow
        $usersOld = $finalValue -split "," | ForEach-Object { $_.Trim() }
        Write-Host "  Result type: $($usersOld.GetType().Name)" -ForegroundColor DarkGray
        Write-Host "  First element: '$($usersOld[0])'" -ForegroundColor DarkGray
        Write-Host "  First element type: $($usersOld[0].GetType().Name)" -ForegroundColor DarkGray

        Write-Host ""
        Write-Host "Splitting by comma (NEW METHOD - with @() wrapper):" -ForegroundColor Green
        $users = @($finalValue -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        Write-Host "  Result type: $($users.GetType().Name)" -ForegroundColor DarkGray
        Write-Host "  Count: $($users.Count)" -ForegroundColor White

        for ($i = 0; $i -lt $users.Count; $i++) {
            $user = [string]$users[$i]  # Force to string
            Write-Host ""
            Write-Host "  User[$i]: '$user'" -ForegroundColor White
            Write-Host "    Length: $($user.Length) chars" -ForegroundColor DarkGray
            Write-Host "    Type: $($user.GetType().Name)" -ForegroundColor DarkGray

            if ($user -and $user.Length -gt 0) {
                # Test URL generation
                $testSlug = $user.Replace("@", "_").Replace(".", "_")
                Write-Host "    Slug would be: '$testSlug'" -ForegroundColor Green
                Write-Host "    Contains @: $($user.Contains('@'))" -ForegroundColor DarkGray
            } else {
                Write-Host "    WARNING: Empty user!" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "WARNING: No uncommented USERS line found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Lines containing 'USERS':" -ForegroundColor Yellow
    Get-Content $envPath -Encoding UTF8 | Select-String -Pattern "USERS" | ForEach-Object {
        Write-Host "  Line $($_.LineNumber): $($_.Line)" -ForegroundColor White
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host "Diagnosis complete" -ForegroundColor Magenta
Write-Host "============================================================" -ForegroundColor Magenta
