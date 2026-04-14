# ==============================================================================
# SCRIPT DE VERIFICACIÓN DE ALMACENAMIENTO
# Compara OneDrive (origen) vs SharePoint (destino)
#
# Autor: Alejandro Ariaz (@Ariaz39)
# Licencia: MIT License
# Repositorio: https://github.com/Ariaz39/sharepoint-from-onedrive-migration-tool
# ==============================================================================
#
# USO:
#   .\check-storage.ps1                                          # Usa el primer usuario del .env
#   .\check-storage.ps1 -UserEmail "user@domain.com"             # Usa un usuario específico
#   .\check-storage.ps1 -ScopeFolder "RUP"                       # Valida solo esa carpeta
#   .\check-storage.ps1 -ScopeFolder "RUP/CONTRATOS FINALIZADOS" # Subcarpeta específica
#
# NOTA: No hardcodea información de la empresa - lee del .env
# ==============================================================================

param(
    [string]$UserEmail   = "",  # Opcional - si no se especifica, usa el primer usuario del .env
    [string]$ScopeFolder = ""   # Opcional - valida solo esa carpeta (igual que en migrationScript.ps1)
)

# Cargar configuración del .env (mismo método que migrationScript.ps1)
$envFile = Join-Path $PSScriptRoot ".env"
if (-not (Test-Path $envFile)) {
    Write-Host "ERROR: Archivo .env no encontrado." -ForegroundColor Red
    exit 1
}

Get-Content $envFile | Where-Object { $_ -match "^[^#=]+=" } | ForEach-Object {
    $name, $value = $_ -split '=', 2
    # Limpiar comentarios inline (todo después de #)
    $cleanValue = ($value -split '#')[0].Trim()
    Set-Variable -Name $name.Trim() -Value $cleanValue -Scope Script
}

# Si no se especificó UserEmail, usar el primer usuario del .env
if ([string]::IsNullOrWhiteSpace($UserEmail)) {
    # IMPORTANTE: usar @() para forzar array
    $userList = @($USERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

    if ($userList.Count -eq 0) {
        Write-Host "ERROR: No hay usuarios configurados en .env (variable USERS)" -ForegroundColor Red
        exit 1
    }

    $UserEmail = $userList[0]
    Write-Host "Usando usuario del .env: $UserEmail" -ForegroundColor DarkGray
}

# Convertir EXCLUDED_FOLDERS a array
$excludedList = $EXCLUDED_FOLDERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

$certPath = Join-Path $PSScriptRoot "PnPMigrationCert.pfx"
$certPassword = ConvertTo-SecureString -String $CERT_PASSWORD -AsPlainText -Force

# Construir URLs
$upnPrefix = $UserEmail.ToLower().Replace(".", "_").Replace("@", "_")
$myUrlBase = $MY_URL -replace '/personal$', ''
$oneDriveUrl = "$myUrlBase/personal/$upnPrefix"
$baseDocUrl = "/personal/$upnPrefix/Documents"

# Scope: prefijo de ruta en OneDrive y en SharePoint destino
# El FileRef en SharePoint destino tiene la forma: /sites/<site>/<library>/<ScopeFolder>/...
$destSiteName      = $DESTINATION_SITE_URL.Split('/')[-1]
$scopeSourcePrefix = if ($ScopeFolder -ne "") { "$baseDocUrl/$ScopeFolder".TrimEnd('/') } else { "" }
$scopeDestPrefix   = if ($ScopeFolder -ne "") { "/sites/$destSiteName/$DESTINATION_LIBRARY/$ScopeFolder".TrimEnd('/') } else { "" }

Write-Host "`n=== COMPARACIÓN DE ALMACENAMIENTO ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Usuario: $UserEmail" -ForegroundColor Yellow
if ($ScopeFolder -ne "") {
    Write-Host "Scope:   $ScopeFolder" -ForegroundColor Yellow
}
Write-Host "⚠️  ADVERTENCIA: Los reportes generados pueden contener información sensible." -ForegroundColor Yellow
Write-Host "   No compartir los archivos CSV fuera de la organización." -ForegroundColor DarkGray
Write-Host ""

# ONEDRIVE
Write-Host "=== ONEDRIVE (Origen) ===" -ForegroundColor Green
try {
    Write-Host "  Conectando a OneDrive..." -ForegroundColor DarkGray
    $sourceConn = Connect-PnPOnline -Url $oneDriveUrl -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection

    Write-Host "  Escaneando archivos (puede tardar si hay muchos archivos)..." -ForegroundColor DarkGray
    $sourceFiles = Get-PnPListItem -List "Documents" -PageSize 1000 -Connection $sourceConn | Where-Object {
        $_.FileSystemObjectType -eq "File" -and
        ($scopeSourcePrefix -eq "" -or $_.FieldValues.FileRef.StartsWith($scopeSourcePrefix, [System.StringComparison]::InvariantCultureIgnoreCase))
    }
    Write-Host "  Archivos encontrados: $($sourceFiles.Count)" -ForegroundColor Cyan

    # Separar archivos: excluidos vs a migrar
    $totalSizeBytes = 0
    $fileCount = 0
    $excludedSizeBytes = 0
    $excludedFileCount = 0
    $toMigrateSizeBytes = 0
    $toMigrateFileCount = 0
    $excludedByFolder = @{}  # Contador por carpeta excluida

    $zeroSizeSource = [System.Collections.Generic.List[string]]::new()

    foreach ($file in $sourceFiles) {
        # File_x0020_Size puede ser null en algunos tipos — tratar null/0 como 0 pero siempre contar el archivo
        $size = [long]($file.FieldValues.File_x0020_Size)
        $totalSizeBytes += $size
        $fileCount++

        if ($size -eq 0) { $zeroSizeSource.Add($file.FieldValues.FileRef) }

        # Verificar si está en carpeta excluida
        $sourceRelUrl = $file.FieldValues.FileRef
        $pathParts = $sourceRelUrl.Replace($baseDocUrl, "").Trim('/').Split('/')
        $directoryParts = if ($pathParts.Count -gt 1) { $pathParts[0..($pathParts.Count - 2)] } else { @() }

        $isExcluded = $false
        foreach ($excl in $excludedList) {
            if ($directoryParts -contains $excl) {
                $isExcluded = $true
                $excludedSizeBytes += $size
                $excludedFileCount++

                # Contar por carpeta excluida
                if (-not $excludedByFolder.ContainsKey($excl)) {
                    $excludedByFolder[$excl] = @{ Count = 0; Size = 0 }
                }
                $excludedByFolder[$excl].Count++
                $excludedByFolder[$excl].Size += $size
                break
            }
        }

        if (-not $isExcluded) {
            $toMigrateSizeBytes += $size
            $toMigrateFileCount++
        }
    }

    if ($zeroSizeSource.Count -gt 0) {
        Write-Host "  Archivos con tamaño 0 o nulo en origen: $($zeroSizeSource.Count)" -ForegroundColor Yellow
        $zeroSizeSource | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }
    }

    $totalSizeGB = [math]::Round($totalSizeBytes / 1GB, 2)
    $excludedSizeGB = [math]::Round($excludedSizeBytes / 1GB, 2)
    $toMigrateSizeGB = [math]::Round($toMigrateSizeBytes / 1GB, 2)

    Write-Host "  Total archivos: $fileCount" -ForegroundColor White
    Write-Host "  Total tamaño: $totalSizeGB GB" -ForegroundColor White
    Write-Host ""
    Write-Host "  Archivos excluidos: $excludedFileCount" -ForegroundColor DarkGray
    Write-Host "  Tamaño excluido: $excludedSizeGB GB" -ForegroundColor DarkGray

    # Desglose por carpeta excluida
    if ($excludedByFolder.Count -gt 0) {
        Write-Host "  Desglose de exclusiones:" -ForegroundColor DarkGray
        foreach ($folder in $excludedByFolder.Keys | Sort-Object) {
            $folderSizeGB = [math]::Round($excludedByFolder[$folder].Size / 1GB, 2)
            Write-Host "    • $folder : $($excludedByFolder[$folder].Count) archivos, $folderSizeGB GB" -ForegroundColor DarkGray
        }
    }

    Write-Host ""
    Write-Host "  Archivos a migrar: $toMigrateFileCount" -ForegroundColor Cyan
    Write-Host "  Tamaño a migrar: $toMigrateSizeGB GB" -ForegroundColor Cyan
} catch {
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# SHAREPOINT
Write-Host ""
Write-Host "=== SHAREPOINT (Destino) ===" -ForegroundColor Green
try {
    Write-Host "  Conectando a SharePoint..." -ForegroundColor DarkGray
    $destConn = Connect-PnPOnline -Url $DESTINATION_SITE_URL -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection

    Write-Host "  Escaneando archivos (puede tardar si hay muchos archivos)..." -ForegroundColor DarkGray
    $destFiles = Get-PnPListItem -List $DESTINATION_LIBRARY -PageSize 1000 -Connection $destConn | Where-Object {
        $_.FileSystemObjectType -eq "File" -and
        ($scopeDestPrefix -eq "" -or $_.FieldValues.FileRef.StartsWith($scopeDestPrefix, [System.StringComparison]::InvariantCultureIgnoreCase))
    }
    Write-Host "  Archivos encontrados: $($destFiles.Count)" -ForegroundColor Cyan

    $totalSizeBytesDestino = 0
    $fileCountDestino = 0
    $zeroSizeDest = [System.Collections.Generic.List[string]]::new()

    foreach ($file in $destFiles) {
        $size = [long]($file.FieldValues.File_x0020_Size)
        $totalSizeBytesDestino += $size
        $fileCountDestino++
        if ($size -eq 0) { $zeroSizeDest.Add($file.FieldValues.FileRef) }
    }

    if ($zeroSizeDest.Count -gt 0) {
        Write-Host "  Archivos con tamaño 0 o nulo en destino: $($zeroSizeDest.Count)" -ForegroundColor Yellow
        $zeroSizeDest | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }
    }

    $totalSizeGBDestino = [math]::Round($totalSizeBytesDestino / 1GB, 2)

    Write-Host "  Archivos: $fileCountDestino"
    Write-Host "  Tamaño total: $totalSizeGBDestino GB" -ForegroundColor Cyan
} catch {
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# COMPARACIÓN (contra archivos que SÍ deben migrarse)
$diffFiles    = $toMigrateFileCount - $fileCountDestino
$diffBytes    = $toMigrateSizeBytes - $totalSizeBytesDestino
$percentMigrated = if ($toMigrateFileCount -gt 0) { [math]::Round(($fileCountDestino / $toMigrateFileCount) * 100, 1) } else { 0 }
$resultadoTexto  = if ($percentMigrated -ge 99.5 -and [math]::Abs($diffBytes) -lt 100MB) { "EXITOSA" } elseif ($percentMigrated -ge 95) { "CASI COMPLETA" } else { "INCOMPLETA" }

Write-Host ""
Write-Host "=== COMPARACIÓN ===" -ForegroundColor Cyan
Write-Host "  Archivos esperados (sin exclusiones): $toMigrateFileCount"
Write-Host "  Archivos en SharePoint:               $fileCountDestino"
Write-Host "  Porcentaje migrado: $percentMigrated%" -ForegroundColor $(if ($percentMigrated -ge 99) { 'Green' } elseif ($percentMigrated -ge 95) { 'Yellow' } else { 'Red' })
Write-Host ""

if ([math]::Abs($diffFiles) -eq 0) {
    Write-Host "  ✓ Todos los archivos migrados correctamente" -ForegroundColor Green
} elseif ($diffFiles -gt 0) {
    Write-Host "  ⚠ Archivos faltantes: $diffFiles" -ForegroundColor Yellow
} else {
    Write-Host "  ℹ Archivos extra en destino: $([math]::Abs($diffFiles))" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "=== RESULTADO ===" -ForegroundColor Cyan
if ($resultadoTexto -eq "EXITOSA") {
    Write-Host "  ✓ MIGRACIÓN EXITOSA" -ForegroundColor Green
    Write-Host "  Todos los archivos fueron migrados correctamente" -ForegroundColor Green
} elseif ($resultadoTexto -eq "CASI COMPLETA") {
    Write-Host "  ⚡ MIGRACIÓN CASI COMPLETA" -ForegroundColor Yellow
    Write-Host "  Revisar archivos faltantes en el log de migración" -ForegroundColor Yellow
} else {
    Write-Host "  ✗ MIGRACIÓN INCOMPLETA" -ForegroundColor Red
    Write-Host "  Se requiere investigar qué archivos no se migraron" -ForegroundColor Red
}

# GENERAR REPORTE CSV PARA GERENCIA
$logPath    = Join-Path $PSScriptRoot ($LOG_DIRECTORY)
if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath | Out-Null }
$scopeLabel = if ($ScopeFolder -ne "") { $ScopeFolder -replace '[/\\]', '-' } else { "completo" }
$reportFile = Join-Path $logPath "reporte-migracion-$($UserEmail.Split('@')[0])-$scopeLabel-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"

$lines = @(
    "REPORTE DE MIGRACIÓN"
    "Fecha;$((Get-Date).ToString('dd/MM/yyyy HH:mm'))"
    "Usuario;$UserEmail"
    "Carpeta migrada;$(if ($ScopeFolder -ne '') { $ScopeFolder } else { 'Completo' })"
    ""
    "ORIGEN (OneDrive)"
    "Total archivos origen;$toMigrateFileCount"
    ""
    "DESTINO (SharePoint)"
    "Total archivos migrados;$fileCountDestino"
    "Porcentaje completado;$percentMigrated%"
    ""
    "RESULTADO;$resultadoTexto"
)
$lines | Out-File $reportFile -Encoding UTF8
Write-Host ""
Write-Host "  Reporte generado: $reportFile" -ForegroundColor DarkGray
Write-Host ""
