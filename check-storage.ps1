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

# Función auxiliar: escaneo paginado por carpeta (evita timeout en bibliotecas grandes)
function Get-FilesPagedByFolder {
    param([string]$List, [string]$FolderUrl, $Conn, [int]$PageSize)
    $result   = [System.Collections.Generic.List[object]]::new()
    $position = $null
    $caml     = "<View Scope='FilesOnly'><RowLimit Paged='TRUE'>$PageSize</RowLimit></View>"
    do {
        $pageArgs = @{ List = $List; Query = $caml; Connection = $Conn; FolderServerRelativeUrl = $FolderUrl; PageSize = $PageSize; ErrorAction = "Stop" }
        if ($position) { $pageArgs["ListItemCollectionPosition"] = $position }
        $page     = Get-PnPListItem @pageArgs
        $position = $page.ListItemCollectionPosition
        $page | Where-Object { $_.FileSystemObjectType -eq "File" } | ForEach-Object { $result.Add($_) }
    } while ($position)
    return ,$result
}

function Get-AllSubFoldersCheck {
    param([string]$FolderUrl, $Conn)
    $queue  = [System.Collections.Generic.Queue[string]]::new()
    $result = [System.Collections.Generic.List[string]]::new()
    $queue.Enqueue($FolderUrl)
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $result.Add($current)
        try {
            $folder = Get-PnPFolder -Url $current -Connection $Conn -ErrorAction Stop
            $folder.Context.Load($folder.Folders)
            $folder.Context.ExecuteQuery()
            foreach ($sf in $folder.Folders) {
                if ($sf.Name -notin @('Forms', '_vti_cnf')) { $queue.Enqueue($sf.ServerRelativeUrl) }
            }
        } catch {}
    }
    return ,$result
}

# ONEDRIVE
Write-Host "=== ONEDRIVE (Origen) ===" -ForegroundColor Green
try {
    Write-Host "  Conectando a OneDrive..." -ForegroundColor DarkGray
    $sourceConn = Connect-PnPOnline -Url $oneDriveUrl -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection

    $scanRoot = if ($scopeSourcePrefix -ne "") { $scopeSourcePrefix } else { $baseDocUrl }
    Write-Host "  Expandiendo árbol de carpetas: $scanRoot" -ForegroundColor DarkGray
    $sourceFolders = Get-AllSubFoldersCheck -FolderUrl $scanRoot -Conn $sourceConn
    Write-Host "  Escaneando $($sourceFolders.Count) carpetas (paginado)..." -ForegroundColor DarkGray

    $totalSizeBytes    = 0
    $fileCount         = 0
    $excludedSizeBytes = 0
    $excludedFileCount = 0
    $toMigrateSizeBytes = 0
    $toMigrateFileCount = 0
    $excludedByFolder  = @{}
    $zeroSizeSource    = [System.Collections.Generic.List[string]]::new()

    foreach ($folder in $sourceFolders) {
        $pageFiles = Get-FilesPagedByFolder -List "Documents" -FolderUrl $folder -Conn $sourceConn -PageSize 500
        foreach ($file in $pageFiles) {
            $size = [long]($file.FieldValues.File_x0020_Size)
            $totalSizeBytes += $size
            $fileCount++
            if ($size -eq 0) { $zeroSizeSource.Add($file.FieldValues.FileRef) }

            $sourceRelUrl   = $file.FieldValues.FileRef
            $pathParts      = $sourceRelUrl.Replace($baseDocUrl, "").Trim('/').Split('/')
            $directoryParts = if ($pathParts.Count -gt 1) { $pathParts[0..($pathParts.Count - 2)] } else { @() }

            $isExcluded = $false
            foreach ($excl in $excludedList) {
                if ($directoryParts -contains $excl) {
                    $isExcluded = $true
                    $excludedSizeBytes += $size
                    $excludedFileCount++
                    if (-not $excludedByFolder.ContainsKey($excl)) { $excludedByFolder[$excl] = @{ Count = 0; Size = 0 } }
                    $excludedByFolder[$excl].Count++
                    $excludedByFolder[$excl].Size += $size
                    break
                }
            }
            if (-not $isExcluded) { $toMigrateSizeBytes += $size; $toMigrateFileCount++ }
        }
    }

    if ($zeroSizeSource.Count -gt 0) {
        Write-Host "  Archivos con tamaño 0 o nulo en origen: $($zeroSizeSource.Count)" -ForegroundColor Yellow
        $zeroSizeSource | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }
    }

    $totalSizeGB    = [math]::Round($totalSizeBytes / 1GB, 2)
    $excludedSizeGB = [math]::Round($excludedSizeBytes / 1GB, 2)
    $toMigrateSizeGB = [math]::Round($toMigrateSizeBytes / 1GB, 2)

    Write-Host "  Total archivos: $fileCount" -ForegroundColor White
    Write-Host "  Total tamaño: $totalSizeGB GB" -ForegroundColor White
    Write-Host ""
    Write-Host "  Archivos excluidos: $excludedFileCount" -ForegroundColor DarkGray
    Write-Host "  Tamaño excluido: $excludedSizeGB GB" -ForegroundColor DarkGray

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

    $destScanRoot = if ($scopeDestPrefix -ne "") { $scopeDestPrefix } else { "/sites/$destSiteName/$DESTINATION_LIBRARY" }
    Write-Host "  Expandiendo árbol de carpetas: $destScanRoot" -ForegroundColor DarkGray
    $destFolders = Get-AllSubFoldersCheck -FolderUrl $destScanRoot -Conn $destConn
    Write-Host "  Escaneando $($destFolders.Count) carpetas (paginado)..." -ForegroundColor DarkGray

    $totalSizeBytesDestino = 0
    $fileCountDestino      = 0
    $zeroSizeDest          = [System.Collections.Generic.List[string]]::new()

    foreach ($folder in $destFolders) {
        $pageFiles = Get-FilesPagedByFolder -List $DESTINATION_LIBRARY -FolderUrl $folder -Conn $destConn -PageSize 500
        foreach ($file in $pageFiles) {
            $size = [long]($file.FieldValues.File_x0020_Size)
            $totalSizeBytesDestino += $size
            $fileCountDestino++
            if ($size -eq 0) { $zeroSizeDest.Add($file.FieldValues.FileRef) }
        }
    }
    Write-Host "  Archivos encontrados: $fileCountDestino" -ForegroundColor Cyan

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
Write-Host "  Tamaño origen  (bytes exactos): $toMigrateSizeBytes"
Write-Host "  Tamaño destino (bytes exactos): $totalSizeBytesDestino"
Write-Host "  Diferencia exacta: $diffBytes bytes ($([math]::Round([math]::Abs($diffBytes)/1MB,2)) MB)"
Write-Host "  Porcentaje migrado: $percentMigrated%" -ForegroundColor $(if ($percentMigrated -ge 99) { 'Green' } elseif ($percentMigrated -ge 95) { 'Yellow' } else { 'Red' })
Write-Host ""

if ([math]::Abs($diffFiles) -eq 0) {
    Write-Host "  ✓ Todos los archivos migrados correctamente" -ForegroundColor Green
} elseif ($diffFiles -gt 0) {
    Write-Host "  ⚠ Archivos faltantes: $diffFiles" -ForegroundColor Yellow
} else {
    Write-Host "  ℹ Archivos extra en destino: $([math]::Abs($diffFiles))" -ForegroundColor Cyan
}

Write-Host "  Diferencia de tamaño: $([math]::Round([math]::Abs($diffBytes)/1GB,2)) GB"
if ([math]::Abs($diffBytes) -lt 100MB) {
    Write-Host "  ✓ Tamaños coinciden (≤100 MB diferencia)" -ForegroundColor Green
} elseif ($diffBytes -gt 0) {
    Write-Host "  ⚠ Faltan $([math]::Round($diffBytes/1GB,2)) GB por migrar" -ForegroundColor Yellow
} else {
    Write-Host "  ℹ Destino tiene $([math]::Round([math]::Abs($diffBytes)/1GB,2)) GB extra" -ForegroundColor Cyan
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

# GENERAR COMPARATIVA CSV (formato limpio para gerencia, sin datos técnicos)
$logPath    = Join-Path $PSScriptRoot ($LOG_DIRECTORY)
if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath | Out-Null }
$scopeLabel  = if ($ScopeFolder -ne "") { $ScopeFolder -replace '[/\\]', '-' } else { "full" }
$compareFile = Join-Path $logPath "storage-comparison-$($UserEmail.Split('@')[0])-$scopeLabel-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"

$lines = @(
    "STORAGE COMPARISON"
    "Date;$((Get-Date).ToString('dd/MM/yyyy HH:mm'))"
    "User;$UserEmail"
    "Migrated folder;$(if ($ScopeFolder -ne '') { $ScopeFolder } else { 'Full' })"
    ""
    "SOURCE (OneDrive)"
    "Total source files;$toMigrateFileCount"
    ""
    "DESTINATION (SharePoint)"
    "Total migrated files;$fileCountDestino"
    "Completion percentage;$percentMigrated%"
    ""
    "RESULT;$resultadoTexto"
)
$lines | Out-File $compareFile -Encoding UTF8
Write-Host ""
Write-Host "  Comparativa generada: $compareFile" -ForegroundColor DarkGray
Write-Host ""
