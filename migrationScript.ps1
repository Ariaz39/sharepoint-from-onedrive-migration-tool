# ==============================================================================
# SCRIPT DE MIGRACIÓN ONEDRIVE A SHAREPOINT - VERSIÓN OPTIMIZADA (V4)
#
# Autor: Alejandro Ariaz (@Ariaz39)
# Licencia: MIT License
# Repositorio: https://github.com/Ariaz39/sharepoint-from-onedrive-migration-tool
#
# Requiere: PowerShell 7+ y PnP.PowerShell
# Novedades V4: Paralelismo aumentado (25 hilos), lotes 1000, paginación configurable
# Optimizado para: 1 millón de archivos (versión conservadora)
# ==============================================================================

param (
    [switch]$Force,
    [switch]$SkipExisting,
    [string]$ScopeFolder = ""  # Ej: -ScopeFolder "Proyectos/2026"
)

# 1. Cargar Variables de Entorno
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

$certPath = Join-Path $PSScriptRoot "PnPMigrationCert.pfx"
$certPassword = ConvertTo-SecureString -String $CERT_PASSWORD -AsPlainText -Force
$logPath = Join-Path $PSScriptRoot "Logs"
$reportFile = Join-Path $logPath "migration-report-v4-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"

if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath | Out-Null }
"Timestamp,User,Step,Item,Status,Message" | Out-File $reportFile -Encoding UTF8

# Convertir variables del .env a Arrays
# IMPORTANTE: usar @() para forzar array y evitar problemas de indexación
$userList = @($USERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
$excludedList = @($EXCLUDED_FOLDERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

# Variables de Control de Rendimiento (leídas desde .env)
$throttleLimit = [int]$THROTTLE_LIMIT
$batchSize = [int]$BATCH_SIZE
$maxRetries = [int]$MAX_RETRIES
$pageSize = [int]$PAGE_SIZE

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "MIGRACIÓN ONEDRIVE → SHAREPOINT V4" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ThrottleLimit: $throttleLimit | BatchSize: $batchSize | PageSize: $pageSize" -ForegroundColor DarkGray
if ($ScopeFolder) { Write-Host "[SCOPE ACTIVADO] Solo se migrará: $ScopeFolder" -ForegroundColor Yellow }
Write-Host ""

# 2. Función de Sanitización de Nombres
function Format-SafeName {
    param([string]$Name)
    # Reemplazar caracteres prohibidos por SharePoint
    $sanitized = $Name -replace '["\*:<>?|\\]', '_'
    # Colapsar espacios múltiples y limpiar extremos
    $sanitized = ($sanitized -replace '\s+', ' ').Trim()
    return $sanitized
}

# 3. Función de Transferencia Optimizada
function Transfer-File {
    param(
        $SourceUrl, $DestSiteUrl,
        $DestFolderPath, $FileName, $UserEmail,
        $CertPath, $CertPassword, $TenantId, $ClientId,
        $MaxRetries, $RetryDelay429, $RetryDelayDefault
    )

    $retryCount = 0
    $success = $false
    $tempFile = $null
    $sourceConn = $null
    $destConn = $null

    try {
        # Crear conexiones (una por archivo, evita problemas de thread-safety)
        $sourceConn = Connect-PnPOnline -Url $SourceUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertPath -CertificatePassword $CertPassword -ReturnConnection -ErrorAction Stop
        $destConn = Connect-PnPOnline -Url $DestSiteUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertPath -CertificatePassword $CertPassword -ReturnConnection -ErrorAction Stop

        while (-not $success -and $retryCount -lt $MaxRetries) {
            $tempFile = $null
            try {
                # Extraer el nombre original del archivo y sanitizarlo
                $rawFileName = $FileName.Substring($FileName.LastIndexOf('/') + 1)
                $originalFileName = Format-SafeName -Name $rawFileName
                $nameSanitized = $originalFileName -ne $rawFileName

                # Crear archivo temporal
                $tempFile = [System.IO.Path]::GetTempFileName()

                # Descargar del origen
                Get-PnPFile -Url $FileName -Path ([System.IO.Path]::GetDirectoryName($tempFile)) -FileName ([System.IO.Path]::GetFileName($tempFile)) -Connection $sourceConn -AsFile -Force -ErrorAction Stop

                # Subir al destino con el nombre sanitizado
                Add-PnPFile -Path $tempFile -Folder $DestFolderPath -NewFileName $originalFileName -Connection $destConn -ErrorAction Stop | Out-Null

                $success = $true
                $message = if ($nameSanitized) { "Copiado (nombre sanitizado: '$rawFileName' -> '$originalFileName')" } else { "Copiado" }
                return [PSCustomObject]@{ Status = "OK"; Item = $FileName; User = $UserEmail; Message = $message }
            } catch {
                $retryCount++
                # Retry con backoff exponencial
                if ($_.Exception.Message -match "429") {
                    Start-Sleep -Seconds ($retryCount * $RetryDelay429)
                } else {
                    Start-Sleep -Seconds $RetryDelayDefault
                }

                if ($retryCount -ge $MaxRetries) {
                    return [PSCustomObject]@{ Status = "ERROR"; Item = $FileName; User = $UserEmail; Message = $_.Exception.Message }
                }
            } finally {
                # CRÍTICO: Limpiar archivo temporal SIEMPRE
                if ($tempFile -and (Test-Path $tempFile)) {
                    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                }
            }
        }
    } finally {
        # Las conexiones se limpian automáticamente al salir del scope
        $sourceConn = $null
        $destConn = $null
    }
}

# 3. Lógica Principal Iterativa
foreach ($email in $userList) {
    Write-Host "`nProcesando usuario: $email" -ForegroundColor Yellow
    
    $upnPrefix = $email.Replace(".", "_").Replace("@", "_")
    $myUrlBase = $MY_URL -replace '/personal$', ''
    $oneDriveUrl = "$myUrlBase/personal/$upnPrefix"
    $baseDocUrl = "/personal/$upnPrefix/Documents"

    # Rutas de destino
    $destSiteName = $DESTINATION_SITE_URL.Split('/')[-1]
    $destBaseUrl = "/sites/$destSiteName/$DESTINATION_LIBRARY"  # Para comparación/logging
    $destLibraryRelative = $DESTINATION_LIBRARY  # Ruta relativa a la biblioteca (para operaciones PnP)

    try {
        Write-Host "  Conectando a SharePoint..." -ForegroundColor DarkGray
        $sourceConn = Connect-PnPOnline -Url $oneDriveUrl -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop
        $destConn = Connect-PnPOnline -Url $DESTINATION_SITE_URL -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop
        
        Write-Host "  Escaneando archivos..." -ForegroundColor DarkGray
        $allSourceFiles = Get-PnPListItem -List "Documents" -PageSize $pageSize -Connection $sourceConn | Where-Object { $_.FileSystemObjectType -eq "File" }

        # Caché de destino
        $destCache = @{}
        if ($SkipExisting) {
            Write-Host "  Verificando archivos existentes..." -ForegroundColor DarkGray
            $allDestFiles = Get-PnPListItem -List $DESTINATION_LIBRARY -PageSize $pageSize -Connection $destConn | Where-Object { $_.FileSystemObjectType -eq "File" }
            foreach ($df in $allDestFiles) { $destCache[$df.FieldValues.FileRef] = $true }
        }

        # Aplicar Filtros: Scope y Exclusiones
        $filesToProcess = @()
        foreach ($file in $allSourceFiles) {
            $sourceRelUrl = $file.FieldValues.FileRef
            
            # FILTRO 1: Scope de Carpeta
            if ($ScopeFolder -ne "") {
                $scopePath = "$baseDocUrl/$ScopeFolder"
                # Si la ruta del archivo no empieza con la ruta del scope, lo saltamos
                if (-not $sourceRelUrl.StartsWith($scopePath, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                    continue
                }
            }

            # FILTRO 2: Carpetas Excluidas
            $isExcluded = $false
            # Extraer solo los nombres de las carpetas de la ruta del archivo
            $pathParts = $sourceRelUrl.Replace($baseDocUrl, "").Trim('/').Split('/')
            $directoryParts = if ($pathParts.Count -gt 1) { $pathParts[0..($pathParts.Count - 2)] } else { @() }
            
            foreach ($excl in $excludedList) {
                if ($directoryParts -contains $excl) {
                    $isExcluded = $true
                    $logLine = "$((Get-Date).ToString('o')),$email,Scan,$sourceRelUrl,SKIPPED,System folder excluded ($excl)"
                    Add-Content -Path $reportFile -Value $logLine
                    break
                }
            }
            if ($isExcluded) { continue }
            
            # FILTRO 3: Archivos Existentes
            $expectedDestPath = $sourceRelUrl -replace [regex]::Escape($baseDocUrl), $destBaseUrl
            if ($SkipExisting -and $destCache.ContainsKey($expectedDestPath)) {
                $logLine = "$((Get-Date).ToString('o')),$email,Migrate,$sourceRelUrl,SKIPPED,File exists in destination"
                Add-Content -Path $reportFile -Value $logLine
                continue
            }

            # Construir ruta RELATIVA a la biblioteca (sin /sites/.../library)
            # Sanitizar cada segmento de carpeta individualmente
            $relativePathFromDocs = $sourceRelUrl -replace [regex]::Escape($baseDocUrl), ""
            $relativePathFromDocs = $relativePathFromDocs.TrimStart('/')
            $destFolderRelative = if ($relativePathFromDocs.Contains('/')) {
                $folderSegments = $relativePathFromDocs.Substring(0, $relativePathFromDocs.LastIndexOf('/')).Split('/') | ForEach-Object { Format-SafeName -Name $_ }
                $destLibraryRelative + "/" + ($folderSegments -join '/')
            } else {
                $destLibraryRelative
            }

            $filesToProcess += [PSCustomObject]@{
                SourceUrl = $oneDriveUrl
                SourceFileRef = $sourceRelUrl
                DestFolder = $destFolderRelative  # Ahora es relativo a la biblioteca
                User = $email
            }
        }

        $totalFiles = $filesToProcess.Count
        Write-Host "  Archivos a migrar tras aplicar filtros: $totalFiles" -ForegroundColor Cyan

        # CRÍTICO: Pre-crear todas las carpetas de destino (OPTIMIZADO con caché)
        $uniqueFolders = $filesToProcess | Select-Object -ExpandProperty DestFolder -Unique | Sort-Object
        Write-Host "  Creando estructura de carpetas ($($uniqueFolders.Count) rutas únicas)..." -ForegroundColor DarkGray

        # Caché de carpetas ya creadas (evita verificar/crear la misma carpeta múltiples veces)
        $createdFolders = @{}
        $folderCount = 0
        $startTime = Get-Date

        foreach ($folder in $uniqueFolders) {
            try {
                # Mostrar progreso cada 50 carpetas
                $folderCount++
                if ($folderCount % 50 -eq 0) {
                    $elapsed = ((Get-Date) - $startTime).TotalSeconds
                    $rate = $folderCount / $elapsed
                    $remaining = ($uniqueFolders.Count - $folderCount) / $rate
                    Write-Host "    Progreso: $folderCount/$($uniqueFolders.Count) carpetas (${remaining}s restantes aprox.)" -ForegroundColor DarkGray
                }

                # Crear carpeta recursivamente parte por parte
                $parts = $folder.Trim('/').Split('/')
                $pathParts = @()

                foreach ($part in $parts) {
                    if ($part) {
                        $pathParts += $part
                        $currentPath = $pathParts -join '/'

                        # OPTIMIZACIÓN: Skip si ya creamos esta carpeta en esta sesión
                        if ($createdFolders.ContainsKey($currentPath)) {
                            continue
                        }

                        try {
                            # Intentar crear directamente (más rápido que verificar primero)
                            $parentPath = if ($pathParts.Count -gt 1) {
                                $pathParts[0..($pathParts.Count - 2)] -join '/'
                            } else {
                                ""
                            }
                            Add-PnPFolder -Name $part -Folder $parentPath -Connection $destConn -ErrorAction Stop | Out-Null
                            $createdFolders[$currentPath] = $true
                        } catch {
                            # Si falla (probablemente ya existe), marcarlo como creado de todos modos
                            $createdFolders[$currentPath] = $true
                        }
                    }
                }
            } catch {
                # Ignorar errores generales
            }
        }

        $totalTime = ((Get-Date) - $startTime).TotalSeconds
        Write-Host "  ✓ Estructura de carpetas creada ($folderCount rutas en ${totalTime}s)" -ForegroundColor Green

        # Obtener definición de funciones para uso paralelo
        $transferFunctionDef = ${function:Transfer-File}.ToString()
        $formatSafeNameDef = ${function:Format-SafeName}.ToString()

        # 4. Procesamiento en Lotes y Paralelo (Evita el colapso de RAM)
        $migrationStartTime = Get-Date
        $processedFiles = 0
        $successCount = 0
        $errorCount = 0

        for ($i = 0; $i -lt $totalFiles; $i += $batchSize) {
            $batch = $filesToProcess | Select-Object -Skip $i -First $batchSize
            $currentBatch = [math]::Floor($i/$batchSize) + 1
            $totalBatches = [math]::Ceiling($totalFiles/$batchSize)

            # Calcular progreso y ETA
            $percentComplete = [math]::Round(($processedFiles / $totalFiles) * 100, 1)
            $elapsed = (Get-Date) - $migrationStartTime
            $rate = if ($processedFiles -gt 0 -and $elapsed.TotalSeconds -gt 0) { $processedFiles / $elapsed.TotalMinutes } else { 0 }
            $remaining = if ($rate -gt 0) { ($totalFiles - $processedFiles) / $rate } else { 0 }

            # Mostrar barra de progreso
            $progressParams = @{
                Activity = "Migrando archivos de $email"
                Status = "Lote $currentBatch de $totalBatches | $processedFiles/$totalFiles archivos | ✓ $successCount ✗ $errorCount"
                PercentComplete = $percentComplete
                CurrentOperation = if ($rate -gt 0) { "Velocidad: $([math]::Round($rate, 1)) archivos/min | ETA: $([math]::Round($remaining, 1)) min" } else { "Calculando velocidad..." }
            }
            Write-Progress @progressParams

            Write-Host "  -> Lote $currentBatch/$totalBatches ($($batch.Count) archivos) | Progreso: $percentComplete% | ✓ $successCount ✗ $errorCount" -ForegroundColor DarkGray

            $results = $batch | ForEach-Object -Parallel {
                # Recrear funciones en el contexto paralelo
                ${function:Format-SafeName} = $using:formatSafeNameDef
                ${function:Transfer-File} = $using:transferFunctionDef

                Transfer-File -SourceUrl $_.SourceUrl `
                              -DestSiteUrl $using:DESTINATION_SITE_URL `
                              -DestFolderPath $_.DestFolder `
                              -FileName $_.SourceFileRef `
                              -UserEmail $_.User `
                              -CertPath $using:certPath `
                              -CertPassword $using:certPassword `
                              -TenantId $using:TENANT `
                              -ClientId $using:CLIENT_ID `
                              -MaxRetries $using:maxRetries `
                              -RetryDelay429 ([int]$using:RETRY_DELAY_429) `
                              -RetryDelayDefault ([int]$using:RETRY_DELAY_DEFAULT)
            } -ThrottleLimit $throttleLimit

            # Guardar resultados del lote
            foreach ($res in $results) {
                $logLine = "$((Get-Date).ToString('o')),$($res.User),Migrate,$($res.Item),$($res.Status),$($res.Message)"
                Add-Content -Path $reportFile -Value $logLine

                $processedFiles++
                if ($res.Status -eq "ERROR") {
                    $errorCount++
                    Write-Host "    [ERROR] $($res.Item)" -ForegroundColor Red
                } else {
                    $successCount++
                }
            }

            # LIMPIEZA DE MEMORIA RAM (Garbage Collection)
            $batch = $null
            $results = $null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        # Completar barra de progreso
        Write-Progress -Activity "Migrando archivos de $email" -Completed

        # Resumen de migración
        $totalTime = (Get-Date) - $migrationStartTime
        $avgRate = if ($totalTime.TotalMinutes -gt 0) { [math]::Round($processedFiles / $totalTime.TotalMinutes, 1) } else { 0 }

        Write-Host ""
        Write-Host "  ========================================" -ForegroundColor Cyan
        Write-Host "  RESUMEN DE MIGRACIÓN - $email" -ForegroundColor Cyan
        Write-Host "  ========================================" -ForegroundColor Cyan
        Write-Host "  Archivos procesados: $processedFiles" -ForegroundColor White
        Write-Host "  Exitosos: $successCount" -ForegroundColor Green
        Write-Host "  Errores: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host "  Tiempo total: $($totalTime.Hours)h $($totalTime.Minutes)m $($totalTime.Seconds)s" -ForegroundColor White
        Write-Host "  Velocidad promedio: $avgRate archivos/min" -ForegroundColor White
        Write-Host "  ========================================" -ForegroundColor Cyan

    } catch {
        Write-Host "  [ERROR FATAL] $email : $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nMigración finalizada. Reporte: $reportFile" -ForegroundColor Green

# Ejecutar verificación automática para cada usuario migrado
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "VERIFICACIÓN AUTOMÁTICA DE MIGRACIÓN" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

foreach ($email in $userList) {
    Write-Host "`nVerificando usuario: $email" -ForegroundColor Yellow
    & "$PSScriptRoot\check-storage.ps1" -UserEmail $email
}
