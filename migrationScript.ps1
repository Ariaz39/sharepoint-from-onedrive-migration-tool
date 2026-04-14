# ==============================================================================
# SCRIPT DE MIGRACIÓN ONEDRIVE A SHAREPOINT - VERSIÓN OPTIMIZADA (V5)
#
# Autor: Alejandro Ariaz (@Ariaz39)
# Licencia: MIT License
# Repositorio: https://github.com/Ariaz39/sharepoint-from-onedrive-migration-tool
#
# Requiere: PowerShell 7+ y PnP.PowerShell
# Novedades V5: Escaneo paginado en streaming, checkpoint por archivo (reanudable),
#               pausa entre lotes configurable, apto para migraciones de 100GB+
# ==============================================================================

param (
    [switch]$Force,
    [switch]$SkipExisting,
    [string]$ScopeFolder = ""  # Ej: -ScopeFolder "Proyectos/2026" o múltiples: -ScopeFolder "Carpeta1,Carpeta2"
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
$firstUser = @($USERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })[0]
$firstUserName = $firstUser.Split('@')[0]
$reportFile = Join-Path $logPath "migration-report-v4-$firstUserName-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"

if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath | Out-Null }
"Timestamp;Step;SourcePath;DestPath;Status;Message" | Out-File $reportFile -Encoding UTF8

# Convertir variables del .env a Arrays
# IMPORTANTE: usar @() para forzar array y evitar problemas de indexación
$userList = @($USERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
$excludedList      = @($EXCLUDED_FOLDERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
$excludedExtensions = @($EXCLUDED_EXTENSIONS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

# Variables de Control de Rendimiento (leídas desde .env)
$throttleLimit  = [int]$THROTTLE_LIMIT
$batchSize      = [int]$BATCH_SIZE
$maxRetries     = [int]$MAX_RETRIES
$pageSize       = [int]$PAGE_SIZE
$batchPauseSec  = if ($BATCH_PAUSE_SECONDS) { [int]$BATCH_PAUSE_SECONDS } else { 0 }

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "MIGRACIÓN ONEDRIVE → SHAREPOINT V5" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ThrottleLimit: $throttleLimit | BatchSize: $batchSize | PageSize: $pageSize | PauseLote: ${batchPauseSec}s" -ForegroundColor DarkGray
# Convertir ScopeFolder en array (soporta múltiples carpetas separadas por coma)
$scopeFolderList = @($ScopeFolder -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
if ($scopeFolderList.Count -gt 0) { Write-Host "[SCOPE ACTIVADO] Solo se migrará: $($scopeFolderList -join ' | ')" -ForegroundColor Yellow }
Write-Host ""

# 2. Función de Sanitización de Nombres
function Format-SafeName {
    param([string]$Name)
    # Reemplazar caracteres prohibidos o problemáticos en SharePoint/URLs
    # Prohibidos estrictos: " * : < > ? | \
    # Problemáticos en URLs/rutas: # % & @ ! $ ' ~ { }
    $sanitized = $Name -replace '["\*:<>?|\\#%&@!$''~{}]', '_'
    # Colapsar espacios múltiples y limpiar extremos
    $sanitized = ($sanitized -replace '\s+', ' ').Trim()
    return $sanitized
}

# 3. Función de Transferencia Optimizada
function Transfer-File {
    param(
        $SourceUrl, $DestSiteUrl,
        $DestFolderPath, $FileName, $UserEmail,
        $CertPath, $CertPasswordPlain, $TenantId, $ClientId,
        $MaxRetries, $RetryDelay429, $RetryDelayDefault
    )

    $retryCount = 0
    $success = $false
    $tempFile = $null
    $sourceConn = $null
    $destConn = $null

    try {
        # Convertir contraseña a SecureString dentro del hilo (evita corrupción de SecureString entre runspaces)
        $CertPassword = ConvertTo-SecureString -String $CertPasswordPlain -AsPlainText -Force

        while (-not $success -and $retryCount -lt $MaxRetries) {
            $tempFile = $null
            try {
                # Reconectar en cada intento: evita "Nullable object must have a value"
                # que ocurre esporádicamente cuando el runspace no inicializa la conexión.
                # Solo impacta hilos con fallo (no los exitosos), delay previo evita presión al API.
                if ($retryCount -gt 0) { Start-Sleep -Seconds $RetryDelayDefault }
                $sourceConn = Connect-PnPOnline -Url $SourceUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertPath -CertificatePassword $CertPassword -ReturnConnection -ErrorAction Stop
                $destConn   = Connect-PnPOnline -Url $DestSiteUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertPath -CertificatePassword $CertPassword -ReturnConnection -ErrorAction Stop

                # Extraer el nombre original del archivo y sanitizarlo
                $rawFileName    = $FileName.Substring($FileName.LastIndexOf('/') + 1)
                $safeFileName   = Format-SafeName -Name $rawFileName
                $fileRenamed    = $safeFileName -ne $rawFileName

                # Detectar qué segmentos de carpeta fueron sanitizados
                $rawFolderSegments      = $FileName.Substring(0, $FileName.LastIndexOf('/')).Split('/')
                $renamedFolderDetails   = $rawFolderSegments | Where-Object { $_ -ne "" } | ForEach-Object {
                    $safe = Format-SafeName -Name $_
                    if ($safe -ne $_) { "carpeta: '$_' -> '$safe'" }
                }

                # Crear archivo temporal
                $tempFile = [System.IO.Path]::GetTempFileName()

                # Descargar del origen
                Get-PnPFile -Url $FileName -Path ([System.IO.Path]::GetDirectoryName($tempFile)) -FileName ([System.IO.Path]::GetFileName($tempFile)) -Connection $sourceConn -AsFile -Force -ErrorAction Stop

                # Subir al destino con el nombre sanitizado
                Add-PnPFile -Path $tempFile -Folder $DestFolderPath -NewFileName $safeFileName -Connection $destConn -ErrorAction Stop | Out-Null

                $success = $true
                $destPath = "$DestFolderPath/$safeFileName"
                $reasons  = @()
                if ($fileRenamed)                        { $reasons += "archivo: '$rawFileName' -> '$safeFileName'" }
                if ($renamedFolderDetails.Count -gt 0)  { $reasons += $renamedFolderDetails }
                $message = if ($reasons.Count -gt 0) { "Copiado (sanitizado - $($reasons -join '; '))" } else { "Copiado" }
                return [PSCustomObject]@{ Status = "OK"; Item = $FileName; DestPath = $destPath; User = $UserEmail; Message = $message }
            } catch {
                $retryCount++
                # Retry con backoff exponencial
                if ($_.Exception.Message -match "429") {
                    Start-Sleep -Seconds ($retryCount * $RetryDelay429)
                } else {
                    Start-Sleep -Seconds $RetryDelayDefault
                }

                if ($retryCount -ge $MaxRetries) {
                    return [PSCustomObject]@{ Status = "ERROR"; Item = $FileName; DestPath = ""; User = $UserEmail; Message = $_.Exception.Message }
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
    
    $upnPrefix = $email.ToLower().Replace(".", "_").Replace("@", "_")
    $myUrlBase = $MY_URL -replace '/personal$', ''
    $oneDriveUrl = "$myUrlBase/personal/$upnPrefix"
    $baseDocUrl = "/personal/$upnPrefix/Documents"

    # Rutas de destino
    $destSiteName = $DESTINATION_SITE_URL.Split('/')[-1]
    $destBaseUrl = "/sites/$destSiteName/$DESTINATION_LIBRARY"  # Para comparación/logging
    $destLibraryRelative = $DESTINATION_LIBRARY  # Ruta relativa a la biblioteca (para operaciones PnP)

    # Archivo de checkpoint: registra archivos ya migrados exitosamente para permitir reanudar
    $checkpointFile = Join-Path $logPath "checkpoint-$firstUserName.json"
    $checkpoint = @{}
    if (Test-Path $checkpointFile) {
        Write-Host "  [CHECKPOINT] Cargando progreso previo desde $checkpointFile..." -ForegroundColor Yellow
        $checkpoint = Get-Content $checkpointFile -Raw | ConvertFrom-Json -AsHashtable
        Write-Host "  [CHECKPOINT] $($checkpoint.Count) archivos ya migrados (se saltarán si no tienen cambios mayores a 20 min)." -ForegroundColor Yellow
    }

    try {
        Write-Host "  Conectando a SharePoint..." -ForegroundColor DarkGray
        $sourceConn = Connect-PnPOnline -Url $oneDriveUrl -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop
        $destConn   = Connect-PnPOnline -Url $DESTINATION_SITE_URL -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop

        # Obtener definición de funciones para uso paralelo
        $transferFunctionDef  = ${function:Transfer-File}.ToString()
        $formatSafeNameDef    = ${function:Format-SafeName}.ToString()

        # Caché de carpetas ya creadas en destino (compartida entre scopes)
        $createdFolders = @{}

        $migrationStartTime = Get-Date
        $processedFiles     = 0
        $successCount       = 0
        $errorCount         = 0
        $skippedCount       = 0
        $currentBatchNum    = 0

        # Determinar qué carpetas escanear (scope o raíz completa)
        $scanFolders = if ($scopeFolderList.Count -gt 0) {
            $scopeFolderList | ForEach-Object { "$baseDocUrl/$_" }
        } else {
            @($baseDocUrl)
        }

        # Expande el árbol de carpetas nivel por nivel usando el objeto Folder de PnP
        # Retorna lista plana de ServerRelativeUrl de todas las subcarpetas
        function Get-AllSubFolders {
            param([string]$FolderServerRelUrl, $Conn)
            $queue  = [System.Collections.Generic.Queue[string]]::new()
            $result = [System.Collections.Generic.List[string]]::new()
            $queue.Enqueue($FolderServerRelUrl)
            while ($queue.Count -gt 0) {
                $current = $queue.Dequeue()
                $result.Add($current)
                try {
                    # Obtener el objeto Folder por su ServerRelativeUrl completa
                    $folder = Get-PnPFolder -Url $current -Connection $Conn -ErrorAction Stop
                    $folder.Context.Load($folder.Folders)
                    $folder.Context.ExecuteQuery()
                    foreach ($sf in $folder.Folders) {
                        # Ignorar carpetas de sistema de SharePoint
                        if ($sf.Name -notin @('Forms', '_vti_cnf')) {
                            $queue.Enqueue($sf.ServerRelativeUrl)
                        }
                    }
                } catch { <# ignorar errores de permiso en subcarpetas individuales #> }
            }
            return ,$result
        }

        foreach ($scanFolder in $scanFolders) {
            Write-Host "  Expandiendo árbol de carpetas: $scanFolder" -ForegroundColor DarkGray
            $allSubFolders = Get-AllSubFolders -FolderServerRelUrl $scanFolder -Conn $sourceConn
            Write-Host "  -> $($allSubFolders.Count) carpetas encontradas (incluyendo raíz)" -ForegroundColor DarkGray

            foreach ($subFolder in $allSubFolders) {
                Write-Host "  Escaneando: $subFolder" -ForegroundColor DarkGray

                # Escaneo PAGINADO con CAML FilesOnly: solo archivos directos de esta carpeta,
                # sin recursión → nunca supera el threshold de 5000 por consulta
                $pageNumber = 0
                $position   = $null
                # CAML Query con Scope FilesOnly: devuelve solo ítems de tipo File en la carpeta actual
                $camlQuery  = "<View Scope='FilesOnly'><RowLimit Paged='TRUE'>$pageSize</RowLimit></View>"

                do {
                    $pageNumber++
                    try {
                        $pageArgs = @{
                            List                    = "Documents"
                            Query                   = $camlQuery
                            Connection              = $sourceConn
                            FolderServerRelativeUrl = $subFolder
                            PageSize                = $pageSize
                            ErrorAction             = "Stop"
                        }
                        if ($position) { $pageArgs["ListItemCollectionPosition"] = $position }

                        $page     = Get-PnPListItem @pageArgs
                        $position = $page.ListItemCollectionPosition

                        $pageFiles = @($page | Where-Object { $_.FileSystemObjectType -eq "File" })
                        if ($pageFiles.Count -gt 0) {
                            Write-Host "    Pág $pageNumber`: $($pageFiles.Count) archivos" -ForegroundColor DarkGray
                        }

                        # Construir lote desde esta página y procesarlo
                        $batchItems = [System.Collections.Generic.List[object]]::new()

                        foreach ($file in $pageFiles) {
                            $sourceRelUrl = $file.FieldValues.FileRef

                            # FILTRO: Checkpoint (ya migrado exitosamente)
                            # Si fue migrado pero el origen se modificó hace más de 20 minutos
                            # respecto a la fecha de migración, se sobreescribe para no perder cambios
                            if ($checkpoint.ContainsKey($sourceRelUrl)) {
                                $migratedAt   = [datetime]::Parse($checkpoint[$sourceRelUrl], [System.Globalization.CultureInfo]::InvariantCulture)
                                $sourceModified = $file.FieldValues.Modified
                                if (($sourceModified - $migratedAt).TotalMinutes -gt 20) {
                                    Write-Host "    [ACTUALIZAR] $sourceRelUrl (modificado $([math]::Round(($sourceModified - $migratedAt).TotalMinutes,1)) min después de migrar)" -ForegroundColor Yellow
                                    # No hacer continue → cae al bloque de migración para sobreescribir
                                } else {
                                    $skippedCount++
                                    continue
                                }
                            }

                            # FILTRO: Extensiones excluidas (locks de ArcGIS, temporales, etc.)
                            if ($excludedExtensions.Count -gt 0) {
                                $fileName = $sourceRelUrl.Substring($sourceRelUrl.LastIndexOf('/') + 1)
                                $isLockFile = $excludedExtensions | Where-Object { $fileName.EndsWith($_, [System.StringComparison]::InvariantCultureIgnoreCase) }
                                if ($isLockFile) {
                                    $logLine = "$((Get-Date).ToString('yyyy-MM-dd_HH:mm:ss'));Scan;$sourceRelUrl;;SKIPPED;Extensión excluida ($fileName)"
                                    Add-Content -Path $reportFile -Value $logLine
                                    $skippedCount++
                                    continue
                                }
                            }

                            # FILTRO: Carpetas Excluidas
                            $pathParts      = $sourceRelUrl.Replace($baseDocUrl, "").Trim('/').Split('/')
                            $directoryParts = if ($pathParts.Count -gt 1) { $pathParts[0..($pathParts.Count - 2)] } else { @() }
                            $isExcluded     = $false
                            foreach ($excl in $excludedList) {
                                if ($directoryParts -contains $excl) {
                                    $isExcluded = $true
                                    $logLine = "$((Get-Date).ToString('yyyy-MM-dd_HH:mm:ss'));Scan;$sourceRelUrl;;SKIPPED;Carpeta Excluida ($excl)"
                                    Add-Content -Path $reportFile -Value $logLine
                                    break
                                }
                            }
                            if ($isExcluded) { $skippedCount++; continue }

                            # Construir ruta de destino sanitizada
                            $relativePathFromDocs = $sourceRelUrl -replace [regex]::Escape($baseDocUrl), ""
                            $relativePathFromDocs = $relativePathFromDocs.TrimStart('/')
                            $destFolderRelative   = if ($relativePathFromDocs.Contains('/')) {
                                $folderSegments = $relativePathFromDocs.Substring(0, $relativePathFromDocs.LastIndexOf('/')).Split('/') |
                                    ForEach-Object { Format-SafeName -Name $_ }
                                $destLibraryRelative + "/" + ($folderSegments -join '/')
                            } else {
                                $destLibraryRelative
                            }

                            # Pre-crear carpeta de destino si no existe aún
                            if (-not $createdFolders.ContainsKey($destFolderRelative)) {
                                $parts     = $destFolderRelative.Trim('/').Split('/')
                                $pathBuild = @()
                                foreach ($part in $parts) {
                                    if ($part) {
                                        $pathBuild += $part
                                        $currentPath = $pathBuild -join '/'
                                        if (-not $createdFolders.ContainsKey($currentPath)) {
                                            $parentPath = if ($pathBuild.Count -gt 1) { $pathBuild[0..($pathBuild.Count - 2)] -join '/' } else { "" }
                                            try { Add-PnPFolder -Name $part -Folder $parentPath -Connection $destConn -ErrorAction Stop | Out-Null } catch {}
                                            $createdFolders[$currentPath] = $true
                                        }
                                    }
                                }
                            }

                            $batchItems.Add([PSCustomObject]@{
                                SourceUrl     = $oneDriveUrl
                                SourceFileRef = $sourceRelUrl
                                DestFolder    = $destFolderRelative
                                User          = $email
                            })
                        }

                        # Procesar batchItems en sub-lotes paralelos del tamaño configurado
                        $allItems = @($batchItems)
                        for ($i = 0; $i -lt $allItems.Count; $i += $batchSize) {
                            $subBatch = $allItems | Select-Object -Skip $i -First $batchSize
                            if (-not $subBatch) { continue }

                            $currentBatchNum++
                            $elapsed = (Get-Date) - $migrationStartTime
                            $rate    = if ($processedFiles -gt 0 -and $elapsed.TotalSeconds -gt 0) { $processedFiles / $elapsed.TotalMinutes } else { 0 }

                            Write-Host "  -> Lote $currentBatchNum ($($subBatch.Count) arch) | ✓ $successCount ✗ $errorCount ⏭ $skippedCount | $(if ($rate -gt 0) { "$([math]::Round($rate,1)) arch/min" } else { '...' })" -ForegroundColor DarkGray

                            $results = $subBatch | ForEach-Object -Parallel {
                                ${function:Format-SafeName} = $using:formatSafeNameDef
                                ${function:Transfer-File}   = $using:transferFunctionDef

                                Transfer-File -SourceUrl          $_.SourceUrl `
                                              -DestSiteUrl        $using:DESTINATION_SITE_URL `
                                              -DestFolderPath     $_.DestFolder `
                                              -FileName           $_.SourceFileRef `
                                              -UserEmail          $_.User `
                                              -CertPath           $using:certPath `
                                              -CertPasswordPlain  $using:CERT_PASSWORD `
                                              -TenantId           $using:TENANT `
                                              -ClientId           $using:CLIENT_ID `
                                              -MaxRetries         $using:maxRetries `
                                              -RetryDelay429      ([int]$using:RETRY_DELAY_429) `
                                              -RetryDelayDefault  ([int]$using:RETRY_DELAY_DEFAULT)
                            } -ThrottleLimit $throttleLimit

                            # Guardar resultados y actualizar checkpoint
                            foreach ($res in $results) {
                                $logLine = "$((Get-Date).ToString('yyyy-MM-dd_HH:mm:ss'));Migrate;$($res.Item);$($res.DestPath);$($res.Status);$($res.Message)"
                                Add-Content -Path $reportFile -Value $logLine
                                $processedFiles++
                                if ($res.Status -eq "ERROR") {
                                    $errorCount++
                                    Write-Host "    [ERROR] $($res.Item)" -ForegroundColor Red
                                } else {
                                    $successCount++
                                    $checkpoint[$res.Item] = (Get-Date).ToString('o')
                                }
                            }

                            # Persistir checkpoint en disco tras cada sub-lote
                            $checkpoint | ConvertTo-Json -Compress | Set-Content $checkpointFile -Encoding UTF8

                            # Pausa entre lotes para evitar throttling sostenido en migraciones grandes
                            if ($batchPauseSec -gt 0) {
                                Write-Host "    Pausa: ${batchPauseSec}s..." -ForegroundColor DarkGray
                                Start-Sleep -Seconds $batchPauseSec
                            }

                            # Liberar RAM
                            $subBatch = $null
                            $results  = $null
                            [System.GC]::Collect()
                            [System.GC]::WaitForPendingFinalizers()
                        }

                    } catch {
                        Write-Host "  [WARN] Error en página $pageNumber de '$subFolder': $($_.Exception.Message)" -ForegroundColor Yellow
                        $position = $null  # Detener paginación de esta carpeta si falla
                    }

                } while ($position)  # Continuar mientras haya más páginas de esta subcarpeta
            }  # fin foreach ($subFolder)
        }  # fin foreach ($scanFolder)

        # Resumen de migración
        $totalTime = (Get-Date) - $migrationStartTime
        $avgRate   = if ($totalTime.TotalMinutes -gt 0) { [math]::Round($processedFiles / $totalTime.TotalMinutes, 1) } else { 0 }

        Write-Host ""
        Write-Host "  ========================================" -ForegroundColor Cyan
        Write-Host "  RESUMEN DE MIGRACIÓN - $email" -ForegroundColor Cyan
        Write-Host "  ========================================" -ForegroundColor Cyan
        Write-Host "  Archivos procesados: $processedFiles" -ForegroundColor White
        Write-Host "  Exitosos:            $successCount"   -ForegroundColor Green
        Write-Host "  Errores:             $errorCount"     -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host "  Saltados (checkpoint/excluidos): $skippedCount" -ForegroundColor DarkGray
        Write-Host "  Tiempo total:        $($totalTime.Hours)h $($totalTime.Minutes)m $($totalTime.Seconds)s" -ForegroundColor White
        Write-Host "  Velocidad promedio:  $avgRate archivos/min" -ForegroundColor White
        Write-Host "  Checkpoint guardado: $checkpointFile" -ForegroundColor DarkGray
        Write-Host "  ========================================" -ForegroundColor Cyan

        Add-Content -Path $reportFile -Value ""
        Add-Content -Path $reportFile -Value ";;RESUMEN DE MIGRACIÓN - $email;;;"
        Add-Content -Path $reportFile -Value "Archivos procesados;$processedFiles;;;;"
        Add-Content -Path $reportFile -Value "Exitosos;$successCount;;;;"
        Add-Content -Path $reportFile -Value "Errores;$errorCount;;;;"
        Add-Content -Path $reportFile -Value "Saltados;$skippedCount;;;;"
        Add-Content -Path $reportFile -Value "Tiempo total;$($totalTime.Hours)h $($totalTime.Minutes)m $($totalTime.Seconds)s;;;;"
        Add-Content -Path $reportFile -Value "Velocidad promedio;$avgRate archivos/min;;;;"

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
