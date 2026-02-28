# ==============================================================================
# SCRIPT DE VALIDACIÓN DE NOMBRES DE ARCHIVO
# Genera reporte de archivos con nombres problemáticos (sin modificar nada)
# ==============================================================================
#
# USO:
#   .\validate-filenames.ps1                              # Usa el primer usuario del .env
#   .\validate-filenames.ps1 -UserEmail "user@domain.com" # Usa un usuario específico
#
# REPORTES GENERADOS:
#   - Logs/filename-validation-report-YYYYMMDD-HHmmss.csv
#   - Contiene rutas relativas (no filtra información de la empresa)
#
# NOTA: No hardcodea información de la empresa - lee del .env
# ==============================================================================

param(
    [string]$UserEmail = ""  # Opcional - si no se especifica, usa el primer usuario del .env
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

# Reglas de validación (configurables)
$MAX_PATH_LENGTH = 350          # Máximo caracteres en ruta (margen de 50)
$WARN_PATH_LENGTH = 300         # Advertencia si supera esto
$MAX_FILENAME_LENGTH = 200      # Máximo caracteres en nombre de archivo
$MAX_FOLDER_DEPTH = 7           # Máxima profundidad de carpetas
$WARN_FOLDER_DEPTH = 6          # Advertencia si supera esto

# Caracteres problemáticos (aunque SharePoint los permita, pueden causar problemas)
$PROBLEMATIC_CHARS = @('"', '*', ':', '<', '>', '?', '|')
$WARNING_CHARS = @('#', '%', '&', '{', '}', '~', '\', '/')

$certPath = Join-Path $PSScriptRoot "PnPMigrationCert.pfx"
$certPassword = ConvertTo-SecureString -String $CERT_PASSWORD -AsPlainText -Force
$logPath = Join-Path $PSScriptRoot "Logs"
$reportFile = Join-Path $logPath "filename-validation-report-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"

if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath | Out-Null }

# Construir URLs
$upnPrefix = $UserEmail.Replace(".", "_").Replace("@", "_")
$myUrlBase = $MY_URL -replace '/personal$', ''
$oneDriveUrl = "$myUrlBase/personal/$upnPrefix"
$baseDocUrl = "/personal/$upnPrefix/Documents"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "VALIDACIÓN DE NOMBRES DE ARCHIVO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Usuario: $UserEmail" -ForegroundColor Yellow
Write-Host "⚠️  ADVERTENCIA: Los reportes contienen rutas de archivos." -ForegroundColor Yellow
Write-Host "   NO compartir los archivos CSV fuera de la organización." -ForegroundColor DarkGray
Write-Host ""

# Inicializar contadores
$totalFiles = 0
$criticalIssues = 0
$warningIssues = 0
$infoIssues = 0
$zipFileCount = 0
$issues = @()

try {
    Write-Host "Conectando a OneDrive..." -ForegroundColor DarkGray
    $conn = Connect-PnPOnline -Url $oneDriveUrl -ClientId $CLIENT_ID -Tenant $TENANT -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop

    Write-Host "Escaneando archivos..." -ForegroundColor DarkGray
    $allFiles = Get-PnPListItem -List "Documents" -PageSize 5000 -Connection $conn | Where-Object { $_.FileSystemObjectType -eq "File" }

    $totalFiles = $allFiles.Count
    Write-Host "Archivos encontrados: $totalFiles" -ForegroundColor Cyan
    Write-Host "Analizando nombres y rutas..." -ForegroundColor DarkGray
    Write-Host ""

    $processedCount = 0
    foreach ($file in $allFiles) {
        $processedCount++
        if ($processedCount % 100 -eq 0) {
            Write-Progress -Activity "Validando archivos" -Status "$processedCount de $totalFiles" -PercentComplete (($processedCount / $totalFiles) * 100)
        }

        $filePath = $file.FieldValues.FileRef
        $fileName = $filePath.Substring($filePath.LastIndexOf('/') + 1)

        # Calcular ruta relativa desde Documents
        $relativePath = $filePath.Replace($baseDocUrl, "").TrimStart('/')

        # Calcular profundidad de carpetas
        $depth = ($relativePath.Split('/')).Count - 1  # -1 porque el archivo no cuenta

        # Calcular longitud de ruta en SharePoint (usando placeholder para no filtrar nombre de sitio)
        # Formato: /sites/[SITIO]/[BIBLIOTECA]/ruta...
        $sitePathLength = "/sites/".Length + ($DESTINATION_SITE_URL -split '/')[-1].Length + 1
        $libraryPathLength = $DESTINATION_LIBRARY.Length + 1
        $pathLength = $sitePathLength + $libraryPathLength + $relativePath.Length

        # Detectar si es archivo .zip (o .rar, .7z, .tar, .gz - archivos comprimidos)
        $isCompressedFile = $fileName -match '\.(zip|rar|7z|tar|gz|bz2|tgz)$'
        if ($isCompressedFile) { $zipFileCount++ }

        # Array de problemas encontrados en este archivo
        $fileIssues = @()
        $severity = "OK"

        # NOTA: Para archivos comprimidos (.zip, .rar, etc.) solo validamos el archivo mismo,
        # NO el contenido interno (ya que permanecerán comprimidos en SharePoint)

        # VALIDACIÓN 1: Longitud de ruta
        if ($pathLength -gt $MAX_PATH_LENGTH) {
            $fileIssues += "Ruta excede límite ($pathLength > $MAX_PATH_LENGTH caracteres)"
            $severity = "CRITICAL"
            $criticalIssues++
        } elseif ($pathLength -gt $WARN_PATH_LENGTH) {
            $fileIssues += "Ruta cerca del límite ($pathLength > $WARN_PATH_LENGTH caracteres)"
            if ($severity -ne "CRITICAL") { $severity = "WARNING" }
            $warningIssues++
        }

        # VALIDACIÓN 2: Longitud de nombre de archivo
        if ($fileName.Length -gt $MAX_FILENAME_LENGTH) {
            $fileIssues += "Nombre de archivo muy largo ($($fileName.Length) > $MAX_FILENAME_LENGTH caracteres)"
            if ($severity -ne "CRITICAL") { $severity = "WARNING" }
            $warningIssues++
        }

        # VALIDACIÓN 3: Profundidad de carpetas
        if ($depth -gt $MAX_FOLDER_DEPTH) {
            $fileIssues += "Estructura muy profunda ($depth niveles > $MAX_FOLDER_DEPTH)"
            if ($severity -ne "CRITICAL") { $severity = "WARNING" }
            $warningIssues++
        } elseif ($depth -gt $WARN_FOLDER_DEPTH) {
            $fileIssues += "Estructura profunda ($depth niveles > $WARN_FOLDER_DEPTH)"
            if ($severity -eq "OK") { $severity = "INFO" }
            $infoIssues++
        }

        # VALIDACIÓN 4: Caracteres problemáticos
        $foundProblematic = @()
        foreach ($char in $PROBLEMATIC_CHARS) {
            if ($fileName.Contains($char)) {
                $foundProblematic += $char
            }
        }
        if ($foundProblematic.Count -gt 0) {
            $fileIssues += "Caracteres prohibidos: $($foundProblematic -join ', ')"
            $severity = "CRITICAL"
            $criticalIssues++
        }

        # VALIDACIÓN 5: Caracteres de advertencia
        $foundWarning = @()
        foreach ($char in $WARNING_CHARS) {
            if ($fileName.Contains($char)) {
                $foundWarning += $char
            }
        }
        if ($foundWarning.Count -gt 0) {
            $fileIssues += "Caracteres problemáticos: $($foundWarning -join ', ')"
            if ($severity -eq "OK") { $severity = "INFO" }
            $infoIssues++
        }

        # VALIDACIÓN 6: Espacios al final del nombre
        if ($fileName.TrimEnd() -ne $fileName -or $fileName.EndsWith('.')) {
            $fileIssues += "Termina con espacio o punto (se truncará)"
            if ($severity -ne "CRITICAL") { $severity = "WARNING" }
            $warningIssues++
        }

        # VALIDACIÓN 7: Nombres de carpetas duplicadas en la ruta
        $pathParts = $relativePath.Split('/')
        $folderNames = if ($pathParts.Count -gt 1) { $pathParts[0..($pathParts.Count - 2)] } else { @() }
        $duplicates = $folderNames | Group-Object | Where-Object { $_.Count -gt 1 }
        if ($duplicates) {
            $fileIssues += "Carpetas duplicadas en ruta: $($duplicates.Name -join ', ')"
            if ($severity -eq "OK") { $severity = "INFO" }
            $infoIssues++
        }

        # VALIDACIÓN 8: Archivos comprimidos en rutas largas (solo informativo)
        if ($isCompressedFile -and $pathLength -gt 250 -and $pathLength -le $WARN_PATH_LENGTH) {
            $fileIssues += "Archivo comprimido en ruta larga ($pathLength caracteres) - Si se extrae en futuro, podría exceder límites"
            if ($severity -eq "OK") { $severity = "INFO" }
            $infoIssues++
        }

        # Solo reportar si hay problemas
        if ($fileIssues.Count -gt 0) {
            $fileType = if ($isCompressedFile) { "Comprimido" } else { "Normal" }
            $issues += [PSCustomObject]@{
                Severity = $severity
                FilePath = $relativePath
                FileName = $fileName
                FileType = $fileType
                PathLength = $pathLength
                FolderDepth = $depth
                Issues = ($fileIssues -join " | ")
            }
        }
    }

    Write-Progress -Activity "Validando archivos" -Completed

    # Generar reporte CSV
    Write-Host "Generando reporte..." -ForegroundColor DarkGray

    # Header con advertencia de seguridad
    "# ============================================================" | Out-File $reportFile -Encoding UTF8
    "# ADVERTENCIA: Este reporte contiene información sensible" | Add-Content $reportFile -Encoding UTF8
    "# NO compartir fuera de la organización" | Add-Content $reportFile -Encoding UTF8
    "# ============================================================" | Add-Content $reportFile -Encoding UTF8
    "Severidad,Ruta_Archivo,Nombre_Archivo,Tipo_Archivo,Longitud_Ruta,Profundidad_Carpetas,Problemas" | Add-Content $reportFile -Encoding UTF8

    # Ordenar por severidad (CRITICAL > WARNING > INFO)
    $sortedIssues = $issues | Sort-Object @{
        Expression = {
            switch ($_.Severity) {
                "CRITICAL" { 1 }
                "WARNING" { 2 }
                "INFO" { 3 }
            }
        }
    }, PathLength -Descending

    foreach ($issue in $sortedIssues) {
        $line = "`"$($issue.Severity)`",`"$($issue.FilePath)`",`"$($issue.FileName)`",`"$($issue.FileType)`",$($issue.PathLength),$($issue.FolderDepth),`"$($issue.Issues)`""
        Add-Content -Path $reportFile -Value $line -Encoding UTF8
    }

    # Resumen en consola
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "RESUMEN DE VALIDACIÓN" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Total de archivos analizados: $totalFiles" -ForegroundColor White
    Write-Host "Archivos comprimidos (.zip, .rar, etc.): $zipFileCount" -ForegroundColor Cyan
    Write-Host "  └─ Solo se valida ruta/nombre (contenido interno ignorado)" -ForegroundColor DarkGray
    Write-Host ""

    if ($criticalIssues -eq 0 -and $warningIssues -eq 0 -and $infoIssues -eq 0) {
        Write-Host "✓ No se encontraron problemas" -ForegroundColor Green
    } else {
        if ($criticalIssues -gt 0) {
            Write-Host "🔴 CRÍTICO: $criticalIssues archivos" -ForegroundColor Red
            Write-Host "   (No se pueden migrar o causarán errores)" -ForegroundColor Red
        }
        if ($warningIssues -gt 0) {
            Write-Host "⚠️  ADVERTENCIA: $warningIssues archivos" -ForegroundColor Yellow
            Write-Host "   (Pueden causar problemas en el futuro)" -ForegroundColor Yellow
        }
        if ($infoIssues -gt 0) {
            Write-Host "ℹ️  INFO: $infoIssues archivos" -ForegroundColor Cyan
            Write-Host "   (Buenas prácticas - considerar revisar)" -ForegroundColor Cyan
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Reporte detallado guardado en:" -ForegroundColor Green
    Write-Host $reportFile -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Mostrar top 10 problemas más críticos
    if ($criticalIssues -gt 0 -or $warningIssues -gt 0) {
        Write-Host "TOP 10 ARCHIVOS MÁS PROBLEMÁTICOS:" -ForegroundColor Yellow
        Write-Host ""
        $top10 = $sortedIssues | Select-Object -First 10
        $i = 1
        foreach ($item in $top10) {
            $color = switch ($item.Severity) {
                "CRITICAL" { 'Red' }
                "WARNING" { 'Yellow' }
                "INFO" { 'Cyan' }
            }
            Write-Host "$i. [$($item.Severity)] $($item.FileName)" -ForegroundColor $color
            Write-Host "   Ruta: $($item.FilePath)" -ForegroundColor DarkGray
            Write-Host "   Problemas: $($item.Issues)" -ForegroundColor DarkGray
            Write-Host ""
            $i++
        }
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Validación completada." -ForegroundColor Cyan
