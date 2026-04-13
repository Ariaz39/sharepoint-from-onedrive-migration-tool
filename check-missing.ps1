# ==============================================================================
# SCRIPT DE VERIFICACIÓN DE ARCHIVOS FALTANTES
# Compara OneDrive (origen) vs SharePoint (destino) archivo por archivo
# e indica cuáles no fueron migrados y el motivo probable.
#
# Autor: Alejandro Ariaz (@Ariaz39)
# Licencia: MIT License
#
# USO:
#   .\check-missing.ps1                               # Usa el primer usuario del .env
#   .\check-missing.ps1 -UserEmail "user@domain.com"  # Usuario específico
#   .\check-missing.ps1 -ScopeFolder "HSEQ/SST"       # Limitar a una carpeta
#   .\check-missing.ps1 -UseLatestLog                 # Cruza con el último log CSV
# ==============================================================================

param(
    [string]$UserEmail    = "",
    [string]$ScopeFolder  = "",
    [switch]$UseLatestLog             # Si se activa, cruza con el log más reciente para obtener el motivo exacto
)

# --- Cargar .env ---
$envFile = Join-Path $PSScriptRoot ".env"
if (-not (Test-Path $envFile)) { Write-Host "ERROR: .env no encontrado." -ForegroundColor Red; exit 1 }

Get-Content $envFile | Where-Object { $_ -match "^[^#=]+=" } | ForEach-Object {
    $name, $value = $_ -split '=', 2
    $cleanValue = ($value -split '#')[0].Trim()
    Set-Variable -Name $name.Trim() -Value $cleanValue -Scope Script
}

if ([string]::IsNullOrWhiteSpace($UserEmail)) {
    $userList  = @($USERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    $UserEmail = $userList[0]
    Write-Host "Usando usuario del .env: $UserEmail" -ForegroundColor DarkGray
}

$excludedList = @($EXCLUDED_FOLDERS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
$certPath     = Join-Path $PSScriptRoot "PnPMigrationCert.pfx"
$certPassword = ConvertTo-SecureString -String $CERT_PASSWORD -AsPlainText -Force
$logPath      = Join-Path $PSScriptRoot "Logs"

$upnPrefix   = $UserEmail.Replace(".", "_").Replace("@", "_")
$myUrlBase   = $MY_URL -replace '/personal$', ''
$oneDriveUrl = "$myUrlBase/personal/$upnPrefix"
$baseDocUrl  = "/personal/$upnPrefix/Documents"
$destSiteName    = $DESTINATION_SITE_URL.Split('/')[-1]
$destBaseUrl     = "/sites/$destSiteName/$DESTINATION_LIBRARY"

# --- Función de sanitización (igual que migrationScript.ps1) ---
function Format-SafeName {
    param([string]$Name)
    $sanitized = $Name -replace '["\*:<>?|\\#%&@!$''~{}]', '_'
    $sanitized = ($sanitized -replace '\s+', ' ').Trim()
    return $sanitized
}

# --- Cargar log CSV más reciente (opcional) ---
$logIndex = @{}   # sourceRelUrl -> motivo registrado en el log
if ($UseLatestLog) {
    $latestLog = Get-ChildItem -Path $logPath -Filter "migration-report-v4-*.csv" -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) {
        Write-Host "Usando log: $($latestLog.Name)" -ForegroundColor DarkGray
        Import-Csv -Path $latestLog.FullName -Delimiter ";" -Header "Timestamp","Step","SourcePath","DestPath","Status","Message" |
            Select-Object -Skip 1 |   # saltar cabecera duplicada
            ForEach-Object {
                if ($_.SourcePath -ne "") {
                    $logIndex[$_.SourcePath] = @{ Status = $_.Status; Message = $_.Message }
                }
            }
    } else {
        Write-Host "No se encontró ningún log CSV en Logs/." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " VERIFICACIÓN DE ARCHIVOS FALTANTES" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Usuario : $UserEmail" -ForegroundColor Yellow
if ($ScopeFolder) { Write-Host " Scope   : $ScopeFolder" -ForegroundColor Yellow }
Write-Host ""

# --- Conectar ---
Write-Host "Conectando a OneDrive..." -ForegroundColor DarkGray
try {
    $sourceConn = Connect-PnPOnline -Url $oneDriveUrl -ClientId $CLIENT_ID -Tenant $TENANT `
        -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop
} catch {
    Write-Host "ERROR conectando a OneDrive: $($_.Exception.Message)" -ForegroundColor Red; exit 1
}

Write-Host "Conectando a SharePoint..." -ForegroundColor DarkGray
try {
    $destConn = Connect-PnPOnline -Url $DESTINATION_SITE_URL -ClientId $CLIENT_ID -Tenant $TENANT `
        -CertificatePath $certPath -CertificatePassword $certPassword -ReturnConnection -ErrorAction Stop
} catch {
    Write-Host "ERROR conectando a SharePoint: $($_.Exception.Message)" -ForegroundColor Red; exit 1
}

# --- Escanear origen ---
Write-Host "Escaneando OneDrive (origen)..." -ForegroundColor DarkGray
$allSourceFiles = Get-PnPListItem -List "Documents" -PageSize 1000 -Connection $sourceConn |
                  Where-Object { $_.FileSystemObjectType -eq "File" }
Write-Host "  $($allSourceFiles.Count) archivos encontrados en OneDrive." -ForegroundColor Gray

# --- Escanear destino (construir índice por FileRef normalizado) ---
Write-Host "Escaneando SharePoint (destino)..." -ForegroundColor DarkGray
$allDestFiles = Get-PnPListItem -List $DESTINATION_LIBRARY -PageSize 1000 -Connection $destConn |
                Where-Object { $_.FileSystemObjectType -eq "File" }
Write-Host "  $($allDestFiles.Count) archivos encontrados en SharePoint." -ForegroundColor Gray

# Índice destino: clave = FileRef en minúsculas
$destIndex = @{}
foreach ($df in $allDestFiles) {
    $destIndex[$df.FieldValues.FileRef.ToLower()] = $true
}

# --- Comparar ---
Write-Host ""
Write-Host "Analizando diferencias..." -ForegroundColor DarkGray

$missing      = [System.Collections.Generic.List[PSCustomObject]]::new()
$excluded     = 0
$skipped      = 0
$ok           = 0

foreach ($file in $allSourceFiles) {
    $sourceRelUrl = $file.FieldValues.FileRef

    # Filtro scope
    if ($ScopeFolder -ne "") {
        $scopePath = "$baseDocUrl/$ScopeFolder"
        if (-not $sourceRelUrl.StartsWith($scopePath, [System.StringComparison]::InvariantCultureIgnoreCase)) {
            continue
        }
    }

    # Verificar si está en carpeta excluida
    $pathParts      = $sourceRelUrl.Replace($baseDocUrl, "").Trim('/').Split('/')
    $directoryParts = if ($pathParts.Count -gt 1) { $pathParts[0..($pathParts.Count - 2)] } else { @() }
    $exclMatch      = $excludedList | Where-Object { $directoryParts -contains $_ } | Select-Object -First 1

    if ($exclMatch) {
        $excluded++
        continue
    }

    # Construir ruta esperada en destino (con sanitización igual que migrationScript.ps1)
    $relativePathFromDocs = $sourceRelUrl -replace [regex]::Escape($baseDocUrl), ""
    $relativePathFromDocs = $relativePathFromDocs.TrimStart('/')

    $segments     = $relativePathFromDocs.Split('/')
    $safeSegments = $segments | ForEach-Object { Format-SafeName -Name $_ }
    $expectedDestRelUrl = ("$destBaseUrl/" + ($safeSegments -join '/')).ToLower()

    if ($destIndex.ContainsKey($expectedDestRelUrl)) {
        $ok++
    } else {
        # Determinar motivo
        $reason = "No encontrado en destino"

        # ¿Figura en el log con ERROR?
        if ($logIndex.ContainsKey($sourceRelUrl)) {
            $entry = $logIndex[$sourceRelUrl]
            if ($entry.Status -eq "ERROR") {
                $reason = "Error en migración: $($entry.Message)"
            } elseif ($entry.Status -eq "SKIPPED") {
                $reason = "Omitido en migración: $($entry.Message)"
                $skipped++
            }
        } else {
            # ¿El nombre fue sanitizado? (podría causar discrepancia de ruta)
            $rawName  = $segments[-1]
            $safeName = Format-SafeName -Name $rawName
            if ($safeName -ne $rawName) {
                $reason = "Nombre sanitizado ('$rawName' → '$safeName') — verificar si llegó con otro nombre"
            }

            # ¿Nunca apareció en el log? → no fue procesado por el script
            if (-not $logIndex.ContainsKey($sourceRelUrl) -and $UseLatestLog) {
                $reason = "No procesado por el script (fuera del scope o migración no ejecutada aún)"
            }
        }

        $missing.Add([PSCustomObject]@{
            SourcePath   = $sourceRelUrl
            ExpectedDest = $expectedDestRelUrl
            Motivo       = $reason
        })
    }
}

# --- Reporte en consola ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " RESUMEN" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Archivos en origen (tras scope)  : $($ok + $missing.Count + $excluded - $excluded)" -ForegroundColor White
Write-Host "  Migrados correctamente           : $ok" -ForegroundColor Green
Write-Host "  Excluidos por configuración      : $excluded" -ForegroundColor DarkGray
Write-Host "  Faltantes en destino             : $($missing.Count)" -ForegroundColor $(if ($missing.Count -eq 0) { 'Green' } else { 'Red' })
Write-Host ""

if ($missing.Count -gt 0) {
    # Agrupar por motivo
    $byReason = $missing | Group-Object Motivo | Sort-Object Count -Descending
    Write-Host "  Desglose por motivo:" -ForegroundColor Yellow
    foreach ($g in $byReason) {
        Write-Host "    [$($g.Count)] $($g.Name)" -ForegroundColor Yellow
    }
    Write-Host ""

    # Mostrar primeros 20 en consola
    $preview = [math]::Min(20, $missing.Count)
    Write-Host "  Primeros $preview archivos faltantes:" -ForegroundColor White
    $missing | Select-Object -First $preview | ForEach-Object {
        Write-Host "    • $($_.SourcePath)" -ForegroundColor Gray
        Write-Host "      Motivo: $($_.Motivo)" -ForegroundColor DarkYellow
    }
    if ($missing.Count -gt 20) {
        Write-Host "    ... y $($missing.Count - 20) más (ver CSV para lista completa)" -ForegroundColor DarkGray
    }

    # Exportar CSV completo
    if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath | Out-Null }
    $reportFile = Join-Path $logPath "missing-files-$($UserEmail.Split('@')[0])-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"
    $missing | Export-Csv -Path $reportFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Reporte completo exportado a:" -ForegroundColor Green
    Write-Host "  $reportFile" -ForegroundColor Green
} else {
    Write-Host "  ✓ Todos los archivos están en destino." -ForegroundColor Green
}

Write-Host ""
