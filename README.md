# 🚀 Migración OneDrive → SharePoint

![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![SharePoint](https://img.shields.io/badge/SharePoint-0078D4?style=for-the-badge&logo=microsoft-sharepoint&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)
![PnP](https://img.shields.io/badge/PnP-PowerShell-blue?style=for-the-badge)
![Version](https://img.shields.io/badge/version-5.0.0-blue?style=for-the-badge)

Sistema automatizado y optimizado para migrar contenido de OneDrive personal a bibliotecas de SharePoint con autenticación por certificado, checkpoint reanudable y verificación por carpeta.

**Tecnología:** PnP PowerShell
**Versión:** 5.0.0
**Última actualización:** Abril 2026

---

## 📑 Tabla de Contenido

- [Características Principales](#-características-principales)
- [Inicio Rápido](#-inicio-rápido)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Scripts Disponibles](#-scripts-disponibles)
- [Requisitos Previos](#-requisitos-previos)
- [Configuración Inicial](#-configuración-inicial)
- [Configuración del Archivo .env](#-configuración-del-archivo-env)
- [Ejecución de la Migración](#-ejecución-de-la-migración)
- [Checkpoint y Reanudación](#-checkpoint-y-reanudación)
- [Verificación y Validación](#-verificación-y-validación)
- [Logs y Reportes](#-logs-y-reportes)
- [Optimización y Rendimiento](#-optimización-y-rendimiento)
- [Solución de Problemas](#-solución-de-problemas)
- [Seguridad](#-seguridad)

---

## ✨ Características Principales

### Migración Optimizada (V5)
- ✅ **Autenticación automática** mediante certificado (sin MFA interactivo)
- ✅ **Escaneo paginado en streaming** - Consultas CAML por carpeta, sin límite de 5000 ítems
- ✅ **Checkpoint reanudable** - Guarda progreso en JSON, retoma desde donde se interrumpió
- ✅ **Sobreescritura inteligente** - Reemplaza archivos modificados más de 20 min después de migrar
- ✅ **Múltiples scopes** - Migra varias carpetas en una sola ejecución con `-ScopeFolder`
- ✅ **Paralelismo configurable** - Hilos simultáneos con control de throttling
- ✅ **Reconexión en retry** - Maneja errores de runspace bajo carga paralela
- ✅ **Exclusión por carpeta y extensión** - Omite carpetas del sistema y archivos temporales (locks de ArcGIS, etc.)
- ✅ **Sanitización de nombres** - Reemplaza caracteres prohibidos en SharePoint automáticamente
- ✅ **Gestión de memoria** - Garbage collection optimizado para migraciones de 100GB+

### Herramientas de Verificación
- 📊 **check-storage.ps1** - Compara origen vs destino por carpeta, genera reporte ejecutivo
- 🔍 **check-missing.ps1** - Diagnóstico de archivos faltantes
- 📋 **validate-filenames.ps1** - Detecta rutas largas y caracteres problemáticos

---

## 🚀 Inicio Rápido

### 1. Configurar el Entorno

```powershell
# Instalar módulo PnP
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Permitir ejecución de scripts
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Editar configuración
notepad .env
```

### 2. Ejecutar la Migración

```powershell
# Primera ejecución completa
.\migrationScript.ps1 -Force

# Migrar solo una carpeta
.\migrationScript.ps1 -Force -ScopeFolder "Proyectos/2026"

# Migrar múltiples carpetas
.\migrationScript.ps1 -Force -ScopeFolder "Carpeta1,Carpeta2"

# Reanudar migración interrumpida (usa checkpoint automáticamente)
.\migrationScript.ps1 -Force
```

### 3. Verificar Resultados

```powershell
# Verificar una carpeta específica
.\check-storage.ps1 -ScopeFolder "Proyectos/2026"

# Verificar con usuario explícito
.\check-storage.ps1 -UserEmail "usuario@dominio.com" -ScopeFolder "Proyectos/2026"

# Verificar todo el OneDrive
.\check-storage.ps1
```

---

## 📁 Estructura del Proyecto

```
SharePointMigration\
├── Scripts Principales
│   └── migrationScript.ps1              # Script de migración V5
│
├── Herramientas de Verificación
│   ├── check-storage.ps1                # Comparación origen vs destino
│   ├── check-missing.ps1                # Diagnóstico de archivos faltantes
│   └── validate-filenames.ps1           # Validación de nombres de archivo
│
├── Configuración
│   ├── .env                             # Configuración (NO VERSIONAR)
│   ├── .gitignore                       # Archivos excluidos del control de versiones
│
├── Certificados
│   ├── PnPMigrationCert.pfx             # Certificado de autenticación (NO VERSIONAR)
│   └── PnPMigrationCert.cer             # Certificado público (para Entra ID)
│
├── Documentación
│   └── README.md                        # Este archivo
│
└── Logs/                                # Generado automáticamente
    ├── migration-report-*.csv           # Reporte detallado de migración
    ├── storage-comparison-*.csv         # Comparativa ejecutiva (para gerencia)
    ├── filename-validation-report-*.csv # Reporte de validación de nombres
    └── checkpoint-<usuario>.json        # Checkpoint de progreso por usuario
```

---

## 🛠️ Scripts Disponibles

### Scripts de Migración

| Script | Descripción | Uso Recomendado |
|--------|-------------|-----------------|
| **migrationScript.ps1** | Migración V5 con streaming, checkpoint y multi-scope | ✅ Producción (100GB+) |

### Scripts de Verificación

| Script | Parámetros | Descripción |
|--------|------------|-------------|
| **check-storage.ps1** | `-UserEmail` `-ScopeFolder` | Compara archivos y tamaño OneDrive vs SharePoint por carpeta |
| **check-missing.ps1** | — | Diagnóstico de archivos faltantes en destino |
| **validate-filenames.ps1** | `-UserEmail` | Detecta nombres problemáticos antes de migrar |

---

## 📋 Requisitos Previos

| Requisito | Detalle |
|-----------|---------|
| **Sistema operativo** | Windows 10/11 con PowerShell 7+ |
| **Cuenta de administrador** | Rol Global Admin o SharePoint Admin en Microsoft 365 |
| **Módulo PnP.PowerShell** | Versión 2.x o superior |
| **Aplicación en Entra ID** | App Registration con permisos de SharePoint |
| **Certificado** | Archivo `.pfx` generado y subido a Entra ID |
| **Red** | Conexión estable de 100+ Mbps (ideal 500+ Mbps para migraciones grandes) |

### Instalación de PowerShell 7+

```powershell
winget install --id Microsoft.Powershell --source winget
pwsh --version
```

---

## ⚙️ Configuración Inicial

Estos pasos se realizan **una sola vez** por tenant.

### 1. Configurar Política de Ejecución

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 2. Instalar PnP.PowerShell

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser -Force
Get-Module PnP.PowerShell -ListAvailable | Select-Object Version
```

### 3. Registrar Aplicación en Entra ID

1. Ir a [https://entra.microsoft.com](https://entra.microsoft.com)
2. Navegar a **Entra ID → Registros de aplicaciones → + Nuevo registro**
3. Completar:
   - **Nombre:** `PnP Migration Tool`
   - **Tipos de cuenta:** Solo inquilino único
   - **URI de redirección:** Web → `http://localhost`
4. Copiar el **Id. de aplicación (cliente)**

**Agregar permisos de API:**

1. **Permisos de API → + Agregar un permiso**
2. SharePoint → Permisos delegados → `AllSites.FullControl`
3. SharePoint → Permisos de aplicación → `Sites.FullControl.All`
4. **Conceder consentimiento de administrador**

**Habilitar flujo de cliente público:**

1. **Autenticación → Configuración avanzada**
2. Activar **Permitir flujos de cliente público → Sí**

### 4. Generar Certificado

```powershell
$projectPath  = "C:\Path\To\Your\Project"
$certName     = "PnPMigrationCert"
$certPassword = "YourSecurePassword123!"

$cert = New-SelfSignedCertificate `
    -Subject "CN=$certName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

Export-Certificate -Cert $cert -FilePath "$projectPath\$certName.cer"

$pwd = ConvertTo-SecureString -String $certPassword -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "$projectPath\$certName.pfx" -Password $pwd

Write-Host "Thumbprint: $($cert.Thumbprint)"
```

### 5. Subir Certificado a Entra ID

1. **Entra ID → Registros de aplicaciones → PnP Migration Tool**
2. **Certificados y secretos → Certificados → + Cargar certificado**
3. Seleccionar el archivo `.cer` generado

---

## 📝 Configuración del Archivo .env

El archivo `.env` contiene toda la configuración. **Nunca versionar este archivo.**

### Plantilla Completa

```bash
# ============================================================
# SharePoint Migration Environment Variables
# ============================================================

# SharePoint URLs
MY_URL=https://your-tenant-my.sharepoint.com/personal
DESTINATION_SITE_URL=https://your-tenant.sharepoint.com/sites/YourSite
DESTINATION_LIBRARY=DestinationLibrary

# Authentication
ADMIN_ACCOUNT=admin@your-domain.com
CLIENT_ID=your-client-id-here
TENANT=your-tenant.onmicrosoft.com

# Certificate (must be in same folder as the script)
CERT_FILENAME=PnPMigrationCert.pfx
CERT_PASSWORD=YourSecurePassword123!

# Exclusions — carpetas del sistema omitidas (comma-separated)
EXCLUDED_FOLDERS=Archivos de Microsoft Copilot Chat,Attachments,Forms,Notebook,Apps,_vti_cnf,Archivos de chat de Microsoft Teams,Documentos de Microsoft Teams,Grabaciones,backup_correos

# Exclusions — extensiones de archivo omitidas (comma-separated, include the dot)
# .lock/.sr.lock = ArcGIS shapefile locks | .ldb = geodatabase locks | .tmp = temporales
EXCLUDED_EXTENSIONS=.lock,.sr.lock,.ldb,.tmp,.~lock

# Users to migrate (comma-separated — solo el primero se migra por ejecución)
USERS=user@your-domain.com

# Performance Settings
THROTTLE_LIMIT=12         # Hilos paralelos simultáneos
BATCH_SIZE=300            # Archivos por sub-lote paralelo
MAX_RETRIES=3             # Reintentos por archivo
RETRY_DELAY_429=5         # Delay base para errores 429 (segundos, se multiplica por intento)
RETRY_DELAY_DEFAULT=2     # Delay para otros errores (segundos)
PAGE_SIZE=500             # Tamaño de página CAML para consultas a SharePoint
BATCH_PAUSE_SECONDS=3     # Pausa entre lotes (0 = sin pausa, recomendado 3+ para 100GB+)
LOG_DIRECTORY=Logs        # Carpeta para logs y checkpoint
```

### Configuración por Escenario

| Escenario | THROTTLE_LIMIT | BATCH_SIZE | BATCH_PAUSE_SECONDS |
|-----------|----------------|------------|---------------------|
| **Pruebas** | 5 | 50 | 0 |
| **Producción estándar** | 10 | 200 | 3 |
| **Producción optimizada** ✅ | 12 | 300 | 3 |
| **Red lenta (<50 Mbps)** | 5 | 100 | 5 |
| **VM en Azure** | 15 | 300 | 1 |

---

## 🎯 Ejecución de la Migración

### Preparación Previa

Antes de migrar cada usuario, abrir su OneDrive en el navegador como administrador para asegurarse de que esté provisionado:

```
https://[tenant]-my.sharepoint.com/personal/[email_con_guiones_bajos]

Ejemplo:
  Email: user@your-domain.com
  URL:   https://your-tenant-my.sharepoint.com/personal/user_your-domain_com
```

### Modos de Ejecución

```powershell
# Migración completa desatendida
.\migrationScript.ps1 -Force

# Migrar solo una carpeta
.\migrationScript.ps1 -Force -ScopeFolder "Proyectos/2026"

# Migrar múltiples carpetas separadas por coma
.\migrationScript.ps1 -Force -ScopeFolder "RUP,Monitoreos Fauna"

# Reanudar — el checkpoint se carga automáticamente, no se necesita flag adicional
.\migrationScript.ps1 -Force
```

### Salida en Tiempo Real

```
========================================
MIGRACIÓN ONEDRIVE → SHAREPOINT V5
========================================
ThrottleLimit: 12 | BatchSize: 300 | PageSize: 500 | PauseLote: 3s
[SCOPE ACTIVADO] Solo se migrará: RUP

Procesando usuario: user@your-domain.com
  [CHECKPOINT] Cargando progreso previo desde Logs/checkpoint-user.json...
  [CHECKPOINT] 1408 archivos ya migrados (se saltarán si no tienen cambios mayores a 20 min).
  Conectando a SharePoint...
  Expandiendo árbol de carpetas: /personal/.../Documents/RUP
  -> 47 carpetas encontradas (incluyendo raíz)
  Escaneando: /personal/.../Documents/RUP
    Pág 1: 12 archivos
  -> Lote 1 (12 arch) | ✓ 0 ✗ 0 ⏭ 0 | ...

  ========================================
  RESUMEN DE MIGRACIÓN - user@your-domain.com
  ========================================
  Archivos procesados: 26
  Exitosos:            26
  Errores:             0
  Saltados (checkpoint/excluidos): 1408
  Tiempo total:        0h 4m 12s
  Velocidad promedio:  6.2 archivos/min
  Checkpoint guardado: Logs/checkpoint-user.json
  ========================================
```

---

## 💾 Checkpoint y Reanudación

El script guarda automáticamente el progreso en un archivo JSON después de cada lote.

**Ubicación:** `Logs/checkpoint-<usuario>.json`

```
Logs\checkpoint-area-biotica.json
Logs\checkpoint-gerencia.json
```

**Comportamiento por archivo:**

| Situación | Acción |
|-----------|--------|
| No está en checkpoint | Se migra normalmente |
| En checkpoint, sin cambios recientes (≤20 min) | Se salta |
| En checkpoint, modificado >20 min después de migrar | Se sobreescribe |
| En checkpoint con error previo | Se reintenta |

Para **reiniciar desde cero** (ignorar checkpoint), eliminar el archivo JSON correspondiente antes de ejecutar.

---

## 🔍 Verificación y Validación

### Verificación de Almacenamiento

```powershell
# Verificar una carpeta específica (recomendado al migrar por scope)
.\check-storage.ps1 -ScopeFolder "RUP"

# Verificar subcarpeta específica
.\check-storage.ps1 -ScopeFolder "RUP/CONTRATOS FINALIZADOS"

# Verificar con usuario explícito
.\check-storage.ps1 -UserEmail "user@your-domain.com" -ScopeFolder "RUP"

# Verificar todo el OneDrive del primer usuario del .env
.\check-storage.ps1
```

**Salida en consola (técnica):**
- Total archivos y tamaño en OneDrive y SharePoint
- Diferencia exacta en bytes y MB
- Archivos con tamaño 0 o nulo (locks, temporales)
- Porcentaje migrado y resultado

**Archivo generado (ejecutivo):** `Logs/storage-comparison-<usuario>-<scope>-<fecha>.csv`
```
STORAGE COMPARISON
Date;13/04/2026 15:30
User;user@your-domain.com
Migrated folder;RUP

SOURCE (OneDrive)
Total source files;1413

DESTINATION (SharePoint)
Total migrated files;1413
Completion percentage;100%

RESULT;EXITOSA
```

### Validación de Nombres de Archivo

```powershell
.\validate-filenames.ps1 -UserEmail "user@your-domain.com"
```

**Detecta:**
- 🔴 **CRÍTICO:** Rutas >350 caracteres, caracteres prohibidos (`" * : < > ? / \ |`)
- ⚠️ **WARNING:** Rutas >300 caracteres, nombres largos, estructura profunda
- ℹ️ **INFO:** Archivos .zip en rutas largas

---

## 📊 Logs y Reportes

| Archivo | Descripción | Generado por |
|---------|-------------|--------------|
| `Logs/migration-report-*.csv` | Log detallado con cada archivo procesado (OK/SKIPPED/ERROR) | migrationScript.ps1 |
| `Logs/checkpoint-<usuario>.json` | Progreso de migración para reanudar | migrationScript.ps1 |
| `Logs/storage-comparison-*.csv` | Comparativa ejecutiva para gerencia | check-storage.ps1 |
| `Logs/filename-validation-report-*.csv` | Archivos con nombres problemáticos | validate-filenames.ps1 |

### Formato del Log de Migración

```csv
Timestamp;Step;SourcePath;DestPath;Status;Message
2026-04-13T14:30:25;Migrate;/personal/.../archivo.pdf;Library/archivo.pdf;OK;Copiado
2026-04-13T14:30:27;Scan;/personal/.../archivo.sr.lock;;SKIPPED;Extensión excluida
2026-04-13T14:30:30;Scan;/personal/.../carpeta/archivo.pdf;;SKIPPED;Carpeta Excluida (backup_correos)
```

**Estatus posibles:**
- `OK` — Archivo copiado exitosamente
- `SKIPPED` — Omitido (checkpoint, carpeta excluida, extensión excluida)
- `ERROR` — Error al copiar (ver columna Message)

---

## ⚡ Optimización y Rendimiento

### Factores que Afectan la Velocidad

| Factor | Impacto | Recomendación |
|--------|---------|---------------|
| **Ancho de banda de red** | 70% | Ethernet 500+ Mbps, evitar WiFi |
| **Latencia de red** | 20% | Ejecutar desde VM en Azure (misma región que SharePoint) |
| **Límites de SharePoint (throttling)** | 10% | Mantener `THROTTLE_LIMIT ≤ 12` y `BATCH_PAUSE_SECONDS ≥ 3` |

### Velocidades Esperadas

| Red | Velocidad estimada |
|-----|--------------------|
| 10 Mbps | ~20 arch/min |
| 100 Mbps | ~50 arch/min |
| 500 Mbps | ~70 arch/min |
| VM en Azure | ~150 arch/min |

### Estrategias para Migraciones Grandes (100GB+)

**Opción 1: Migrar por carpetas con `-ScopeFolder`**
```powershell
.\migrationScript.ps1 -Force -ScopeFolder "Carpeta1"
.\migrationScript.ps1 -Force -ScopeFolder "Carpeta2"
```
Control granular, fácil de verificar por partes con `check-storage.ps1 -ScopeFolder`.

**Opción 2: VM en Azure**
1. Crear VM Windows en Azure (misma región que el tenant de SharePoint)
2. Copiar scripts, `.env` y `.pfx` a la VM
3. Ejecutar desde la VM — latencia mínima, ancho de banda máximo
4. Costo estimado: ~$50-150 USD por migración completa

**Opción 3: Horario no laboral**
Ejecutar de noche para aprovechar la red sin competencia y reducir impacto en usuarios.

---

## ❗ Solución de Problemas

### Errores Comunes

| Error | Causa | Solución |
|-------|-------|----------|
| `Access is denied` | OneDrive no provisionado | Abrir OneDrive del usuario en navegador como admin |
| `Certificate not found` | `.pfx` no en la carpeta del script | Verificar que `.pfx` está junto a `migrationScript.ps1` |
| `Unauthorized` | Permisos insuficientes | Verificar permisos de la app en Entra ID y consentimiento de admin |
| `429 Too Many Requests` | Throttling de SharePoint | Manejado automáticamente con backoff exponencial |
| `Nullable object must have a value` | Error de conexión en runspace paralelo | Manejado automáticamente con reconexión en retry |
| `String was not recognized as DateTime` | Cultura regional diferente entre equipos | Resuelto en V5 con `InvariantCulture` |

### Archivos con Nombres Problemáticos

SharePoint rechaza algunos caracteres: `" * : < > ? / \ |`

El script los sanitiza automáticamente al migrar (reemplaza por `_`). Para identificarlos previamente:

```powershell
.\validate-filenames.ps1 -UserEmail "user@your-domain.com"
```

### Transferir el Proyecto a Otro Equipo

El proyecto es portable — el certificado no está atado al hardware. Para mover a otro equipo:

1. Copiar carpeta del proyecto (o clonar con `git clone`)
2. Copiar manualmente `.env` y `PnPMigrationCert.pfx` (no están en git)
3. En el equipo destino: instalar PowerShell 7+ y `PnP.PowerShell`
4. Ejecutar `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

> **Importante:** Transferir via `git clone` o ZIP — copiar archivos sueltos puede corromper el encoding UTF-8 del script y causar errores de parseo en PowerShell.

---

## 🔐 Seguridad

### Archivos Sensibles — NO Compartir

- ❌ `.env` — Credenciales, URLs, contraseña del certificado
- ❌ `*.pfx` — Certificado de autenticación
- ❌ `Logs/*.csv` — Pueden contener rutas y nombres de archivos privados
- ❌ `Logs/*.json` — Checkpoint con rutas de archivos del usuario

El `.gitignore` ya excluye estos archivos automáticamente.

### Sobre el Certificado

- **¿Está atado a un equipo?** No. Funciona en cualquier equipo que tenga el `.pfx` y la contraseña del `.env`.
- **¿Cuándo vence?** A los 2 años. Generar uno nuevo y subirlo a Entra ID cuando venza.
- **¿Es seguro?** El `.pfx` está protegido con contraseña. Sin ella no puede usarse.

---

## 🤝 Contribuciones

- ✅ Crear rama desde `develop`, hacer PR para merge
- ✅ Actualizar documentación al agregar funcionalidades
- ✅ Probar con datos de prueba antes de producción
- ❌ **NUNCA** incluir credenciales reales en commits
- ❌ **NUNCA** subir `.env`, `.pfx` o archivos de `Logs/`

---

## 📞 Soporte

- 🐛 Reportar problemas: [Crear issue](https://github.com/Ariaz39/sharepoint-from-onedrive-migration-tool/issues)
- 💡 Sugerencias: Pull requests bienvenidos
- 📧 Contacto: [LinkedIn - Alejandro Ariaz](https://www.linkedin.com/in/alejandro-ariaz/)

---

## 👨‍💻 Autor

**Desarrollado por:** Alejandro Ariaz
**LinkedIn:** [linkedin.com/in/alejandro-ariaz](https://www.linkedin.com/in/alejandro-ariaz/)
**GitHub:** [@Ariaz39](https://github.com/Ariaz39)

---

## 📄 Licencia

MIT License — puedes usar, modificar y distribuir libremente, incluso para proyectos comerciales. Solo se requiere mantener el aviso de copyright. Ver [LICENSE](LICENSE).
