# 🚀 Migración OneDrive → SharePoint

![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![SharePoint](https://img.shields.io/badge/SharePoint-0078D4?style=for-the-badge&logo=microsoft-sharepoint&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)
![PnP](https://img.shields.io/badge/PnP-PowerShell-blue?style=for-the-badge)

Sistema automatizado y optimizado para migrar contenido de OneDrive personal a bibliotecas de SharePoint con autenticación por certificado, progreso en tiempo real y verificación automática.

**Tecnología:** PnP PowerShell
**Versión:** 4.0 (Optimizado para millones de archivos)
**Última actualización:** Febrero 2026

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
- [Verificación y Validación](#-verificación-y-validación)
- [Logs y Reportes](#-logs-y-reportes)
- [Optimización y Rendimiento](#-optimización-y-rendimiento)
- [Solución de Problemas](#-solución-de-problemas)
- [Seguridad](#-seguridad)

---

## ✨ Características Principales

### Migración Optimizada (V4)
- ✅ **Autenticación automática** mediante certificado (sin MFA interactivo)
- ✅ **Paralelismo inteligente** - 10 hilos simultáneos con control de throttling
- ✅ **Procesamiento por lotes** - 200 archivos por lote para mejor visibilidad
- ✅ **Barra de progreso en tiempo real** - Velocidad, ETA, contadores de éxito/error
- ✅ **Verificación automática** - Compara origen vs destino al finalizar
- ✅ **Validación de nombres** - Detecta archivos con nombres problemáticos
- ✅ **Gestión de memoria** - Garbage collection optimizado para migraciones grandes
- ✅ **Exclusión inteligente** - Omite carpetas del sistema automáticamente
- ✅ **Resistente a errores** - Reintentos automáticos con backoff exponencial

### Herramientas Adicionales
- 📊 **Monitoreo de rendimiento** - CPU, RAM, disco en tiempo real
- 🔍 **Validación de archivos** - Detecta rutas largas, caracteres especiales, archivos comprimidos
- 🧹 **Limpieza de pruebas** - Vacía bibliotecas de SharePoint para testing
- 🔓 **Desbloqueo de archivos** - Libera archivos bloqueados en SharePoint

---

## 🚀 Inicio Rápido

### 1. Configurar el Entorno

```powershell
# Copiar la plantilla de configuración
Copy-Item .env.example .env

# Editar el archivo .env con tus credenciales
notepad .env
```

### 2. Ejecutar la Migración

```powershell
# Primera ejecución (modo conservador)
.\migrationScriptGemini.ps1

# Ejecución desatendida
.\migrationScriptGemini.ps1 -Force

# Reanudar migración interrumpida
.\migrationScriptGemini.ps1 -Force -SkipExisting

# Migrar solo una carpeta específica
.\migrationScriptGemini.ps1 -ScopeFolder "Proyectos/2026"
```

### 3. Verificar Resultados

La verificación se ejecuta automáticamente al finalizar. También puedes ejecutarla manualmente:

```powershell
# Verificar migración de un usuario
.\check-storage.ps1 -UserEmail "usuario@dominio.com"
```

---

## 📁 Estructura del Proyecto

```
SharePointMigration\
├── Scripts Principales
│   ├── migrationScriptGemini.ps1    # Script de migración V4 (optimizado)
│   ├── migrationScript.ps1          # Script de migración V1 (legacy)
│
├── Herramientas de Verificación
│   ├── check-storage.ps1            # Verificación de almacenamiento (origen vs destino)
│   ├── validate-filenames.ps1       # Validación de nombres de archivo
│   ├── check-performance.ps1        # Monitoreo de rendimiento del sistema
│
├── Utilidades
│   ├── delete-test-folder.ps1       # Vaciar biblioteca de SharePoint para pruebas
│   ├── unlock-sharepoint-files.ps1  # Desbloquear archivos en SharePoint
│
├── Configuración
│   ├── .env                         # Configuración (NO VERSIONAR)
│   ├── .env.example                 # Plantilla de ejemplo
│   ├── .gitignore                   # Archivos excluidos del control de versiones
│
├── Certificados
│   ├── PnPMigrationCert.pfx         # Certificado de autenticación (NO VERSIONAR)
│   └── PnPMigrationCert.cer         # Certificado público (para Entra ID)
│
├── Documentación
│   └── README.md                    # Este archivo
│
└── Logs/                            # Generado automáticamente
    ├── migration-report-v4-*.csv    # Reporte de migración
    └── filename-validation-report-*.csv  # Reporte de validación
```

---

## 🛠️ Scripts Disponibles

### Scripts de Migración

| Script | Descripción | Uso Recomendado |
|--------|-------------|-----------------|
| **migrationScriptGemini.ps1** | Versión optimizada V4 con progreso en tiempo real | ✅ Producción (1M+ archivos) |
| **migrationScript.ps1** | Versión original V1 | 📦 Legacy (migraciones pequeñas) |

### Scripts de Verificación

| Script | Descripción | Ejecutar |
|--------|-------------|----------|
| **check-storage.ps1** | Compara OneDrive vs SharePoint (archivos, tamaño, porcentaje migrado) | Post-migración |
| **validate-filenames.ps1** | Detecta archivos con nombres problemáticos (rutas largas, caracteres especiales) | Pre/Post-migración |
| **check-performance.ps1** | Monitorea CPU, RAM, disco durante la migración | Durante migración |

### Scripts de Utilidades

| Script | Descripción | Uso |
|--------|-------------|-----|
| **delete-test-folder.ps1** | Vacía una biblioteca de SharePoint (sin eliminar la biblioteca raíz) | Testing |
| **unlock-sharepoint-files.ps1** | Libera archivos bloqueados (checked out) en SharePoint | Troubleshooting |

---

## 📋 Requisitos Previos

| Requisito | Detalle |
|-----------|---------|
| **Sistema operativo** | Windows 10/11 con PowerShell 7+ (recomendado) o PowerShell 5.1 |
| **Cuenta de administrador** | Cuenta con rol Global Admin o SharePoint Admin en Microsoft 365 |
| **Módulo PnP.PowerShell** | Versión 3.1.0 o superior |
| **Aplicación en Entra ID** | App Registration con permisos de SharePoint |
| **Certificado** | Archivo `.pfx` generado y subido a Entra ID |
| **Red** | Conexión estable de 100+ Mbps (ideal 500+ Mbps para migraciones grandes) |

### Instalación de PowerShell 7+ (Recomendado)

```powershell
# Instalar PowerShell 7+
winget install --id Microsoft.Powershell --source winget

# Verificar versión
pwsh -Version
```

---

## ⚙️ Configuración Inicial

Estos pasos se realizan **una sola vez**. No es necesario repetirlos para migraciones futuras.

### 1. Configurar Política de Ejecución

```powershell
# Permitir ejecución de scripts locales
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Desbloquear el script (si fue descargado de internet)
Unblock-File -Path ".\migrationScriptGemini.ps1"
```

### 2. Instalar PnP.PowerShell

```powershell
# Instalar módulo
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Verificar instalación
Get-Module PnP.PowerShell -ListAvailable | Select-Object Version
```

### 3. Registrar Aplicación en Entra ID

1. Ir a [https://entra.microsoft.com](https://entra.microsoft.com)
2. Navegar a **Entra ID → Registros de aplicaciones**
3. Hacer clic en **+ Nuevo registro**
4. Completar:
   - **Nombre:** `PnP Migration Tool`
   - **Tipos de cuenta:** Solo inquilino único
   - **URI de redirección:** Web → `http://localhost`
5. Copiar el **Id. de aplicación (cliente)**

**Agregar permisos de API:**

1. Ir a **Permisos de API → + Agregar un permiso**
2. Seleccionar **SharePoint → Permisos delegados** → `AllSites.FullControl`
3. Seleccionar **SharePoint → Permisos de aplicación** → `Sites.FullControl.All`
4. Hacer clic en **Conceder consentimiento de administrador**

**Habilitar flujo de cliente público:**

1. Ir a **Autenticación → Configuración avanzada**
2. Activar **Permitir flujos de cliente público** → **Sí**
3. Guardar

### 4. Generar Certificado

```powershell
# Ajustar ruta a la carpeta del proyecto
$projectPath = "C:\Path\To\Your\Project"

# Definir nombre y contraseña
$certName = "PnPMigrationCert"
$certPassword = "YourSecurePassword123!"

# Generar certificado
$cert = New-SelfSignedCertificate `
    -Subject "CN=$certName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Exportar certificado público (.cer) para Entra ID
Export-Certificate `
    -Cert $cert `
    -FilePath "$projectPath\$certName.cer"

# Exportar certificado completo (.pfx) para el script
$pwd = ConvertTo-SecureString -String $certPassword -Force -AsPlainText
Export-PfxCertificate `
    -Cert $cert `
    -FilePath "$projectPath\$certName.pfx" `
    -Password $pwd

Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
Write-Host "Certificate exported: $certName.pfx" -ForegroundColor Green
```

### 5. Subir Certificado a Entra ID

1. Ir a **Entra ID → Registros de aplicaciones → PnP Migration Tool**
2. Hacer clic en **Certificados y secretos → Certificados**
3. Hacer clic en **+ Cargar certificado**
4. Seleccionar el archivo `.cer` generado
5. Verificar que aparece con el thumbprint correcto

---

## 📝 Configuración del Archivo .env

El archivo `.env` contiene toda la configuración necesaria.

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

# Certificate
CERT_FILENAME=PnPMigrationCert.pfx
CERT_PASSWORD=YourSecurePassword123!

# Exclusions (carpetas del sistema que se omiten)
EXCLUDED_FOLDERS=Archivos de Microsoft Copilot Chat,Attachments,Forms,Notebook,Apps,_vti_cnf,Archivos de chat de Microsoft Teams,Documentos de Microsoft Teams,Grabaciones,backup_correos

# Users to migrate (solo el primero se migra por ejecución)
USERS=user@your-domain.com

# Performance Settings (OPTIMIZADO V4)
THROTTLE_LIMIT=10         # Hilos paralelos (10 = óptimo para estabilidad)
BATCH_SIZE=200            # Archivos por lote (200 = feedback cada 2 min)
MAX_RETRIES=3             # Reintentos por archivo
RETRY_DELAY_429=5         # Delay para errores de throttling (segundos)
RETRY_DELAY_DEFAULT=2     # Delay para otros errores (segundos)
PAGE_SIZE=1000            # Tamaño de página para consultas a SharePoint
LOG_DIRECTORY=Logs        # Carpeta para logs
```

### Configuración por Escenario

| Escenario | THROTTLE_LIMIT | BATCH_SIZE | Notas |
|-----------|----------------|------------|-------|
| **Pruebas/desarrollo** | 5 | 50 | Feedback cada 30 seg |
| **Producción balanceada** | 10 | 200 | ✅ **Recomendado** - Feedback cada 2 min |
| **Migración masiva (1M archivos)** | 10 | 200 | Mismo que producción |
| **Red lenta (<50 Mbps)** | 5 | 100 | Reduce carga de red |
| **Red muy rápida (1 Gbps+)** | 15 | 300 | Máximo rendimiento (riesgo throttling) |

---

## 🎯 Ejecución de la Migración

### Preparación Previa

**IMPORTANTE:** Antes de migrar cada usuario, abrir su OneDrive en el navegador:

```
URL: https://[tenant]-my.sharepoint.com/personal/[email_con_guiones_bajos]

Ejemplo:
Email: user@your-domain.com
URL: https://your-tenant-my.sharepoint.com/personal/user_your-domain_com
```

Esto garantiza que el OneDrive esté provisionado y accesible.

### Modos de Ejecución

```powershell
# Modo básico (con confirmaciones)
.\migrationScriptGemini.ps1

# Modo desatendido (sin confirmaciones)
.\migrationScriptGemini.ps1 -Force

# Reanudar migración interrumpida (omite archivos existentes)
.\migrationScriptGemini.ps1 -Force -SkipExisting

# Migrar solo una carpeta específica
.\migrationScriptGemini.ps1 -ScopeFolder "Proyectos/2026"
```

### Salida en Tiempo Real

Durante la ejecución verás:

```
========================================
MIGRACIÓN ONEDRIVE → SHAREPOINT V4
========================================
ThrottleLimit: 10 | BatchSize: 200 | PageSize: 1000

Procesando usuario: user@your-domain.com
  Conectando a SharePoint...
  Escaneando archivos...
  Archivos a migrar tras aplicar filtros: 1899
  Creando estructura de carpetas (45 rutas únicas)...
  ✓ Estructura de carpetas creada (45 rutas en 8.2s)

Migrando archivos de user@your-domain.com
Lote 5 de 10 | 800/1899 archivos | ✓ 795 ✗ 5
[████████████████░░░░░░░░] 42.1%
Velocidad: 71.3 archivos/min | ETA: 15.4 min

  -> Lote 5/10 (200 archivos) | Progreso: 42.1% | ✓ 795 ✗ 5

  ========================================
  RESUMEN DE MIGRACIÓN - user@your-domain.com
  ========================================
  Archivos procesados: 1899
  Exitosos: 1894
  Errores: 5
  Tiempo total: 0h 26m 40s
  Velocidad promedio: 71.2 archivos/min
  ========================================

========================================
VERIFICACIÓN AUTOMÁTICA DE MIGRACIÓN
========================================

Verificando usuario: user@your-domain.com

=== ONEDRIVE (Origen) ===
  Total archivos: 1899
  Total tamaño: 3.69 GB

  Archivos excluidos: 503
  Tamaño excluido: 1.48 GB
  Desglose de exclusiones:
    • backup_correos: 500 archivos, 1.45 GB

  Archivos a migrar: 1396
  Tamaño a migrar: 2.21 GB

=== SHAREPOINT (Destino) ===
  Archivos: 1396
  Tamaño total: 2.21 GB

=== COMPARACIÓN ===
  Archivos esperados (sin exclusiones): 1396
  Archivos en SharePoint: 1396
  Porcentaje migrado: 100.0%

  ✓ Todos los archivos migrados correctamente
  ✓ Tamaños coinciden

=== RESULTADO ===
  ✓ MIGRACIÓN EXITOSA
```

---

## 🔍 Verificación y Validación

### Verificación de Almacenamiento

```powershell
# Verificar migración de un usuario específico
.\check-storage.ps1 -UserEmail "user@your-domain.com"
```

**Reporte generado:**
- Total de archivos y tamaño en OneDrive
- Archivos excluidos (con desglose por carpeta)
- Total esperado vs migrado
- Porcentaje de éxito
- Resultado automático (EXITOSA/INCOMPLETA)

### Validación de Nombres de Archivo

```powershell
# Validar nombres de archivo antes o después de migrar
.\validate-filenames.ps1 -UserEmail "user@your-domain.com"
```

**Detecta:**
- 🔴 **CRÍTICO:** Rutas >350 caracteres, caracteres prohibidos (`" * : < > ? / \ |`)
- ⚠️ **WARNING:** Rutas >300 caracteres, nombres largos, estructura profunda
- ℹ️ **INFO:** Archivos .zip en rutas largas, carpetas duplicadas

**Reporte CSV generado:**
```csv
Severidad,Ruta_Archivo,Nombre_Archivo,Tipo_Archivo,Longitud_Ruta,Profundidad_Carpetas,Problemas
CRITICAL,Projects/.../archivo.pdf,archivo.pdf,Normal,405,8,"Ruta excede límite (405 > 350 caracteres)"
WARNING,Documents/.../doc.docx,doc.docx,Normal,320,5,"Ruta cerca del límite (320 > 300 caracteres)"
INFO,Data/.../datos.zip,datos.zip,Comprimido,285,6,"Archivo comprimido en ruta larga"
```

### Monitoreo de Rendimiento

```powershell
# Monitorear rendimiento del sistema durante la migración
.\check-performance.ps1
```

**Muestra:**
- CPU usage
- RAM usage (total y disponible)
- Procesos PowerShell activos
- Top 5 procesos por memoria
- Espacio en disco C:
- Estado general del sistema

---

## 📊 Logs y Reportes

### Logs de Migración

**Ubicación:** `Logs/migration-report-v4-YYYYMMDD-HHmmss.csv`

**Formato:**
```csv
Timestamp,User,Step,Item,Status,Message
2026-02-28T14:30:25,user@your-domain...,Migrate,Documents/archivo.pdf,OK,Copiado
2026-02-28T14:30:27,user@your-domain...,Migrate,Documents/doc.docx,SKIPPED,File exists in destination
2026-02-28T14:30:30,user@your-domain...,Migrate,Documents/inv@lid.pdf,ERROR,Caracteres no permitidos
```

**Estatus posibles:**
- `OK` - Archivo copiado exitosamente
- `SKIPPED` - Archivo omitido (ya existe o carpeta excluida)
- `ERROR` - Error al copiar archivo

### Reportes de Validación

**Ubicación:** `Logs/filename-validation-report-YYYYMMDD-HHmmss.csv`

**Categorías:**
- 🔴 **CRITICAL** - Problemas que impedirán la migración
- ⚠️ **WARNING** - Problemas que pueden causar errores en el futuro
- ℹ️ **INFO** - Buenas prácticas, no crítico

---

## ⚡ Optimización y Rendimiento

### Factores que Afectan la Velocidad

| Factor | Impacto | Recomendación |
|--------|---------|---------------|
| **Ancho de banda de red** | 70% | Usar conexión de 500+ Mbps, cable Ethernet (no WiFi) |
| **Latencia de red** | 20% | Ejecutar desde red cercana a servidores de Microsoft (ideal: VM en Azure) |
| **Límites de SharePoint** | 10% | Respetar THROTTLE_LIMIT=10 para evitar error 429 |
| **CPU/RAM del equipo** | 5% | Equipo moderno con 16GB+ RAM |

### Velocidades Esperadas

| Escenario | Red | Velocidad Estimada |
|-----------|-----|-------------------|
| **Red lenta** | 10 Mbps | ~20 archivos/min |
| **Red media** | 100 Mbps | ~50 archivos/min |
| **Red rápida** | 500 Mbps | ~70 archivos/min |
| **Red empresarial** | 1 Gbps | ~100 archivos/min |
| **VM en Azure** | Misma red que SharePoint | ~150 archivos/min |

### Optimización para Migraciones Grandes (1TB+)

**Opción 1: VM en Azure (Más rápida)**
1. Crear VM Windows en Azure (misma región que SharePoint)
2. Copiar scripts y certificado a la VM
3. Ejecutar migración desde la VM
4. **Ventaja:** Latencia mínima + ancho de banda gigante
5. **Costo:** ~$100-200 USD por migración completa

**Opción 2: Ejecución en horario no laboral**
1. Ejecutar de noche/madrugada
2. Red libre sin competencia
3. **Ventaja:** Gratis, usa infraestructura existente

**Opción 3: Procesamiento por lotes**
1. Migrar carpetas específicas con `-ScopeFolder`
2. Ejecutar múltiples veces para diferentes carpetas
3. **Ventaja:** Control granular, fácil de reanudar

---

## ❗ Solución de Problemas

### Errores Comunes

| Error | Causa | Solución |
|-------|-------|----------|
| `Access is denied` | OneDrive no accesible | Abrir OneDrive del usuario en navegador como admin |
| `Certificate not found` | Certificado no en carpeta correcta | Verificar que `.pfx` está en misma carpeta que script |
| `Unauthorized` | Permisos insuficientes | Verificar permisos de la app en Entra ID |
| `429 Too Many Requests` | Throttling de SharePoint | Ya manejado automáticamente con reintentos |
| `No connection to disconnect` | Error de desconexión | Ignorar - no afecta la migración |
| `Server relative urls must start with...` | Error de ruta | Ya corregido en V4 |

### Archivos con Nombres Problemáticos

SharePoint rechaza algunos caracteres: `" * : < > ? / \ |`

**Solución:**
1. Ejecutar `validate-filenames.ps1` antes de migrar
2. Revisar reporte CSV de archivos problemáticos
3. Renombrar archivos en OneDrive
4. Re-ejecutar migración con `-SkipExisting`

### Archivos Bloqueados en SharePoint

Si archivos quedan bloqueados (checked out):

```powershell
.\unlock-sharepoint-files.ps1
```

### Limpiar Biblioteca de Pruebas

Para vaciar una biblioteca de SharePoint sin eliminarla:

```powershell
.\delete-test-folder.ps1 -FolderToClean "DestinationLibrary"
```

---

## 🔐 Seguridad

### ⚠️ ARCHIVOS SENSIBLES - NO COMPARTIR

Los siguientes archivos contienen información confidencial y **NUNCA** deben subirse a control de versiones:

- ❌ `.env` - Contiene todas las credenciales y configuración
- ❌ `*.pfx` - Certificados de autenticación
- ❌ `Logs/*.log` - Pueden contener información sensible
- ❌ `Logs/*.csv` - Pueden contener nombres de archivos privados

### ✅ El archivo `.gitignore` ya está configurado

Si usas Git, estos archivos ya están excluidos automáticamente.

### Sobre el Certificado

**¿Está atado a un equipo específico?**
No. El certificado está en el archivo `.pfx`, no en el hardware. Funciona en cualquier equipo donde esté el `.pfx` y el `.env`.

**¿Cuándo vence?**
El certificado tiene vigencia de 2 años. Cuando venza, generar uno nuevo y subirlo a Entra ID.

**¿Es seguro?**
El archivo `.pfx` está protegido con contraseña definida en `.env`. Sin esa contraseña no puede usarse.

**IMPORTANTE:**
- Tratar `.pfx` y `.env` como contraseñas
- No compartir ni dejar en carpetas públicas
- No sincronizar con OneDrive/Dropbox
- No enviar por email sin cifrar

---

## 🤝 Contribuciones

Al contribuir a este proyecto:

- ✅ Actualizar documentación cuando agregues funcionalidades
- ✅ Usar `.env.example` para documentar nuevas variables
- ✅ Probar con datos de prueba antes de producción
- ❌ **NUNCA** incluir credenciales reales en commits
- ❌ **NUNCA** subir archivos `.env` o `.pfx`

---

## 📞 Soporte

Para más información:

- 🐛 Reportar problemas: [Crear issue en el repositorio](https://github.com/Ariaz39/sharepoint-from-onedrive-migration-tool/issues)
- 💡 Sugerencias: Contribuir con pull requests
- 📧 Contacto: [LinkedIn - Alejandro Ariaz](https://www.linkedin.com/in/alejandro-ariaz/)

---

## 👨‍💻 Autor

**Desarrollado por:** Alejandro Ariaz
**LinkedIn:** [linkedin.com/in/alejandro-ariaz](https://www.linkedin.com/in/alejandro-ariaz/)
**GitHub:** [@Ariaz39](https://github.com/Ariaz39)

Si este proyecto te fue útil, no dudes en dar una ⭐ al repositorio o compartirlo.

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más detalles.

**¿Qué significa?** Puedes usar, modificar y distribuir este código libremente, incluso para proyectos comerciales. Solo se requiere mantener el aviso de copyright.

---

**Sistema de Migración OneDrive a SharePoint**
*Herramienta optimizada para migración de grandes volúmenes de datos*
