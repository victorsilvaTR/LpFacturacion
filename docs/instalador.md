# Instalador TR Facturación

## Descripción

Instalador generado con **Inno Setup 6** para el sistema TR Facturación.

**Archivo:** `TRFactura_2.0.3_240229.exe` (3.1 MB)

## Características

- **Producto:** TR Facturación
- **Versión:** 2.0.3
- **Fecha:** 29 Feb 2024
- **Editor:** Legal Publishing
- **Tipo:** Actualizador (requiere instalación completa previa)
- **Plataforma:** Windows 7 o superior (64/32 bits)

## Contenido del Instalador

### Ejecutables
- `TRFacturacion.exe` (3.1 MB) - Aplicación principal

### DLLs
- `FwZip32.dll` (138 KB) - Compresión ZIP
- `FairDll32.dll` (58 KB) - Librería Fairware

### Documentación
- `LicenciaLPContab.pdf` (396 KB) - Licencia del producto

### Bases de Datos
- `TRFactura.mdb` - Base de datos principal
- `TRFacturaDemo.mdb` - Base demo
- `EmpresaVacia-DTE.mdb` - Base vacía con DTE

## Requisitos del Sistema

### Sistema Operativo
- Windows 7 o superior
- Windows Server 2008 R2 o superior

### Espacio en Disco
- Mínimo: 5 MB
- Recomendado: 20 MB

### Permisos
- Usuario con permisos de escritura en la unidad de instalación
- NO requiere permisos de administrador (PrivilegesRequired=lowest)

## Proceso de Instalación

### 1. Ejecutar Instalador
Doble click en `TRFactura_2.0.3_240229.exe`

### 2. Pantalla de Bienvenida
- Muestra versión y fecha
- Advierte cerrar TR Facturación si está ejecutándose
- Advierte que no haya usuarios conectados

### 3. Selección de Unidad
- Selector desplegable con unidades disponibles (C:, D:, E:, etc.)
- Directorio: `[Unidad]:\HR\TRFactura`
- Ejemplo: `C:\HR\TRFactura`

### 4. Verificaciones
El instalador verifica:
- ✅ Que TR Facturación no esté ejecutándose
- ✅ Que no haya usuarios conectados (archivo `.ldb` no exista)

Si alguna verificación falla, el instalador se detiene.

### 5. Instalación
- Copia archivos
- Reemplaza ejecutable existente
- Actualiza DLLs

### 6. Finalización
- Opción de ejecutar TR Facturación
- Crea iconos en:
  - Menú Inicio
  - Escritorio (opcional)
  - Barra de inicio rápido (opcional, Windows <7)

## Estructura del Script (installer.iss)

### Variables Principales
```ini
#define MyExeSetup "TRFactura_2.0.3_240229"
#define MyAppDate "29 Feb 2024"
#define MyAppVersion "2.0.3"
#define MyAppName "TR Facturación"
#define MyAppPublisher "Legal Publishing"
#define MyAppExeName "TRFacturacion.exe"
```

### Configuración [Setup]
```ini
AppId={{0F04D1A7-22AB-40E9-885B-562A8922DB71}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName=HR\TRFactura
MinVersion=6.1
PrivilegesRequired=lowest
Compression=lzma/ultra
```

### Archivos [Files]
```ini
Source: Archivos\TRFacturacion.exe; DestDir: {app}
Source: ..\libs\FwZip32.dll; DestDir: {app}
Source: ..\libs\FairDll32.dll; DestDir: {app}
Source: Archivos\LicenciaLPContab.pdf; DestDir: {app}
```

### Iconos [Icons]
```ini
Name: {group}\TR Facturación; Filename: {app}\TRFacturacion.exe
Name: {commondesktop}\TR Facturación; Filename: {app}\TRFacturacion.exe
```

### Verificaciones [Code]
```pascal
Function InitializeSetup : Boolean;
begin
  // Verifica que la aplicación no esté ejecutándose
  // FindWindowByWindowName('TR Facturación')
end;

Function NextButtonClick(PageId: Integer): Boolean;
begin
  // Verifica que no haya usuarios conectados (.ldb)
end;
```

## Personalización del Instalador

### Cambiar Versión
Editar `installer/installer.iss` líneas 8-10:

```ini
#define MyExeSetup "TRFactura_2.0.4_260315"
#define MyAppDate "15 Mar 2026"
#define MyAppVersion "2.0.4"
```

### Cambiar Imágenes
Reemplazar archivos en `installer/Images/`:

| Archivo | Dimensiones | Uso |
|---------|-------------|-----|
| OpenFact.ico | 256x256 | Icono del instalador |
| TRFactura1.bmp | 566x551 | Imagen grande del wizard |
| TRFacturaIco.bmp | 477x217 | Imagen pequeña del wizard |

**Formato:** BMP 24-bit, resolución 3780x3780 px/m

### Agregar Archivos
Editar sección `[Files]`:

```ini
Source: Archivos\MiArchivo.txt; DestDir: {app}
```

### Cambiar Directorio Destino
Editar línea 37:

```ini
DefaultDirName=HR\MiDirectorio
```

## Compilar el Instalador

### Método 1: GUI
```
1. Abrir Inno Setup Compiler
2. File → Open → installer\installer.iss
3. Build → Compile
```

### Método 2: Línea de comandos
```cmd
cd C:\lpfacturacion-repo\installer
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

### Método 3: Pipeline CI/CD
```bash
git push
# El pipeline compila automáticamente
```

## Desinstalación

### Desde Panel de Control
1. Panel de Control → Programas → Desinstalar un programa
2. Buscar "TR Facturación"
3. Click en Desinstalar

### Desde Menú Inicio
1. Menú Inicio → TR Facturación
2. Click en "Uninstall TR Facturación"

### Manual
Ejecutar: `[DirectorioInstalación]\Uninstall\unins000.exe`

## Logs de Instalación

Inno Setup genera logs en:
```
%TEMP%\Setup Log YYYY-MM-DD #001.txt
```

**Contenido:**
- Fecha/hora de instalación
- Versión instalada
- Archivos copiados
- Errores (si hay)

## Problemas Comunes

### "La aplicación se debe cerrar antes de proceder"

**Causa:** TR Facturación está ejecutándose

**Solución:** Cerrar todas las instancias de TRFacturacion.exe

### "Hay al menos un usuario conectado"

**Causa:** Archivo `TRFactura.ldb` existe (usuario conectado a la base)

**Solución:**
1. Cerrar TR Facturación en todas las máquinas
2. Verificar que no haya procesos colgados
3. Eliminar manualmente `HR\TRFactura\Datos\TRFactura.ldb` si es necesario

### "Bitmap image is not valid"

**Causa:** Imágenes del wizard en formato incorrecto

**Solución:**
- Verificar que las imágenes sean BMP (no JPG)
- Recompilar el instalador

### "Error writing to file"

**Causa:** Permisos insuficientes

**Solución:**
- Ejecutar instalador como Administrador
- Verificar permisos de escritura en la unidad destino

### Instalador no se ejecuta

**Causa:** Bloqueado por Windows SmartScreen

**Solución:**
1. Click derecho → Propiedades
2. Marcar "Desbloquear"
3. Aplicar

## Distribución

### Interno (Red Local)
Copiar a carpeta compartida:
```
\\servidor\instaladores\TRFactura_2.0.3_240229.exe
```

### Externo (Descarga)
- Subir a servidor web/FTP
- Compartir enlace de descarga

### Artifacts GitHub
Descargar desde:
```
https://github.com/victorsilvaTR/LpFacturacion/actions
```

## Versionamiento

Esquema de nombres:
```
TRFactura_[MAJOR].[MINOR].[PATCH]_[YYMMDD].exe

Ejemplo: TRFactura_2.0.3_240229.exe
         └─┬──┘ └┬┘ └┬┘ └─┬──┘
           │     │   │     └─ Fecha (29 Feb 2024)
           │     │   └─────── Patch 3
           │     └─────────── Minor 0
           └───────────────── Major 2
```

## Firmar el Instalador (Opcional)

Para evitar advertencias de SmartScreen:

```cmd
signtool sign /f certificado.pfx /p password /t http://timestamp.verisign.com/scripts/timstamp.dll TRFactura_2.0.3_240229.exe
```

**Requisitos:**
- Certificado de firma de código
- SignTool (Windows SDK)

## Checksum para Verificación

Generar hash del instalador:

```powershell
Get-FileHash TRFactura_2.0.3_240229.exe -Algorithm SHA256
```

Publicar el hash junto al instalador para verificación de integridad.
