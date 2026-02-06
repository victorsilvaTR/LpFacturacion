# Guía de Compilación

## Requisitos Previos

### Software Necesario

| Software | Versión | Ubicación |
|----------|---------|-----------|
| Visual Basic 6.0 | Enterprise/Professional | `C:\Program Files (x86)\Microsoft Visual Studio\VB98\` |
| Inno Setup | 6.x | `C:\Program Files (x86)\Inno Setup 6\` |
| Git | Última | - |

### Controles OCX Requeridos

Los controles deben estar registrados en el sistema:

```cmd
# Ejecutar como Administrador
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid2.ocx"
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid3.ocx"
```

**Verificar registro exitoso:**
- Mensaje: "DllRegisterServer in [archivo].ocx succeeded"

## Secuencia de Compilación

### 1. NetCode (Proyecto Auxiliar)

**Proyecto:** `Facturacion\NetCode\Project1.vbp`

**Pasos:**
1. Abrir VB6
2. File → Open Project → `Project1.vbp`
3. Verificar referencias: Project → References
   - ✅ DAO 2.5/3.51
   - ✅ ADO 2.8
4. Verificar componentes: Project → Components
   - ✅ TABCTL32.OCX
   - ✅ MSFLXGRD.OCX
5. File → Make Project1.exe

**Output:** `NetCodePrueba1.exe` (aprox. 148 KB)

### 2. LPFacturacion (Proyecto Principal)

**Proyecto:** `Facturacion\LPFacturacion\LPFacturacion.vbp`

**Pasos:**
1. Abrir VB6
2. File → Open Project → `LPFacturacion.vbp`
3. Verificar referencias: Project → References
   - ✅ DAO 3.6
   - ✅ ADO 2.8
   - ✅ MSXML 3.0
   - ✅ Scripting Runtime
4. Verificar componentes: Project → Components
   - ✅ FlexEdGrid2.ocx ⚠️ (debe estar registrado)
   - ✅ FlexEdGrid3.ocx ⚠️ (debe estar registrado)
   - ✅ MSFLXGRD.OCX
   - ✅ TABCTL32.OCX
   - ✅ COMDLG32.OCX
   - ✅ MSCOMCTL.OCX
5. File → Make TRFacturacion.exe

**Output:** `TRFacturacion.exe` (aprox. 3.1 MB)

## Compilación del Instalador

### Preparar Archivos

1. **Copiar ejecutables compilados:**
   ```cmd
   copy TRFacturacion.exe C:\lpfacturacion-repo\installer\Archivos\
   ```

2. **Verificar estructura:**
   ```
   installer/
   ├── Archivos/
   │   ├── TRFacturacion.exe     ✅
   │   ├── LicenciaLPContab.pdf  ✅
   │   └── Datos/*.mdb            ✅
   ├── Images/
   │   ├── OpenFact.ico           ✅
   │   ├── TRFactura1.bmp         ✅
   │   └── TRFacturaIco.bmp       ✅
   └── installer.iss              ✅
   ```

### Compilar con Inno Setup

**Método 1: GUI**
1. Abrir Inno Setup Compiler
2. File → Open → `installer\installer.iss`
3. Build → Compile

**Método 2: Línea de comandos**
```cmd
cd C:\lpfacturacion-repo\installer
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

**Output:** `TRFactura_2.0.3_240229.exe` (aprox. 3.1 MB)

## Compilación por Línea de Comandos (VB6)

```cmd
set VB6="C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"

REM Compilar NetCode
%VB6% /make "C:\lpfacturacion-repo\Facturacion\NetCode\Project1.vbp"

REM Compilar LPFacturacion
%VB6% /make "C:\lpfacturacion-repo\Facturacion\LPFacturacion\LPFacturacion.vbp"
```

**Opciones útiles:**
- `/make` - Compilar proyecto
- `/outdir [ruta]` - Especificar directorio de salida
- `/makedll` - Compilar como DLL (no aplicable aquí)

## Verificación Line Endings (CRÍTICO)

⚠️ **Los archivos VB6 DEBEN tener line endings CRLF (Windows)**

### Verificar archivos:
```bash
file archivo.frm
# Debe decir: "ISO-8859 text, with CRLF line terminators"
```

### Si están corruptos (LF):
```bash
# Convertir a CRLF
unix2dos archivo.frm
```

### Protección automática:
El archivo `.gitattributes` con `* binary` previene este problema.

## Problemas Comunes

### Error: "File not found" en referencias

**Causa:** Referencias faltantes o rutas incorrectas

**Solución:**
1. Project → References
2. Desmarcar referencias con "MISSING"
3. Browse y reactivar la referencia correcta

### Error: "Object library not registered"

**Causa:** OCX no registrados

**Solución:**
```cmd
# Como Administrador
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid2.ocx"
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid3.ocx"
```

### Error: "Path/File access error"

**Causa:** Rutas relativas incorrectas

**Solución:**
- Verificar que existan `VB50/`, `Contabilidad70/`
- Abrir proyecto desde la ubicación correcta

### Error: "License information not found"

**Causa:** OCX sin licencia en tiempo de diseño

**Solución:**
- Usar versión completa del OCX
- Verificar que el .oca esté presente

### Instalador: "Bitmap image is not valid"

**Causa:** Imágenes en formato incorrecto

**Solución:**
- Verificar que las imágenes sean BMP válidos
- Script debe referenciar archivos .bmp (no .jpg)

### Instalador: Caracteres corruptos (tildes)

**Causa:** Codificación incorrecta del archivo .iss

**Solución:**
- Archivo debe estar en UTF-8 con BOM
- Verificar caracteres especiales en el script

## Actualizar Versión del Instalador

Editar `installer/installer.iss`:

```ini
#define MyExeSetup "TRFactura_2.0.4_260315"
#define MyAppDate "15 Mar 2026"
#define MyAppVersion "2.0.4"
```

## Compilación Limpia

Para asegurar compilación desde cero:

```cmd
REM Eliminar archivos temporales VB6
del /s *.vbw
del /s *.log

REM Recompilar
%VB6% /make proyecto.vbp
```

## Output y Artifacts

### Ejecutables VB6
- `NetCodePrueba1.exe` - 148 KB
- `TRFacturacion.exe` - 3.1 MB

### Instalador
- `TRFactura_2.0.3_240229.exe` - 3.1 MB

### Ubicación (manual)
Los ejecutables se generan en el mismo directorio que el .vbp

### Ubicación (CI/CD)
- Build: `C:\lpfacturacion-build\LpFacturacion\installer\Archivos\`
- Artifacts: Disponibles en GitHub Actions
