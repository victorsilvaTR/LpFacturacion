# Pipeline CI/CD - GitHub Actions

## Descripción

Pipeline automatizado para compilar proyectos VB6 y generar instalador usando GitHub Actions con self-hosted runner.

**Archivo:** `.github/workflows/build.yml`

## Triggers (Cuándo se Ejecuta)

### Push a rama principal
```yaml
on:
  push:
    branches: [ master, main ]
```

### Solo cuando cambian archivos relevantes
```yaml
paths:
  - '**.vbp'      # Proyectos VB6
  - '**.frm'      # Formularios
  - '**.bas'      # Módulos
  - '**.cls'      # Clases
  - '**.ctl'      # Controles de usuario
  - 'installer/**'
  - '.github/workflows/build.yml'
```

### Pull Requests
```yaml
pull_request:
  branches: [ master, main ]
```

### Manual (Workflow Dispatch)
- Desde GitHub: Actions → Build and Package → Run workflow

## Configuración del Runner

### Self-Hosted Runner
```yaml
runs-on: self-hosted
```

**Requisitos del runner:**
- Windows con VB6 instalado
- Inno Setup 6 instalado
- Configurado en GitHub Settings → Actions → Runners

### Variables de Entorno
```yaml
env:
  VB6: 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE'
  ISCC: 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
  BUILD_DIR: 'C:\lpfacturacion-build'
```

## Pipeline Steps

### 1. Checkout
```yaml
- name: Checkout
  uses: actions/checkout@v4
```

**Función:** Descarga el código del repositorio

### 2. Prepare build directory
```yaml
- name: Prepare build directory
  shell: cmd
  run: |
    if exist "%BUILD_DIR%\LpFacturacion" (
      echo Removing existing build directory...
      rd /s /q "%BUILD_DIR%\LpFacturacion" 2>nul
      timeout /t 2 /nobreak >nul
    )
    echo Copying files to build directory...
    robocopy "%GITHUB_WORKSPACE%" "%BUILD_DIR%\LpFacturacion" /E /NFL /NDL /NJH /NJS /NC /NS /XD .git
```

**Función:**
- Limpia directorio de build anterior
- Copia archivos del workspace a `C:\lpfacturacion-build`
- Excluye carpeta `.git`

**Flags robocopy:**
- `/E` - Copia subdirectorios (incluye vacíos)
- `/NFL` - No lista archivos
- `/NDL` - No lista directorios
- `/NJH /NJS` - No muestra headers/summary
- `/NC /NS` - No muestra clase/tamaño de archivos
- `/XD .git` - Excluye carpeta .git

### 3. Build VB6 projects
```yaml
- name: Build VB6 projects
  shell: cmd
  run: |
    echo === Compilando proyectos VB6 ===

    echo [1/2] NetCodePrueba1.exe
    "%VB6%" /make "%BUILD_DIR%\LpFacturacion\Facturacion\NetCode\Project1.vbp" /outdir "%BUILD_DIR%\LpFacturacion\installer\Archivos"
    if errorlevel 1 exit /b 1

    echo [2/2] TRFacturacion.exe
    "%VB6%" /make "%BUILD_DIR%\LpFacturacion\Facturacion\LPFacturacion\LPFacturacion.vbp" /outdir "%BUILD_DIR%\LpFacturacion\installer\Archivos"
    if errorlevel 1 exit /b 1

    dir "%BUILD_DIR%\LpFacturacion\installer\Archivos\*.exe"
```

**Función:**
- Compila NetCode → `NetCodePrueba1.exe`
- Compila LPFacturacion → `TRFacturacion.exe`
- Output a: `C:\lpfacturacion-build\LpFacturacion\installer\Archivos\`
- Falla el build si algún proyecto falla

**Secuencia de compilación:**
1. NetCode (más simple, menos dependencias)
2. LPFacturacion (proyecto principal)

### 4. Build Installer
```yaml
- name: Build Installer
  shell: cmd
  run: |
    echo === Compilando instalador con Inno Setup ===
    "%ISCC%" "%BUILD_DIR%\LpFacturacion\installer\installer.iss"
    if errorlevel 1 exit /b 1

    dir "%BUILD_DIR%\LpFacturacion\installer\*.exe"
```

**Función:**
- Compila `installer.iss` con Inno Setup
- Genera: `TRFactura_2.0.3_240229.exe`
- Falla el build si la compilación falla

### 5. Upload VB6 executables
```yaml
- name: Upload VB6 executables
  uses: actions/upload-artifact@v4
  with:
    name: vb6-executables-${{ github.sha }}
    path: |
      C:\lpfacturacion-build\LpFacturacion\installer\Archivos\NetCodePrueba1.exe
      C:\lpfacturacion-build\LpFacturacion\installer\Archivos\TRFacturacion.exe
    retention-days: 30
```

**Función:**
- Sube ejecutables VB6 como artifacts
- Nombre único por commit (SHA)
- Retención: 30 días

### 6. Upload Installer
```yaml
- name: Upload Installer
  uses: actions/upload-artifact@v4
  with:
    name: installer-${{ github.sha }}
    path: C:\lpfacturacion-build\LpFacturacion\installer\TRFactura*.exe
    retention-days: 30
```

**Función:**
- Sube instalador como artifact
- Nombre único por commit (SHA)
- Retención: 30 días

## Artifacts Generados

### vb6-executables-[sha]
**Contenido:**
- `NetCodePrueba1.exe` (148 KB)
- `TRFacturacion.exe` (3.1 MB)

**Ubicación:** GitHub → Actions → [run] → Artifacts

### installer-[sha]
**Contenido:**
- `TRFactura_2.0.3_240229.exe` (3.1 MB)

**Ubicación:** GitHub → Actions → [run] → Artifacts

## Descargar Artifacts

### Desde GitHub Web
1. Ir a: https://github.com/victorsilvaTR/LpFacturacion/actions
2. Click en el workflow run deseado
3. Scroll down a "Artifacts"
4. Click en el artifact para descargar

### Usando GitHub CLI
```bash
# Listar artifacts
gh run list --limit 5

# Descargar artifacts del último run
gh run download
```

## Configurar Self-Hosted Runner

### Paso 1: Agregar Runner en GitHub
1. Repository → Settings → Actions → Runners
2. Click "New self-hosted runner"
3. Seleccionar Windows
4. Seguir instrucciones

### Paso 2: Instalar en Windows
```powershell
# Descargar y extraer
mkdir actions-runner; cd actions-runner
Invoke-WebRequest -Uri https://github.com/actions/runner/releases/download/v[version]/actions-runner-win-x64-[version].zip -OutFile actions-runner-win-x64.zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD/actions-runner-win-x64.zip", "$PWD")

# Configurar
./config.cmd --url https://github.com/victorsilvaTR/LpFacturacion --token [TOKEN]

# Ejecutar
./run.cmd
```

### Paso 3: Ejecutar como Servicio (Opcional)
```powershell
# Instalar servicio
./svc.cmd install

# Iniciar servicio
./svc.cmd start
```

### Paso 4: Verificar
1. Settings → Actions → Runners
2. El runner debe aparecer como "Idle" (verde)

## Monitoreo del Pipeline

### Ver Logs en Tiempo Real
1. GitHub → Actions → [running workflow]
2. Click en el job "Build VB6 and Installer"
3. Ver logs de cada step

### Notificaciones
- Email automático en caso de fallo
- Configurar en: Settings → Notifications

## Troubleshooting

### Build falla en "Build VB6 projects"

**Posibles causas:**
- OCX no registrados en el runner
- Referencias faltantes
- Rutas incorrectas

**Solución:**
```cmd
# Registrar OCX en el runner
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid2.ocx"
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid3.ocx"
```

### Build falla en "Build Installer"

**Posibles causas:**
- Inno Setup no instalado
- Ejecutables VB6 no generados
- Error en installer.iss

**Verificar:**
- Inno Setup instalado en: `C:\Program Files (x86)\Inno Setup 6\`
- Step anterior completado exitosamente

### Runner offline

**Causas:**
- Servicio detenido
- Máquina apagada
- Configuración incorrecta

**Solución:**
```powershell
# Verificar servicio
Get-Service actions.runner.*

# Reiniciar servicio
Restart-Service actions.runner.*
```

### Artifacts no se suben

**Causa:** Rutas incorrectas

**Verificar:**
- Los archivos existen en `C:\lpfacturacion-build\...`
- Permisos de lectura en el directorio

## Flujo Completo

```
1. Developer hace commit/push
   ↓
2. GitHub detecta cambios en archivos .vbp/.frm/.bas/etc
   ↓
3. Workflow se activa en self-hosted runner
   ↓
4. Checkout del código
   ↓
5. Copia a C:\lpfacturacion-build
   ↓
6. Compila NetCode
   ↓
7. Compila LPFacturacion
   ↓
8. Compila instalador con Inno Setup
   ↓
9. Sube artifacts a GitHub
   ↓
10. Notifica resultado (éxito/fallo)
```

## Tiempo Estimado de Ejecución

| Step | Tiempo |
|------|--------|
| Checkout | ~5s |
| Prepare build | ~10s |
| Build VB6 | ~30s |
| Build Installer | ~15s |
| Upload artifacts | ~10s |
| **Total** | **~70s** |

## Variables de GitHub

Disponibles automáticamente en el workflow:

| Variable | Ejemplo | Descripción |
|----------|---------|-------------|
| `${{ github.sha }}` | `36ee8d1...` | SHA del commit |
| `${{ github.ref }}` | `refs/heads/main` | Referencia del branch |
| `${{ github.actor }}` | `victorsilvaTR` | Usuario que hizo push |
| `${{ github.workspace }}` | `C:\actions-runner\_work\...` | Directorio de trabajo |

## Personalización del Pipeline

### Cambiar versión del instalador

Editar `.github/workflows/build.yml`:

```yaml
- name: Update version
  shell: cmd
  run: |
    REM Actualizar versión en installer.iss antes de compilar
```

### Agregar tests

```yaml
- name: Run tests
  shell: cmd
  run: |
    REM Ejecutar tests aquí
```

### Notificaciones Slack/Discord

```yaml
- name: Notify Slack
  uses: slackapi/slack-github-action@v1
  with:
    webhook-url: ${{ secrets.SLACK_WEBHOOK }}
```

## Seguridad

### Secrets
Guardar información sensible en GitHub Secrets:
- Settings → Secrets and variables → Actions → New repository secret

### Permisos del Runner
El runner tiene acceso completo al sistema. Asegurar:
- Runner en máquina dedicada/VM
- Firewall configurado
- Antivirus actualizado
