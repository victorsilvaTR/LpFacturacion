# Configuración Git

## .gitattributes - Protección de Archivos VB6

### Problema: Line Endings Corruptos

Visual Basic 6 **requiere** archivos con line endings CRLF (Windows). Git por defecto puede convertir archivos a LF (Unix), lo que **corrompe completamente** los archivos VB6.

**Síntoma:**
```
Error: '0' could not be loaded - Line 0: The file X.frm could not be loaded
```

### Solución: Archivos Binarios

**Archivo:** `.gitattributes`

```
# Tratar todos los archivos como binarios para evitar conversión de line endings
* binary
```

**Efecto:**
- Git trata **TODOS** los archivos como binarios
- NO convierte line endings (CRLF ↔ LF)
- NO modifica archivos al hacer checkout/commit
- Protege archivos VB6, imágenes, ejecutables, etc.

### Verificar Line Endings

**Linux/Git Bash:**
```bash
file archivo.frm
# Correcto: "ISO-8859 text, with CRLF line terminators"
# Incorrecto: "ISO-8859 text" (sin CRLF)
```

**Windows (PowerShell):**
```powershell
Format-Hex archivo.frm | Select-String "0D 0A"
# Debe mostrar bytes 0D 0A (CRLF)
```

### Corregir Archivos Corruptos

Si los archivos ya tienen LF en lugar de CRLF:

**Convertir a CRLF:**
```bash
# Unix/Git Bash
find . -type f \( -name "*.frm" -o -name "*.bas" -o -name "*.cls" -o -name "*.vbp" -o -name "*.vbw" \) ! -path "./.git/*" -exec unix2dos {} \;

# Windows (PowerShell)
Get-ChildItem -Recurse -Include *.frm,*.bas,*.cls,*.vbp,*.vbw | ForEach-Object {
    $content = [IO.File]::ReadAllText($_.FullName)
    [IO.File]::WriteAllText($_.FullName, $content)
}
```

## Configuración Git Recomendada

### Desactivar Autocrlf (Global)

```bash
git config --global core.autocrlf false
git config --system core.autocrlf false
```

**¿Por qué?**
- Con `.gitattributes` configurado como `* binary`, autocrlf no debería interferir
- Pero es mejor desactivarlo para total seguridad
- Previene conversiones inesperadas

### Verificar Configuración

```bash
git config --list --show-origin | grep autocrlf
# Debe mostrar: false
```

## Estrategia de Tratamiento por Tipo de Archivo

### Opción 1: Todo Binario (Actual)

```gitattributes
* binary
```

**Pros:**
- Máxima protección
- Simple
- No hay conversiones inesperadas

**Contras:**
- Archivos de texto (README.md, etc.) pueden tener LF en Linux
- Diffs menos legibles para archivos texto

### Opción 2: Selectivo (Alternativa)

```gitattributes
# VB6 - Siempre binario
*.vbp binary
*.frm binary
*.bas binary
*.cls binary
*.ctl binary
*.vbw binary
*.dca binary
*.dsr binary
*.dob binary

# Binarios obvios
*.exe binary
*.dll binary
*.ocx binary
*.oca binary
*.mdb binary
*.ico binary
*.bmp binary
*.jpg binary
*.png binary

# Texto - Normalizar a LF en repo, CRLF en working
*.md text eol=crlf
*.txt text eol=crlf
*.yml text eol=lf
*.yaml text eol=lf
*.sh text eol=lf
```

**Pros:**
- Más control granular
- Diffs legibles para archivos de texto

**Contras:**
- Más complejo
- Riesgo de olvidar extensiones
- No recomendado para repositorios VB6

## .gitignore - Archivos Temporales

**Archivo:** `.gitignore`

```gitignore
# VB6 temporales
*.vbw
*.log

# Builds
*.exe
*.dll
!libs/*.dll
!libs/*.ocx

# Instalador compilado
installer/*.exe
!installer/Archivos/*.exe

# Build directory
C:\lpfacturacion-build/

# IDE
.vs/
*.suo
*.user

# Sistema
Thumbs.db
Desktop.ini
```

**Nota:** Actualmente NO hay `.gitignore` en el repositorio. Los archivos `.vbw` y `.exe` compilados **SÍ** están en el repo para facilitar distribución.

## Repositorio Remoto

### Configuración Actual

```bash
git remote -v
# origin  https://github.com/victorsilvaTR/LpFacturacion.git (fetch)
# origin  https://github.com/victorsilvaTR/LpFacturacion.git (push)
```

### Branch Principal
```bash
git branch
# * main
```

### Migración desde Repositorio Original

El repositorio fue migrado desde:
```
https://github.com/tr/LpFacturacion (original - descartado)
↓
https://github.com/victorsilvaTR/LpFacturacion (actual)
```

**Comando de migración:**
```bash
git remote set-url origin https://github.com/victorsilvaTR/LpFacturacion.git
```

## Workflow de Commits

### 1. Verificar Estado
```bash
git status
```

### 2. Ver Cambios
```bash
git diff
```

### 3. Agregar Archivos
```bash
# Específicos
git add archivo.frm archivo.bas

# Todos
git add .
```

### 4. Commit
```bash
git commit -m "Descripción del cambio

Detalles adicionales si son necesarios.

Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>"
```

### 5. Push
```bash
git push
```

## Protección del Repositorio

### Branch Protection (GitHub)

Configurar en: Settings → Branches → Add rule

**Reglas recomendadas:**
- ✅ Require pull request reviews before merging
- ✅ Require status checks to pass (CI/CD)
- ✅ Require branches to be up to date
- ✅ Include administrators

### Secrets (GitHub)

Para información sensible:
- Settings → Secrets and variables → Actions
- Ejemplo: certificados de firma, credenciales

## Clonar el Repositorio

### Primera vez
```bash
# HTTPS
git clone https://github.com/victorsilvaTR/LpFacturacion.git

# SSH (si tienes configurado)
git clone git@github.com:victorsilvaTR/LpFacturacion.git
```

### Verificar después de clonar
```bash
cd LpFacturacion
file Facturacion/LPFacturacion/LPFacturacion.vbp
# Debe decir: "... with CRLF line terminators"
```

Si NO tiene CRLF, el `.gitattributes` no funcionó correctamente.

## Historial del Repositorio

### Ver commits recientes
```bash
git log --oneline -10
```

### Ver cambios en archivos VB6
```bash
git log --follow -- Facturacion/LPFacturacion/FrmMain.frm
```

### Comparar versiones
```bash
git diff HEAD~5 HEAD -- Facturacion/LPFacturacion/LPFacturacion.vbp
```

## Troubleshooting Git

### "warning: LF will be replaced by CRLF"

**Causa:** autocrlf está activo

**Solución:**
```bash
git config core.autocrlf false
```

### Archivos aparecen modificados después de checkout

**Causa:** Line endings convertidos

**Solución:**
1. Verificar `.gitattributes` existe y tiene `* binary`
2. Rehacer checkout:
```bash
rm .git/index
git reset
git checkout -f
```

### Push rechazado

**Causa:** Branch desactualizado

**Solución:**
```bash
git pull --rebase
git push
```

### Conflictos en archivos VB6

**Evitar:** Los archivos VB6 son difíciles de mergear. Evitar trabajar simultáneamente en el mismo archivo.

**Resolver:**
```bash
# Aceptar versión local
git checkout --ours archivo.frm

# Aceptar versión remota
git checkout --theirs archivo.frm

# Marcar como resuelto
git add archivo.frm
git commit
```

## Best Practices

### ✅ DO
- Usar `.gitattributes` con `* binary`
- Verificar line endings después de clonar
- Hacer commits pequeños y frecuentes
- Escribir mensajes descriptivos
- Hacer pull antes de push

### ❌ DON'T
- NO activar autocrlf
- NO editar archivos VB6 en Linux sin cuidado
- NO commitear archivos compilados innecesarios (opcional)
- NO hacer force push a main
- NO mergear archivos VB6 manualmente si hay conflictos

## Referencias

- Git Documentation: https://git-scm.com/doc
- GitHub Actions: https://docs.github.com/en/actions
- Inno Setup: https://jrsoftware.org/isinfo.php
