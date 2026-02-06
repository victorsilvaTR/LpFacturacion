# Estructura del Proyecto

## Árbol de Directorios

```
LpFacturacion/
├── .github/
│   └── workflows/
│       └── build.yml              # Pipeline CI/CD de GitHub Actions
├── .gitattributes                 # Configuración Git para archivos binarios
├── docs/                          # Documentación del proyecto
├── Facturacion/
│   ├── Contabilidad70/
│   │   ├── Administrador/         # Formularios compartidos
│   │   ├── Comun/                 # Recursos comunes
│   │   └── HyperContabilidad/     # Módulos de contabilidad
│   ├── LPFacturacion/             # Proyecto principal
│   │   ├── Images/                # Recursos gráficos
│   │   ├── Docs/                  # Documentación técnica
│   │   ├── Exportar/              # Módulos de exportación
│   │   ├── Importar/              # Módulos de importación
│   │   ├── LPFacturacion.vbp      # Proyecto VB6 principal
│   │   ├── LPFacturacion.vbw      # Workspace VB6
│   │   ├── FrmMain.frm            # Formulario principal
│   │   └── *.frm, *.bas, *.cls    # Código fuente
│   ├── NetCode/                   # Proyecto auxiliar
│   │   ├── Project1.vbp           # Proyecto VB6 NetCode
│   │   ├── FrmMain.frm            # Formulario principal
│   │   └── ModNetCode.bas         # Módulo principal
│   ├── VB50/                      # Módulos compartidos
│   │   ├── PAM.bas                # Módulo PAM
│   │   ├── SDK50.bas              # SDK
│   │   ├── Franca.bas             # Utilidades
│   │   └── *.bas, *.cls, *.frm    # Otros módulos compartidos
│   └── RUTS/                      # Datos de RUTs
├── installer/
│   ├── Archivos/
│   │   ├── Datos/
│   │   │   ├── TRFactura.mdb           # Base de datos principal
│   │   │   ├── TRFacturaDemo.mdb       # Base de datos demo
│   │   │   └── EmpresaVacia-DTE.mdb    # Base vacía con DTE
│   │   ├── TRFacturacion.exe           # Ejecutable compilado
│   │   └── LicenciaLPContab.pdf        # Licencia
│   ├── Images/
│   │   ├── OpenFact.ico                # Icono principal
│   │   ├── TRFactura1.bmp              # Imagen wizard grande
│   │   └── TRFacturaIco.bmp            # Imagen wizard pequeña
│   ├── installer.iss                   # Script Inno Setup
│   └── TRFactura_2.0.3_240229.exe      # Instalador compilado
├── libs/
│   ├── FlexEdGrid2.ocx            # Control OCX v2
│   ├── FlexEdGrid2.oca            # Type library cache
│   ├── FlexEdGrid3.ocx            # Control OCX v3
│   ├── FwZip32.dll                # DLL compresión ZIP
│   └── FairDll32.dll              # DLL Fairware
└── README.md                      # README principal del repositorio
```

## Directorios Principales

### `.github/workflows/`
Contiene los workflows de GitHub Actions para CI/CD.

**Archivos:**
- `build.yml` - Pipeline de compilación automática

### `Facturacion/`
Código fuente de los proyectos VB6.

**Subdirectorios principales:**
- `LPFacturacion/` - Sistema principal de facturación
- `NetCode/` - Códigos de red
- `VB50/` - Módulos compartidos entre proyectos
- `Contabilidad70/` - Recursos compartidos de contabilidad

### `installer/`
**Nueva carpeta** agregada para gestionar el proceso de instalación.

**Contenido:**
- `installer.iss` - Script de Inno Setup
- `Archivos/` - Ejecutables y archivos a instalar
- `Images/` - Imágenes del wizard de instalación
- `TRFactura_*.exe` - Instalador compilado

### `libs/`
**Nueva carpeta** con controles OCX y DLLs requeridos.

**Contenido:**
- Controles OCX personalizados (FlexEdGrid2, FlexEdGrid3)
- DLLs auxiliares (FwZip32, FairDll32)

## Archivos de Configuración

### `.gitattributes`
**Nuevo archivo** crítico para proteger archivos VB6.

```
* binary
```

Configura Git para tratar TODOS los archivos como binarios, evitando conversión de line endings que corrompe archivos VB6.

### `installer.iss`
Script de Inno Setup para generar el instalador.

**Configuración principal:**
- Versión: 2.0.3
- Nombre: TR Facturación
- Ejecutable: TRFacturacion.exe
- Directorio destino: `HR\TRFactura`
- Requisitos: Windows 7 o superior (MinVersion 6.1)

## Dependencias entre Proyectos

```
LPFacturacion (Principal)
├── Usa: VB50/* (módulos compartidos)
├── Usa: Contabilidad70/Comun/* (formularios)
├── Usa: Contabilidad70/Administrador/* (formularios)
├── Usa: Contabilidad70/HyperContabilidad/* (módulos)
└── Requiere: FlexEdGrid2.ocx, FlexEdGrid3.ocx

NetCode (Auxiliar)
├── Usa: VB50/* (módulos compartidos)
└── Independiente de LPFacturacion
```

## Archivos Importantes

| Archivo | Descripción | Ubicación |
|---------|-------------|-----------|
| LPFacturacion.vbp | Proyecto principal VB6 | `Facturacion/LPFacturacion/` |
| Project1.vbp | Proyecto NetCode VB6 | `Facturacion/NetCode/` |
| installer.iss | Script Inno Setup | `installer/` |
| build.yml | Pipeline CI/CD | `.github/workflows/` |
| .gitattributes | Configuración Git | Raíz del repositorio |

## Cambios Recientes

### Carpetas Agregadas
- ✅ `.github/workflows/` - Pipeline CI/CD
- ✅ `installer/` - Instalador y recursos
- ✅ `libs/` - Controles OCX y DLLs
- ✅ `docs/` - Documentación

### Archivos Agregados
- ✅ `.gitattributes` - Protección binaria
- ✅ `installer/installer.iss` - Script instalador
- ✅ `installer/TRFactura_2.0.3_240229.exe` - Instalador compilado
- ✅ `.github/workflows/build.yml` - Pipeline build

### Configuración Migrada
- ✅ Repositorio migrado a: `https://github.com/victorsilvaTR/LpFacturacion`
- ✅ Build directory configurado: `C:\lpfacturacion-build`
