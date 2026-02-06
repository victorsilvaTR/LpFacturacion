# Documentación LPFacturacion

Sistema de facturación TR Facturación desarrollado en Visual Basic 6.0.

## Índice

- [Estructura del Proyecto](estructura-proyecto.md)
- [Guía de Compilación](guia-compilacion.md)
- [Pipeline CI/CD](pipeline-cicd.md)
- [Instalador](instalador.md)
- [Configuración Git](configuracion-git.md)

## Versión Actual

**Versión:** 2.0.3
**Fecha:** 29 Feb 2024
**Repositorio:** https://github.com/victorsilvaTR/LpFacturacion

## Proyectos VB6

| Proyecto | Archivo | Ejecutable | Descripción |
|----------|---------|------------|-------------|
| NetCode | `Facturacion/NetCode/Project1.vbp` | NetCodePrueba1.exe | Códigos de Red Contabilidad |
| LPFacturacion | `Facturacion/LPFacturacion/LPFacturacion.vbp` | TRFacturacion.exe | Sistema principal de facturación |

## Componentes Requeridos

### Controles OCX
- FlexEdGrid2.ocx - Control de grilla editable v2
- FlexEdGrid3.ocx - Control de grilla editable v3

### DLLs
- FwZip32.dll - Librería de compresión ZIP
- FairDll32.dll - Librería Fairware

### Referencias
- DAO 3.6 - Data Access Objects
- ADO 2.8 - ActiveX Data Objects
- MSXML 3.0 - Microsoft XML Parser
- Scripting Runtime - Microsoft Scripting

## Inicio Rápido

### Compilación Manual
```cmd
# 1. Registrar controles OCX (como Administrador)
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid2.ocx"
regsvr32 "C:\lpfacturacion-repo\libs\FlexEdGrid3.ocx"

# 2. Abrir VB6 y compilar proyectos
# - Abrir Facturacion\NetCode\Project1.vbp
# - Abrir Facturacion\LPFacturacion\LPFacturacion.vbp
# - File → Make [proyecto].exe
```

### Compilación Automática (CI/CD)
```bash
# Hacer cambios en el código
git add .
git commit -m "Descripción de cambios"
git push

# El pipeline compila automáticamente y genera artifacts
```

## Soporte

Para problemas o consultas, crear un issue en el repositorio de GitHub.
