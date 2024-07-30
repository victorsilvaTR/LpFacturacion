Attribute VB_Name = "LPFactura"
'Pendientes:    18 ago 16
   ' Manejo de estados de DTE
   ' Restringir cambios en caso de Nota de Crédito
   ' Agregar imp. adicional a productos (?)
   


'Ver el tema de avisar de nueva actualización, función FwWebCommGUIA

'Ver que pasa con FUVER (tabla Param) y con la tabla LParamFact (viene de LParam en Contabilidad)

Option Explicit

'******************* NO MODIFICAR **********************
'Password DB
Public Const SG_SEGCFG = "FW6T9R54WX3A"  'archivo cfg para eliminar clave

'password LexContab.mdb
Public Const PASSW_LEXCONT = "Fw#420!&+"
Public Const PASSW_PREFIX = "Fw#42+"   'prefijo password empresa (sigue RUT sin puntos, ni guión, ni dígito verificador

Public Const FAIRFACT_CODE = 6278547   'Version 1

' Para generar el código de red
Public Const PC_SEED = 637717          'Versión 1

Public Const APP_NAME = "LPFactura"
Public Const APP_URL = "http://www.fairware.cl/TRFactura.asp"
Public Const APP_FULLNAME = "TR_Facturación"

Public Const APP_DEMO = False

'******************* NO MODIFICAR **********************

' Informacion para el archivo de licencias
Public gLicFile As String
Public Const KEY_CRYP = 7827141
Public gCantLicencias As Integer  ' cantidad de licencias autorizadas, se llena en ChkInscPC

' Pam: 13 dic 2010: Licencias
Public Const VER_ILIM = 800   ' Como la actual
Public Const VER_5EMP = 700   ' 5 empresas
Public Const VER_DEMO = 600   ' Demo 3 empresas

'cantidad máxima de comprobantes para la versión Demo
Public Const MAX_FACTDEMO = 20

' *********************************************************************



'nombre aplicación
Public gLPFactura As String


'Nombre base de datos Adm Factura
Public Const BD_COMUNDTE = "TRFactura.mdb"
Public Const BD_COMUNDTEVACIA = "TRFacturaVacia.mdb"

'Nombre Base de Datos Vacia
Public Const BD_VACIADTE = "EmpresaVacia-DTE.mdb"

'clases de impresión
Public gPrtReportes As ClsPrtFlxGrid

'Path de importación/Exportación
Public gImportPath As String
Public gExportPath As String

'path de PDFs de DTE
Public gPdfDTE As String
Public gPdfDTEEmp As String   'con RUT empresa

Type VarIniFile_t
   SelEmprPorRUT As Integer         'Ordenamiento de empresas por RUT en la ventana de selección de empresas
End Type

Public gVarIniFile As VarIniFile_t

'Privilegios
Public Const PRVF_ADM_SIS = &H1&          ' Administrar Sistema (usuarios, perfiles, base de datos)
Public Const PRVF_CFG_EMP = &H2&          ' Configurar empresa
Public Const PRVF_ADM_EMPRESA = &H4&      ' Administrar entidades y productos
Public Const PRVF_ADM_EXP = &H8&          ' Administrar Exportación e Importación de Datos
Public Const PRVF_EMITIR_FACT = &H10&     ' Emitir DTE
Public Const PRVF_ADM_FACT = &H20&        ' Administrar/listar documentos emitidos
Public Const PRVF_MANT_DATOS = &H40&        ' Mantención de datos

Public Const LAST_PRVF = PRVF_MANT_DATOS


'Datos de conexión con Sistema de Facturación electrónica

Public Const PROV_LP = 1
Public Const PROV_ACEPTA = 2

Public gProvFactElect(PROV_ACEPTA) As String

Public Const PUERTO_ACEPTA = 8083  ' OJO: con el puerto de la siguiente URL
Public Const URL_ACEPTA = "http://200.27.203.6:8083/Ws_FacElectronica.asmx"

'Parametros de configuración empresa: proveedor, usuario y clave
   'tipo de parámetro
Public Const CONECT_PROV = "CONECTPROV"
Public Const CONECT_DATA = "CONECTDATA"

   'código parámetro
Public Const CONECT_USUARIO = 1
Public Const CONECT_CLAVE = 2
Public Const CONECT_CLAVECERT = 3
Public Const CONECT_MAILEMISOR = 4
Public Const CONECT_RUTFIRMA = 5

Type ConectData_t
   Proveedor As Integer
   Usuario As String
   Clave As String
   ClaveCert As String
   MailEmisor As String
   RutFirma As String
End Type

Public gConectData As ConectData_t

'Estado de los meses
Global Const EM_NOEXISTE = 0
Global Const EM_ABIERTO = 1
Global Const EM_CERRADO = 2
Global Const EM_ERRONEO = 3
Global Const MAX_ESTADOMES = EM_ERRONEO

Global gEstadoMes(MAX_ESTADOMES)    '"Abierto", "Cerrado", "Erróneo"

Public gAbrirMesesParalelo As Boolean

'Código SII, índice y TipoDoc de Guia de Despacho Electrónica, se pone el Lib_otros
Public Const CODDOCDTESII_GUIADESPACHO = 52     ' código de la Guía de Despacho en el SII
Public Const TIPODOC_GUIADESPACHO = 1000        'TipoDoc de la Guía de Despacho en el arreglo gTipoDocDTE (no está en la lista de TipoDocs de la base de datos)
Public Const IDXTIPODOCDTE_GUIADESPACHO = 0     'indice de la Guía de Despacho en el arreglo gTipoDocDTE

'Tipo de Despacho
Public Const GD_SINDESPACHO = 0        ' Sin Despacho.
Public Const GD_DESPRECEPTOR = 1       ' Despacho por cuenta del receptor del documento (cliente o vendedor en caso de Facturas de compra.)
Public Const GD_DESPEMICLI = 2         ' Despacho por cuenta del emisor a instalaciones del cliente
Public Const GD_DESPEMIOTRO = 3        ' Despacho por cuenta del emisor a otras instalaciones (Ejemplo: entrega en Obra)

Public Const MAX_TIPODESPACHO = GD_DESPEMIOTRO

Public gTipoDespacho(MAX_TIPODESPACHO) As String


'Traslado
Public Const GT_SINTRASLADO = 0        ' Sin Traslado.
Public Const GT_VENTA = 1              ' Operación constituye venta.
Public Const GT_VENTAPORFACT = 2       ' Ventas por efectuar.
Public Const GT_CONSIGNACION = 3       ' Consignaciones.
Public Const GT_ENTGRATIS = 4          ' Entrega gratuita.
Public Const GT_TRASINTERNO = 5        ' Traslados internos.
Public Const GT_OTRONOVENTA = 6        ' Otros traslados no venta.
Public Const GT_DEVOLUCION = 7         ' Guía de devolución.

Public Const MAX_TIPOTRASLADO = GT_DEVOLUCION

Public gTipoTraslado(MAX_TIPOTRASLADO) As String



'Constantes de parámetros de configuración empresa, se almacenan en tabla ParamEmpDTE
Public Const OPTEDFACT_VERCOLCODPROD = 1
Public Const OPTEDFACT_VERCOLDESC = 2
Public Const OPTEDFACT_VERCOLUMED = 3
Public Const OPTEDFACT_VERCOLIMPADIC = 4
Public Const OPTEDFACT_VERREF = 5
Public Const OPTEDFACT_NOTSELPROD = 6
Public Const OPTEDFACT_MODPRECIO = 7
Public Const MAX_OPTEDFACT = OPTEDFACT_MODPRECIO
Public Const MAX_OPTEDFACTVERCOL = OPTEDFACT_VERCOLIMPADIC


Type EmpConfig_t
   OptEdFact(MAX_OPTEDFACT) As Integer      'Opciones new DTE
End Type
Public gEmpConfig As EmpConfig_t


Public Const MAX_DIGITOSCANT = 10   '999.999.999.999
Public Const MAX_DIGITOSVALOR = 15   '999.999.999.999

Public Const MAX_LEN_DESCRIPPROD = 1000   'largo máximo de la descripción de un prodcto, en el XML DTE



'Datos Referencias DTE

' Tipo Referencia(1:Anula Documento de Referencia,2: Corrige Texto Documento de Referencia,3: Corrige montos)
Public Const REF_ANULA = 1
Public Const REF_CORRIGETEXTO = 2
Public Const REF_CORRIGEMONTOS = 3
Public Const MAX_TIPOREF = REF_CORRIGEMONTOS

Public gTipoRefSII(MAX_TIPOREF) As String

'Formas de Pago Factura   '<FmaPago> 1:Contado, 2:Crédito 3:Sin costo (entrega Gratuita)
Public Const FP_CONTADO = 1
Public Const FP_CREDITO = 2
Public Const FP_SINCOSTO = 3

'Estados
Public Const ES_INACTIVO = 0
Public Const ES_ACTIVO = 1

Public gFormaDePago(FP_SINCOSTO) As String

Public gEstado(ES_ACTIVO) As String


'Datos DTE

Public Const MAX_ITEMDTE = 60          'max permitido por el SII: 60, máximo en impresión
Public Const MAX_REFDTE = 40           'max permitido por el SII: 40, máximo en impresión
Public Const MAX_IMPADICDTE = 20       'max permitido por el SII: 20, máximo en impresión: 20?

Public Const RUT_DEFEXPORT = "55555555"   'Rut por omisión en el caso de facturas de exportacíón

'Detalle de un DTE
Type DetDTE_t
   IdDetDTE As Long
   IdDTE As Long
   IdEmpresa As Long
   IdProducto As Long
   TipoCod As String
   CodProd As String
   Producto As String
   Descrip As String
   UMedida As String
   Cantidad As Double
   Precio As Double
   EsExento As Boolean
   IdImpAdic As Long
   CodImpAdicSII As String
   TasaImpAdic As Single
   MontoImpAdic As Double
   DescImpAdic As String
   PjeDescto As Single
   MontoDescto As Double
   SubTotal As Double
End Type

'Referencias de un DTE
Type Referencia_t
   IdReferencia As Long
   IdDTE As Long
   IdEmpresa As Long
   Ano As Integer
   IdTipoDocRef As Long
   CodDocRefSII As String
   FolioRef As String
   FechaRef As Long
   CodRefSII As Integer '(1:Anula Documento de Referencia,2: Corrige Texto Documento de Referencia,3: Corrige montos)
   RazonReferencia As String
End Type

Type DTEImpAdic_t
   IdImpAdic As Integer
   IdImpAdicSII As Long
   TasaImpAdic As Single
   MontoImpAdic As Double
   NetoImpAdic As Double
   DescImpAdic As String
End Type

Type DTEFactExp_t
   IdDTEFactExp As Long
   CodIndServicio As String
   CodPais As String
   CodPuertoEmbarque As String
   CodPuertoDesembarque As String
   CodMoneda As String
   TipoCambio As Single
   TotalBultos As Double
   CodModVenta As String
   CodClausulaVenta As String
   TotClausulaVenta As Double
   CodViaTransporte As String
End Type

Type DTEGuiaDesp_t
   IdDTEGuiaDesp As Long
   Patente As String
   RutChofer As String
   NombreChofer As String
End Type
   

'Encabezado DTE
Type DTE_t
   IdDTE As Long
   IdEmpresa As Long
   Ano As Integer
   TipoDoc As Integer
   TipoLib As Integer
   CodDocSII As Integer
   Folio As Long
   idEstado As Integer
   Fecha As Long
   FechaVenc As Long         '<FchVenc>
   FormaDePago As Integer  '<FmaPago> 1:Contado, 2:Crédito 3:Sin costo (entrega Gratuita)
   IdEntidad As Long
   Rut As String
   NotValidRut As Boolean
   RazonSocial As String
   Giro As String
   Direccion As String
   Comuna As String
   Ciudad As String
   Contacto As String
   MailReceptor As String     'mail de recepción de facturas
   SubTotal As Double
   PjeDestoGlobal As Single
   DesctoGlobal As Double
   Neto As Double
   Exento As Double
   TasaIVA As Single
   Iva As Double
   ImpAdicional As Double
   Total As Double
   DetDTE(MAX_ITEMDTE) As DetDTE_t
   ImpAdic(MAX_IMPADICDTE) As DTEImpAdic_t
   Referencia(MAX_REFDTE) As Referencia_t
   EsExport As Boolean
   FactExp As DTEFactExp_t
   TrackID As String
   TipoDespacho As Integer
   Traslado As Integer
   EsGuiaDesp As Boolean
   GuiaDesp As DTEGuiaDesp_t
   Observaciones As String
   Rebaja As Double
   DetFormaPago As Long
   TextDetFormaPago As String
   Vendedor As Long
   TextVendedor As String
End Type

'Tipo de docs DTE con la codificación interna y el código que maneja el SII
Type TipoDocDTE_t
   IdxTipoDoc     As Integer
   TipoLib        As Integer
   TipoDoc        As Integer
   Nombre         As String
   Diminutivo     As String
   CodDocDTESII   As String
End Type

Public gTipoDocDTE() As TipoDocDTE_t

'Antecedentes asociados a una Factura de Exportación

Type DatoFactExp_t
   codigo As String
   Nombre As String
End Type

'Indicador de Servicio
Public gIndServicio() As DatoFactExp_t

Public Const INDSERV_MERCADERIAS = "99"
Public Const INDSERV_TRANSPTERRESTRE = "5"



'Modalidad de Venta
Public gModVenta() As DatoFactExp_t

'Vía de Transporte
Public gViaTransporte() As DatoFactExp_t


'Estado de un DTE de acuerdo a la codificación interna
Public Const EDTE_ENVIADO = 1
Public Const EDTE_PROCESADO = 2
Public Const EDTE_EMITIDO = 3
Public Const EDTE_FOLIONODISP = 4
Public Const EDTE_ERROR = 5
Public Const EDTE_ANULADO = 6

Public Const MAX_ESTADODTE = EDTE_ANULADO
'Public Const MAX_ESTADODTE = EDTE_PAGADO
Public gEstadoDTE(EDTE_ANULADO) As String

'Estado de un DTE de acuerdo al SII
Public Const EDTESII_DESCONOCIDO = 0
Public Const EDTESII_PROCESADO = 1
Public Const EDTESII_ACEPTADO = 2
Public Const EDTESII_REPARO = 3        'con reparos
Public Const EDTESII_RECHAZADO = 4     'se puede reprocesar entregando el folio ¿?
Public Const EDTESII_PAGADO = 5
Public Const EDTESII_ENVIADO = 6
Public Const EDTESII_ANULADO = 7

'Public Const MAX_ESTADODTESII = EDTESII_RECHAZADO
Public Const MAX_ESTADODTESII = EDTESII_ANULADO

Public gEstadoDTESII(MAX_ESTADODTESII) As String     'estado interno
Public gDesEstadoDTESII(MAX_ESTADODTESII) As String

Public gTxtEstadoDTESII(MAX_ESTADODTESII) As String   'estado que se muestra en la traza del DTE, en el área SII
   
Type FmtImp_t
   Campo As String
   Formato As String
End Type

'Lista de últimas empresas abiertas en Menu Empresa
Type LastOpen_t
   Nombre As String
   Id     As Long
   
End Type



  

Public Sub IniLPFactura()
   
   Call ReadIni
   
   ReDim gPrivilegios(Log2(LAST_PRVF))

   gPrivilegios(Log2(PRVF_ADM_SIS)) = "Configurar Usuarios y Administrar Sistema"
   gPrivilegios(Log2(PRVF_CFG_EMP)) = "Crear y Configurar Empresa"
   gPrivilegios(Log2(PRVF_ADM_EMPRESA)) = "Administrar entidades, productos, sucursales"
   gPrivilegios(Log2(PRVF_ADM_EXP)) = "Administrar Exportación e Importación de Datos"
   gPrivilegios(Log2(PRVF_EMITIR_FACT)) = "Emitir DTE"
   gPrivilegios(Log2(PRVF_ADM_FACT)) = "Administrar/listar documentos emitidos"
   gPrivilegios(Log2(PRVF_MANT_DATOS)) = "Mantener datos básicos (Monedas, Países, etc.)"
 
   
   gProvFactElect(PROV_LP) = "Thomson Reuters"
   gProvFactElect(PROV_ACEPTA) = "Acepta"
   
   gEstadoMes(EM_NOEXISTE) = "No Existe"
   gEstadoMes(EM_ABIERTO) = "Abierto"
   gEstadoMes(EM_CERRADO) = "Cerrado"
   gEstadoMes(EM_ERRONEO) = "Erróneo"

   gAbrirMesesParalelo = False
   
   gEstadoEntidad(EE_ACTIVO) = "Activo"
   gEstadoEntidad(EE_INACTIVO) = "Inactivo"
   gEstadoEntidad(EE_BLOQUEADO) = "Bloqueado"
   
   gClasifEnt(ENT_CLIENTE) = "Cliente"
   gClasifEnt(ENT_PROVEEDOR) = "Proveedor"
   gClasifEnt(ENT_EMPLEADO) = "Empleado"
   gClasifEnt(ENT_SOCIO) = "Socio"
   gClasifEnt(ENT_DISTRIB) = "Distribuidor"
   gClasifEnt(ENT_OTRO) = "Otro"

   gTipoRefSII(REF_ANULA) = "Anula Documento de Referencia"
   gTipoRefSII(REF_CORRIGETEXTO) = "Corrige Texto Documento de Referencia"
   gTipoRefSII(REF_CORRIGEMONTOS) = "Corrige Montos"
   
   
   'Formas de Pago Factura
   gFormaDePago(FP_CONTADO) = "Contado"
   gFormaDePago(FP_CREDITO) = "Crédito"
   gFormaDePago(FP_SINCOSTO) = "Sin costo (Ent. Gratuita)"
   
   'Estados
   gEstado(ES_INACTIVO) = "Inactivo"
   gEstado(ES_ACTIVO) = "Activo"
   
   'Estado DTE Interno
   gEstadoDTE(EDTE_ENVIADO) = "Enviado"
   gEstadoDTE(EDTE_PROCESADO) = "Procesado"
   gEstadoDTE(EDTE_EMITIDO) = "Emitido"
   gEstadoDTE(EDTE_FOLIONODISP) = "Folio No Disponible"
   gEstadoDTE(EDTE_ERROR) = "Error"
   gEstadoDTE(EDTE_ANULADO) = "Anulado"
   'gEstadoDTE(EDTE_PAGADO) = "PAG"
   
   'Estado DTE de acuerdo al SII
'   gEstadoDTESII(EDTESII_DESCONOCIDO) = "Desconocido"
'   gEstadoDTESII(EDTESII_PROCESADO) = "Procesado"
'   gEstadoDTESII(EDTESII_ACEPTADO) = "Aceptado"
'   gEstadoDTESII(EDTESII_REPARO) = "Reparo"
'   gEstadoDTESII(EDTESII_RECHAZADO) = "Rechazado"
   gEstadoDTESII(EDTESII_DESCONOCIDO) = "Desconocido"
   gEstadoDTESII(EDTESII_PROCESADO) = "Procesado"
   gEstadoDTESII(EDTESII_ACEPTADO) = "AceptadoSii"
   gEstadoDTESII(EDTESII_REPARO) = "Reparo"
   gEstadoDTESII(EDTESII_RECHAZADO) = "RechazadoSii"
   gEstadoDTESII(EDTESII_PAGADO) = "PAG"
   gEstadoDTESII(EDTESII_ENVIADO) = "EnviadoSii"
   gEstadoDTESII(EDTESII_ANULADO) = "Anulado"
   
   gDesEstadoDTESII(EDTESII_DESCONOCIDO) = "Desconocido"
   gDesEstadoDTESII(EDTESII_PROCESADO) = "Procesado"
   gDesEstadoDTESII(EDTESII_ACEPTADO) = "Aceptado"
   gDesEstadoDTESII(EDTESII_REPARO) = "Reparo"
   gDesEstadoDTESII(EDTESII_RECHAZADO) = "Rechazado"
   gDesEstadoDTESII(EDTESII_PAGADO) = "Pagado"
   gDesEstadoDTESII(EDTESII_ENVIADO) = "Enviado"
   gDesEstadoDTESII(EDTESII_ANULADO) = "Anulado"
   
   'Estado DTE de acuerdo a la traza del DTE en área SII
   gTxtEstadoDTESII(EDTESII_DESCONOCIDO) = "Desconocido"   '¿?
   gTxtEstadoDTESII(EDTESII_PROCESADO) = "Procesado por SII"
   gTxtEstadoDTESII(EDTESII_ACEPTADO) = "Aceptado por SII"
   gTxtEstadoDTESII(EDTESII_REPARO) = "Reparo"
   gTxtEstadoDTESII(EDTESII_RECHAZADO) = "Rechazado por SII"
   gTxtEstadoDTESII(EDTESII_PAGADO) = "DTE Pagado al Contado"
   gTxtEstadoDTESII(EDTESII_ENVIADO) = "Enviado al SII TrackId="
   
'   gTxtEstadoDTESII(EDTESII_DESCONOCIDO) = "Desconocido"   '¿?
'   gTxtEstadoDTESII(EDTESII_PROCESADO) = "Procesado por SII"
'   gTxtEstadoDTESII(EDTESII_ACEPTADO) = "AceptadoSii"
'   gTxtEstadoDTESII(EDTESII_REPARO) = "Reparo"
'   gTxtEstadoDTESII(EDTESII_RECHAZADO) = "RechazadoSii"
'   gTxtEstadoDTESII(EDTESII_PAGADO) = "PAG"
'   gTxtEstadoDTESII(EDTESII_ENVIADO) = "EnviadoSii"
   
   
   'Tipo de Despacho
   gTipoDespacho(GD_SINDESPACHO) = "Sin Despacho"
   gTipoDespacho(GD_DESPRECEPTOR) = "Despacho por cta. receptor"
   gTipoDespacho(GD_DESPEMICLI) = "Despacho por cta. emisor a inst. cliente"
   gTipoDespacho(GD_DESPEMIOTRO) = "Despacho por cta. emisor a otras inst."


   'Traslado
   gTipoTraslado(GT_SINTRASLADO) = "Sin Traslado"
   gTipoTraslado(GT_VENTA) = "Operación constituye venta"
   gTipoTraslado(GT_VENTAPORFACT) = "Ventas por efectuar"
   gTipoTraslado(GT_CONSIGNACION) = "Consignaciones"
   gTipoTraslado(GT_ENTGRATIS) = "Entrega gratuita"
   gTipoTraslado(GT_TRASINTERNO) = "Traslados internos"
   gTipoTraslado(GT_OTRONOVENTA) = "Otros traslados no venta"
   gTipoTraslado(GT_DEVOLUCION) = "Guía de devolución"
   
   'Indicador de Servicio
   ReDim gIndServicio(4)
'   gIndServicio(1).Codigo = "1"
'   gIndServicio(1).Nombre = "Factura de servicios periódicos domiciliarios"
'   gIndServicio(2).Codigo = "2"
'   gIndServicio(2).Nombre = "Factura de otros servicios periódicos."
   gIndServicio(1).codigo = INDSERV_MERCADERIAS
   gIndServicio(1).Nombre = "Mercaderías"
   gIndServicio(2).codigo = "3"
   gIndServicio(2).Nombre = "Serv. calificados como tal por Aduana."
   gIndServicio(3).codigo = "4"
   gIndServicio(3).Nombre = "Serv. de Hotelería."
   gIndServicio(4).codigo = INDSERV_TRANSPTERRESTRE
   gIndServicio(4).Nombre = "Servicio de Transporte Terrestre Internacional"
   
   'Modalidad de Venta
   ReDim gModVenta(5)
   gModVenta(1).codigo = "1"
   gModVenta(1).Nombre = "A Firme."
   gModVenta(2).codigo = "2"
   gModVenta(2).Nombre = "Bajo Condición."
   gModVenta(3).codigo = "3"
   gModVenta(3).Nombre = "En consignación libre."
   gModVenta(4).codigo = "4"
   gModVenta(4).Nombre = "En consignación con un mínimo a firme."
   gModVenta(5).codigo = "9"
   gModVenta(5).Nombre = "Sin Pago."
   
   'Vía de Transporte
   ReDim gViaTransporte(7)
   gViaTransporte(1).codigo = "01"
   gViaTransporte(1).Nombre = "Marítima, Fluvial Y Lacustre"
   gViaTransporte(2).codigo = "04"
   gViaTransporte(2).Nombre = "Aéreo"
   gViaTransporte(3).codigo = "05"
   gViaTransporte(3).Nombre = "Postal"
   gViaTransporte(4).codigo = "06"
   gViaTransporte(4).Nombre = "Ferroviario"
   gViaTransporte(5).codigo = "07"
   gViaTransporte(5).Nombre = "Carretero / Terrestre"
   gViaTransporte(6).codigo = "08"
   gViaTransporte(6).Nombre = "Oleoductos, Gasoductos"
   gViaTransporte(7).codigo = "09"
   gViaTransporte(7).Nombre = "Tendido Eléctrico (Aéreo, Subterráneo)"
   gViaTransporte(8).codigo = "10"
   gViaTransporte(8).Nombre = "Otra"
  
   gOcultarImpAdicDescont = True

End Sub
Public Sub ReadIni()
   Dim Rc As Integer
   Dim Buf As String * 21
   
End Sub

Public Function IniEmpresa() As Boolean
   
   IniEmpresa = False
   
   'abrimos la base de datos de la empresa
   Call AddLog("IniEmpresa: a OpenDbEmp", 2)
   If OpenDbEmpFact() = False Then
    '  End
      Exit Function
   End If
   
   Call AddLog("IniEmpresa: a CorrigeBase", 2)
   Call CorrigeBase
   
   Call AddLog("IniEmpresa: a ChkDbInfo", 2)
   If ChkDbInfoFact(DbMain, gEmpresa.Rut, gEmpresa.Id) = False Then
      Call CloseDb(DbMain)
      Exit Function
   End If
   
      
   'linkeamos las tablas por si se movieron
   Call AddLog("IniEmpresa: a LinkMdbFact", 2)
   Call LinkMdbFact
   
      
   'inicializamos datos básicos del sistema
   Call AddLog("IniEmpresa: a ReadParam", 2)
   Call ReadParam
   
   'inicializamos datos básicos de la empresa
   Call AddLog("IniEmpresa: a ReadEmpresa", 2)
   Call ReadEmpresa
   
   'inicializamos path para los PDFs de los DTE
   gPdfDTEEmp = gPdfDTE & "\" & gEmpresa.Rut
   On Error Resume Next
   MkDir gPdfDTEEmp
   On Error GoTo 0
   
   IniEmpresa = True
   
   Call AddLog("IniEmpresa: nos vamos OK", 1)

End Function
Public Sub LinkMdbFact()
   Dim DbComun As String, DbEmp As String
   Dim ConnStr As String
   Dim Tm As Double

   Tm = CDbl(Now)
         
   'linkeamos las tablas de LPFactura por si se movieron
   DbComun = gDbPath & "\" & BD_COMUNDTE
   ConnStr = gComunConnStr
  
   Call LinkMdbTable(DbMain, DbComun, "Empresas", , , , ConnStr, True)
'   Call LinkMdbTable(DbMain, DbComun, "EmpresasAno", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "PcUsr", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "ParamDTE", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "TipoDocs", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "TipoValor", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "TipoDocRef", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Impuestos", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "IPC", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Usuarios", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Perfiles", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Param", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Regiones", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "CodActiv", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Monedas", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Equivalencia", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Paises", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Puertos", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "ClauCompraVenta", , , , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "TipoVehiculo", , , , ConnStr, True)
 
   Debug.Print "LinkMdbTableFact: Tiempo: " & Format((CDbl(Now) - Tm) / TimeSerial(0, 0, 1), NUMFMT) & " [s]"
   
End Sub

Private Sub ReadParam()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, j As Integer
   Dim NLib As Integer
   Dim CodTV As String
   
   Call ReadComun

   'clasificación de libros: código
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM Param WHERE Tipo='TIPOLIBCOD' ORDER BY Codigo")
   
   i = 1
   Do While Rs.EOF = False
   
      ReDim Preserve gTipoLibCod(i)
      gTipoLibCod(i) = vFld(Rs("Valor"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

   'clasificación de libros: nombre
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM Param WHERE Tipo='TIPOLIB' ORDER BY Codigo")
   
   i = 1
   Do While Rs.EOF = False
   
      ReDim Preserve gTipoLib(i)
      gTipoLib(i) = vFld(Rs("Valor"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

   Call ReadTipoDocs
   
   Call AddLog("ReadParam: pasamos lectura ReadTipoDocs", 1)
   
   NLib = UBound(gTipoLib)
      
   Call ReadTipoValor
   
'   Call ReadIndices
   
   Call AddLog("ReadParam: Nos vamos", 1)
      
End Sub

Public Sub ReadEmpresa()
   Dim Q1 As String
   Dim Rs As Recordset
     
   Call ReadParamEmp
   Call ReadDatosBasEmpresa
      
End Sub
Public Sub ReadParamEmp()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer

   'Configuraciones para New DTE
   For i = 1 To MAX_OPTEDFACT
      gEmpConfig.OptEdFact(i) = 1
      If i = OPTEDFACT_NOTSELPROD Then
         gEmpConfig.OptEdFact(i) = 0
      End If
         
      Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo='OPTEDFACT' AND Codigo=" & i & " AND IdEmpresa = " & gEmpresa.Id)
      If Rs.EOF = False Then
         gEmpConfig.OptEdFact(i) = Val((vFld(Rs("Valor"))))
      End If
      Call CloseRs(Rs)
   Next i
   
   'Parámetros de conexión
      'Proveedor
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo ='" & CONECT_PROV & "'" & " AND IdEmpresa = " & gEmpresa.Id)
   
   If Rs.EOF = False Then 'está
      gConectData.Proveedor = Val(vFld(Rs("Valor")))
   Else
      gConectData.Proveedor = 0
      Q1 = "INSERT INTO ParamEmpDTE(Tipo, Codigo, Valor, IdEmpresa) VALUES('" & CONECT_PROV & "', 0, '0', " & gEmpresa.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If

   Call CloseRs(Rs)
   
      'usuario
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_USUARIO & " AND IdEmpresa = " & gEmpresa.Id)
   
   If Rs.EOF = False Then 'está
      gConectData.Usuario = vFld(Rs("Valor"))
   Else
      gConectData.Usuario = ""
      Q1 = "INSERT INTO ParamEmpDTE(Tipo, Codigo, Valor, IdEmpresa) VALUES('" & CONECT_DATA & "'," & CONECT_USUARIO & ",' ', " & gEmpresa.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If

   Call CloseRs(Rs)
   
      'clave
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_CLAVE & " AND IdEmpresa = " & gEmpresa.Id)
   
   If Rs.EOF = False Then 'está
      gConectData.Clave = vFld(Rs("Valor"))
   Else
      gConectData.Clave = ""
      Q1 = "INSERT INTO ParamEmpDTE(Tipo, Codigo, Valor, IdEmpresa) VALUES('" & CONECT_DATA & "'," & CONECT_CLAVE & ",' ', " & gEmpresa.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If

   Call CloseRs(Rs)

      'clave Cert
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_CLAVECERT & " AND IdEmpresa = " & gEmpresa.Id)
   
   If Rs.EOF = False Then 'está
      gConectData.ClaveCert = vFld(Rs("Valor"))
   Else
      gConectData.ClaveCert = ""
      Q1 = "INSERT INTO ParamEmpDTE(Tipo, Codigo, Valor, IdEmpresa) VALUES('" & CONECT_DATA & "'," & CONECT_CLAVECERT & ",' ', " & gEmpresa.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If

   Call CloseRs(Rs)

      'mail emisor
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_MAILEMISOR & " AND IdEmpresa = " & gEmpresa.Id)
   
   If Rs.EOF = False Then 'está
      gConectData.MailEmisor = vFld(Rs("Valor"))
   Else
      gConectData.MailEmisor = ""
      Q1 = "INSERT INTO ParamEmpDTE(Tipo, Codigo, Valor, IdEmpresa) VALUES('" & CONECT_DATA & "'," & CONECT_MAILEMISOR & ",' ', " & gEmpresa.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If

   Call CloseRs(Rs)
   
   'Rut Firma
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_RUTFIRMA & " AND IdEmpresa = " & gEmpresa.Id)
   
   If Rs.EOF = False Then 'está
      gConectData.RutFirma = vFld(Rs("Valor"))
   Else
      gConectData.RutFirma = ""
      Q1 = "INSERT INTO ParamEmpDTE(Tipo, Codigo, Valor, IdEmpresa) VALUES('" & CONECT_DATA & "'," & CONECT_RUTFIRMA & ",' ', " & gEmpresa.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If

   Call CloseRs(Rs)

   If W.InDesign And gConectData.Usuario = "" Then
      gConectData.Usuario = "Fairware"
      gConectData.Clave = "123456"
      gConectData.Proveedor = PROV_ACEPTA ' PROV_LP
   End If

End Sub
Public Function OpenDbEmpFact(Optional ByVal Rut As String = "") As Integer
   Dim DbName As String
   Dim Passw As String, SqlErr As String
   Dim fType As String
   
   On Error Resume Next
   
   OpenDbEmpFact = True
   
   fType = "-DTE.mdb"
          
   If Rut <> "" Then
      DbName = gDbPath & "\Empresas\" & Rut & fType
   Else
      DbName = gDbPath & "\Empresas\" & gEmpresa.Rut & fType
   End If

   If Rut <> "" Then
      Passw = PASSW_PREFIX & Rut
   Else
      Passw = PASSW_PREFIX & gEmpresa.Rut
   End If
   
   Call AddLog("OpenDbEmpFact: DbName:[" & DbName & "]", 2)
   
   Call SetDbSecurity(DbName, Passw, gCfgFile, SG_SEGCFG, gEmpresa.ConnStr)

   If Not (DbMain Is Nothing) Then
      Call CloseDb(DbMain)
   End If
   
   Err.Clear
   'Set DbMain = OpenDatabase(DbName, True, False, ConnStr) ' MODO EXCLUSIVO
   Set DbMain = OpenDatabase(DbName, False, False, gEmpresa.ConnStr)
'   gEmpresa.ConnStr = Mid(gEmpresa.ConnStr, 2) 'sin el ; del principio
   
   If Err Then
      SqlErr = "Error " & Err & ", '" & Error & "'"
   
      If Err = 3356 Then
         MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
         OpenDbEmpFact = False
      End If
   
   End If
   
   If (Err Or DbMain Is Nothing) And Err <> 3356 Then
      MsgBox SqlErr & vbCrLf & DbName, vbExclamation
      OpenDbEmpFact = False
   End If
   
   Call ChkDbSize(DbMain, 200 * 1024) ' 200 MB
   
   Call AddLog("OpenDbEmpFact: fin OK", 2)

End Function
Public Sub ReadDatosBasEmpresa()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MesActual As Integer, i As Integer
   Dim InitAno As String

      
   Call AddLog("ReadDatosBasEmpresa: llegamos", 1)
        
   gEmpresa.Direccion = ""
   gEmpresa.Telefono = ""
   gEmpresa.RazonSocial = ""
   gEmpresa.Comuna = ""
   gEmpresa.Ciudad = ""
   gEmpresa.Giro = ""
   gEmpresa.CodActEcono = ""
   gEmpresa.RepConjunta = False
   gEmpresa.RutRepLegal1 = ""
   gEmpresa.RepLegal1 = ""
   gEmpresa.RutRepLegal2 = ""
   gEmpresa.RepLegal2 = ""
   gEmpresa.Opciones = 0
   gEmpresa.Franq14Ter = 0
   gEmpresa.ObligaLibComprasVentas = 0
   gEmpresa.email = ""
   gEmpresa.ObsDTE = ""

   'leemos los datos de la empresa en el único registro de esta tabla
   Q1 = "SELECT NombreCorto, Calle, Numero, Dpto, Telefonos, RazonSocial, ApMaterno,Nombre, EMail, ObsDTE "
   Q1 = Q1 & ", Giro, CodActEconom, RepConjunta, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2"
   Q1 = Q1 & ", Regiones.Comuna, Ciudad, Opciones "
   Q1 = Q1 & "  FROM Empresa LEFT JOIN Regiones ON Empresa.Comuna=Regiones.id"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then
      'ES LA PRIMERA Y SE HACE INSERT EN ESTA TABLA
      Q1 = "INSERT INTO Empresa (id, Rut, NombreCorto, RazonSocial) VALUES ("
      Q1 = Q1 & gEmpresa.Id
      Q1 = Q1 & ",'" & gEmpresa.Rut & "'"
      Q1 = Q1 & ",'" & gEmpresa.NombreCorto & "'"
      Q1 = Q1 & ",'" & gEmpresa.NombreCorto & "')"
      Call ExecSQL(DbMain, Q1)
      
   Else
      gEmpresa.Direccion = vFld(Rs("Calle"), True) & " " & vFld(Rs("Numero"), True) & " " & vFld(Rs("Dpto"), True)
      gEmpresa.Telefono = vFld(Rs("Telefonos"), True)
      gEmpresa.RazonSocial = vFld(Rs("RazonSocial"), True) & " " & vFld(Rs("ApMaterno"), True) & " " & vFld(Rs("Nombre"), True)
      gEmpresa.Comuna = vFld(Rs("Comuna"), True)
      gEmpresa.Ciudad = vFld(Rs("Ciudad"), True)
      gEmpresa.email = vFld(Rs("EMail"), True)
      gEmpresa.ObsDTE = vFld(Rs("ObsDTE"), True)
      gEmpresa.Giro = vFld(Rs("Giro"), True)
      gEmpresa.CodActEcono = vFld(Rs("CodActEconom"), True)
      gEmpresa.RepConjunta = vFld(Rs("RepConjunta"))
      gEmpresa.RutRepLegal1 = vFld(Rs("RutRepLegal1"))
      gEmpresa.RepLegal1 = vFld(Rs("RepLegal1"), True)
      gEmpresa.RutRepLegal2 = vFld(Rs("RutRepLegal2"))
      gEmpresa.RepLegal2 = vFld(Rs("RepLegal2"), True)
      gEmpresa.Opciones = vFld(Rs("Opciones"))
   End If
   
   Call CloseRs(Rs)
   
   Call AddLog("ReadDatosBasEmpresa: pasamos datos empresa", 1)
      
   'IVA
   gIVA = 0.19
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo='VALORIVA'" & " AND IdEmpresa = " & gEmpresa.Id)
   If Rs.EOF = False Then
      gIVA = Val((vFld(Rs("Valor"))))
   End If
   Call CloseRs(Rs)
 
   'estado meses
   
'   Set Rs = OpenRs(DbMain, "SELECT Mes, Estado FROM EstadoMes ORDER BY Mes desc")
'
'   MesActual = 0
'   i = 0
'
'   'calculamos el mes actual
'   Do While Rs.EOF = False
'      i = i + 1
'      If vFld(Rs("Estado")) = EM_ABIERTO Then
'         MesActual = vFld(Rs("Mes"))
'         Exit Do
'      End If
'      Rs.MoveNext
'   Loop
'
'   Call CloseRs(Rs)
'
'   If i = 0 Then    'la tabla está vacía
'      Call LlenarTablaMeses("")
'      MesActual = 1          'parte con enero
'      AbrirMes (MesActual)
'
'   'ElseIf MesActual = 0 Then   'no hay ningún mes abierto => se terminó el año
'
'   End If

   'Configuraciones para New DTE
   For i = 1 To MAX_OPTEDFACT
      gEmpConfig.OptEdFact(i) = 1
      If i = OPTEDFACT_NOTSELPROD Then
         gEmpConfig.OptEdFact(i) = 0
      End If
         
      Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo='OPTEDFACT' AND Codigo=" & i & " AND IdEmpresa = " & gEmpresa.Id)
      If Rs.EOF = False Then
         gEmpConfig.OptEdFact(i) = Val((vFld(Rs("Valor"))))
      End If
      Call CloseRs(Rs)
   Next i
   
   

End Sub

Public Function ReadTipoDocs()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, j As Integer
   Dim IdxFAV As Integer
 
   'tipos de docs, independiente de los libros   (no se usa la de LPContabilidad porque en ella se asignan los tipos de LAU (Libros Auxiliares HR)
   
   ReDim gTipoDoc(10)
   ReDim gTipoDocDTE(10)
   
   Q1 = "SELECT Id, TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TieneAfecto,"
   Q1 = Q1 & "TieneExento, ExigeRUT, EsRebaja, DocImpExp, CodDocSII, CodDocDTESII, AceptaPropIVA "
   Q1 = Q1 & " FROM TipoDocs WHERE Atributo='ACTIVO' ORDER BY TipoLib, TipoDoc"
   Set Rs = OpenRs(DbMain, Q1)
      
   i = 0
   j = 1
   
   Do While Rs.EOF = False
   
      If i > UBound(gTipoDoc) Then
         ReDim Preserve gTipoDoc(i + 10)
      End If
      
      gTipoDoc(i).Id = vFld(Rs("Id"))
      gTipoDoc(i).TipoLib = vFld(Rs("TipoLib"))
      gTipoDoc(i).TipoDoc = vFld(Rs("TipoDoc"))
      gTipoDoc(i).Nombre = vFld(Rs("Nombre"))
      gTipoDoc(i).Diminutivo = vFld(Rs("Diminutivo"))
      gTipoDoc(i).Atributo = vFld(Rs("Atributo"))
      gTipoDoc(i).TieneAfecto = vFld(Rs("TieneAfecto"))
      gTipoDoc(i).TieneExento = vFld(Rs("TieneExento"))
      gTipoDoc(i).ExigeRUT = vFld(Rs("ExigeRUT"))
      gTipoDoc(i).EsRebaja = vFld(Rs("EsRebaja"))
      gTipoDoc(i).TipoDocLAU = -1
      gTipoDoc(i).DocImpExp = vFld(Rs("DocImpExp"))
      gTipoDoc(i).CodDocSII = vFld(Rs("CodDocSII"))
      gTipoDoc(i).CodDocDTESII = Val(vFld(Rs("CodDocDTESII")))  'se elimina el cero a la izquierda
      gTipoDoc(i).AceptaPropIVA = vFld(Rs("AceptaPropIVA"))
            
      If gTipoDoc(i).TipoLib = LIB_VENTAS And Val(gTipoDoc(i).CodDocDTESII) <> 0 And InStr(LCase(gTipoDoc(i).Nombre), "boleta") = 0 Then 'no se consideran las boletas
         If j > UBound(gTipoDocDTE) Then
            ReDim Preserve gTipoDocDTE(j + 10)
         End If
         gTipoDocDTE(j).IdxTipoDoc = i
         gTipoDocDTE(j).TipoLib = gTipoDoc(i).TipoLib
         gTipoDocDTE(j).TipoDoc = gTipoDoc(i).TipoDoc
         gTipoDocDTE(j).Nombre = gTipoDoc(i).Nombre
         gTipoDocDTE(j).Diminutivo = gTipoDoc(i).Diminutivo
         gTipoDocDTE(j).CodDocDTESII = gTipoDoc(i).CodDocDTESII
         
         If gTipoDocDTE(j).Diminutivo = "FAV" Then
            IdxFAV = i
         End If
         
         j = j + 1
      End If
            
      i = i + 1
      
      Rs.MoveNext
   Loop

   'Agregamos las guías de despacho en el item 0
      
   gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).IdxTipoDoc = IdxFAV     'asimiliamos una guía de despacho a una factura de ventas
   gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).TipoLib = LIB_OTROS
   gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).TipoDoc = TIPODOC_GUIADESPACHO
   gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).Nombre = "Guía de Despacho"
   gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).Diminutivo = "GDE"
   gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).CodDocDTESII = CODDOCDTESII_GUIADESPACHO   'aquí se marca que es guía de despacho

   If i > 0 Then
      ReDim Preserve gTipoDoc(i - 1)
   End If
   If j > 0 Then
      ReDim Preserve gTipoDocDTE(j - 1)
   End If
   
   Call CloseRs(Rs)

End Function
Public Function GetTipoDoc(ByVal TipoLib As Integer, ByVal TipoDoc As Integer) As Integer
   Dim i As Integer
   
   GetTipoDoc = -1
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).TipoDoc = TipoDoc Then
         GetTipoDoc = i
         Exit Function
      End If
   
   Next i
   
End Function
Public Function GetTipoDocFromCodDocDTESII(ByVal TipoLib As Integer, ByVal CodDocDTESII As String) As Integer
   Dim i As Integer
   
   GetTipoDocFromCodDocDTESII = 0
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).CodDocDTESII = CodDocDTESII Then
         GetTipoDocFromCodDocDTESII = gTipoDoc(i).TipoDoc
         Exit Function
      End If
   
   Next i
   
End Function

'Crea nueva empresa, si no existe,  con DB vacía.
'Supone que está abierta la DB LpFactura
Public Function CrearNuevaEmprFact(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal NombreEmpresa As String) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RutMdb As String
   Dim Msg As String
   Dim EmpVacia As Boolean
   Dim DbActual As Database
   Dim PathDbActual As String
   Dim CopyErr As Boolean
   Dim FCierre As Long
   Dim DbPath As String
   Dim Frm As Form
   Dim Rc As Integer
   Dim ConnStr As String
   Dim IdCompAperTrib As Long
   Dim NuevoAnoVacio As Boolean
      Dim Tbl As TableDef
   Dim fld As Field

      
   'Chequeo si está creada la base de datos para nueva empresa, si no, la creo
   RutMdb = Rut & "-DTE.mdb"
   
   EmpVacia = False
   
   Call AddLog("CrearNuevaEmprFact 1", 1)

   If ExistFile(gDbPath & "\Empresas\" & RutMdb) = True Then     'ya existe la empresa
      CrearNuevaEmprFact = True
      Call AddLog("CrearNuevaEmprFact 2", 1)
      
      Exit Function
   End If
   
   Call AddLog("CrearNuevaEmprFact 3", 1)
   
   If MsgBox1("No existe información para esta empresa." & vbCrLf & vbCrLf & "¿Desea crearla?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      Exit Function
   End If
   
   CrearNuevaEmprFact = False
   
  'No existe, lo creamos
        
   On Error Resume Next
  
   Err.Clear
   
   If CrearMdbVaciaFact(RutMdb) = False Then
      Exit Function
   End If
   
   EmpVacia = True
            
   Call AddLog("CrearNuevaEmprFact 4", 1)
            
   'Guardo la base de datos actual y abro la DB nueva, para borrar los registros de las tablas que corresponden
   Set DbActual = DbMain
   Set DbMain = Nothing
   
   If OpenDbEmpFact(Rut) = False Then
      Exit Function
   End If
     
   Q1 = "DELETE * FROM ParamEmpDTE WHERE Tipo = " & TPE_DBINFO ' & " AND IdEmpresa = " & gEmpresa.id
   Call ExecSQL(DbMain, Q1)

   Call ChkDbInfoFact(DbMain, Rut, IdEmpresa)
         
   'cierro la nueva DB
   Call CloseDb(DbMain)
   
   Set DbMain = DbActual
      
   CrearNuevaEmprFact = True
   
   Call AddLog("CrearNuevaEmprFact FIN ", 1)
   
End Function

' Verifica que no se haya copiado un archivo a otro RUT
Public Function ChkDbInfoFact(Db As Database, ByVal Rut As String, ByVal IdEmpresa As Long) As Boolean
   Dim Q1 As String, Rs As Recordset
   
   ChkDbInfoFact = True
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpDTE WHERE Tipo = " & TPE_DBINFO ' & " AND IdEmpresa = " & gEmpresa.id & " ORDER BY Codigo"
   Set Rs = OpenRs(Db, Q1)
   
   If Rs.EOF Then ' es nueva o no se había agregado
      Call CloseRs(Rs)

      ' Guardo datos que identifican esta base
      Q1 = "INSERT INTO ParamEmpDTE (Tipo, Codigo, Valor)"
      Q1 = Q1 & " VALUES(" & TPE_DBINFO & ", 1, '" & Rut & "')"
      Call ExecSQL(Db, Q1)
         
      Q1 = "INSERT INTO ParamEmpDTE (Tipo, Codigo, Valor)"
      Q1 = Q1 & " VALUES(" & TPE_DBINFO & ", 3, '" & IdEmpresa & "')"
      Call ExecSQL(Db, Q1)
      
   Else
      
      Do Until Rs.EOF
         
         Select Case vFld(Rs("Codigo"))
         
            Case 1: ' RUT
               If Rut <> vFld(Rs("Valor")) Then
                  MsgBox1 "Esta base de datos es del RUT " & vFld(Rs("Valor")) & " y no corresponde al RUT de la Empresa seleccionada.", vbCritical
                  ChkDbInfoFact = False
               End If
            
            Case 3: ' idEmpresa
               If IdEmpresa <> Val(vFld(Rs("Valor"))) Then   'permitimos esto para el caso en que hay que reconstruir la LPFActura
               
                  If ChkPriv(PRV_ADM_EMPRESA) Then
               
                     If MsgBox1("ATENCIÓN" & vbCrLf & "Esta base de datos no corresponde a la Empresa seleccionada." & vbCrLf & vbCrLf & "¿ Desea continuar bajo su responsabilidad ?", vbCritical Or vbYesNo) = vbYes Then
                        Call AddLog("Cambia idEmpresa=" & IdEmpresa & " para " & DbMain.Name)
                        Q1 = "UPDATE ParamEmpDTE SET Valor = " & IdEmpresa & " WHERE Tipo = " & TPE_DBINFO & " AND Codigo = 3 AND IdEmpresa = " & gEmpresa.Id
                        Call ExecSQL(DbMain, Q1)
                        Q1 = "UPDATE Empresa SET id = " & IdEmpresa    'es un solo registro en esta tabla
                        Call ExecSQL(DbMain, Q1)
                     Else
                        ChkDbInfoFact = False
                     End If
                  Else
                     MsgBox1 "Esta base de datos no corresponde a la Empresa seleccionada.", vbCritical
                     ChkDbInfoFact = False
                  End If
               End If
            
         End Select
      
         Rs.MoveNext
      Loop
      
      Call CloseRs(Rs)
      
   End If
      
End Function

Public Function CrearMdbVaciaFact(ByVal RutMdb As String) As Boolean
   Dim FName As String
   
   CrearMdbVaciaFact = False
   
   On Error Resume Next
   
   If ExistFile(gDbPath & "\Empresas\" & RutMdb) = False Then
      'If MsgBox1("¡ADVERTENCIA!, no existe información de la empresa para este año. ¿Desea crearla?", vbYesNo Or vbDefaultButton1 Or vbQuestion) <> vbYes Then
      '   Exit Function
      'End If
      
      If Not ExistFile(gDbPath & "\" & BD_VACIADTE) Then
         MsgBox1 "No se encontró el archivo """ & gDbPath & "\" & BD_VACIADTE & """." & vbCrLf & "Por favor, contacte a personal de soporte del sistema.", vbExclamation + vbOKOnly
         Exit Function
      End If
            
      Call FileCopy(gDbPath & "\" & BD_VACIADTE, gDbPath & "\Empresas\" & RutMdb)
   
      If Err = 75 Then
         MsgBox1 "¡ADVERTENCIA!, no se podrá crear la empresa porque no se ha encontrado " & BD_VACIADTE & " en el directorio " & gDbPath & "\Datos. Verifique si el archivo existe. Si no es así, búsquelo en el CD de instalación del sistema o bien comuníquese con soporte.", vbExclamation
         Exit Function
        
      End If
      
      If Err = 76 Then
         MsgBox1 "¡ADVERTENCIA!, no se podrá crear la empresa, porque no existe el directorio ..\Empresas bajo el directorio ..\Datos.", vbExclamation
         Exit Function
      
      End If
      
   End If
         
   CrearMdbVaciaFact = True
   
End Function


Public Sub SetDbPathFact(Drv As DriveListBox)
   Dim DbPath As String, Rc As Long
   Dim Q1 As String, Rs As Recordset
   Dim i As Integer, j As Integer, k As Integer
      
   DbPath = GetAbsPath(gDbPath, Drv)
   If DbPath <> gDbPath Then
      Call AddLog("SetDbPathFact: Se cambia [" & gDbPath & "] por [" & DbPath & "]")
      gDbPath = DbPath
   End If
   
'OJO ver con Pablo
   
   ' 16 mar 2012: para poder forzar a que no lea la tabla LParam (Lnk = 0)
   i = Val(GetIniString(gCfgFile, "Config", "Local", "0"))
   
'   If Left(gDbPath, 2) <> "\\"  Then
   If Left(gDbPath, 2) <> "\\" And i = 0 Then

      Q1 = "SELECT Valor FROM LParam WHERE Codigo=1"
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         DbPath = vFld(Rs("Valor"), True)
      End If
      Call CloseRs(Rs)

      If Left(DbPath, 2) = "\\" Then
         If ExistFile(DbPath & "\" & BD_COMUNDTE) Then      'ver con Pablo, demora mucho!
            gDbPath = DbPath
         Else
            Q1 = "UPDATE LParam SET Valor='" & ParaSQL(gDbPath) & "' WHERE Codigo=1"
            Rc = ExecSQL(DbMain, Q1)
         End If
      End If

   Else
      Q1 = "UPDATE LParam SET Valor='" & ParaSQL(gDbPath) & "' WHERE Codigo=1"
      Rc = ExecSQL(DbMain, Q1)

   End If

   Call AddLog("SetDbPathFact: gDbPath= [" & gDbPath & "]", 2)

End Sub

Public Function OpenDbAdmFact()
   OpenDbAdmFact = OpenDbAdm(BD_COMUNDTE)
End Function

Public Sub LlenarTablaMeses(ByVal Tbl As String)
   Dim Q1 As String
   Dim i As Integer
   
   If Tbl = "" Then
      Tbl = "EstadoMes"
   End If
   
   For i = 1 To 12
      Q1 = "INSERT INTO " & Tbl
      Q1 = Q1 & " (Mes, Estado, FechaApertura, FechaCierre) "
      Q1 = Q1 & " VALUES(" & i & "," & EM_CERRADO & ", 0, 0 )"
      Call ExecSQL(DbMain, Q1)
   Next i

End Sub

Public Function AbrirMes(ByVal Mes As Integer) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim EstadoMes As Integer
   Dim Impreso As Boolean
   
   AbrirMes = False
   
   EstadoMes = GetEstadoMes(Mes)
   
   Select Case EstadoMes
      Case EM_ABIERTO
         MsgBox1 "Este mes ya está abierto.", vbExclamation + vbOKOnly
         Exit Function
      Case EM_CERRADO
      Case EM_NOEXISTE
         MsgBox1 "Este mes no existe en la base de datos.", vbExclamation + vbOKOnly
         Exit Function
      Case EM_ERRONEO
         MsgBox1 "Este mes no está cuadrado. Se abrirá para que lo cuadre.", vbExclamation + vbOKOnly
   End Select
      
   Q1 = "UPDATE EstadoMes SET Estado = " & EM_ABIERTO & ", FechaApertura = " & CLng(Int(Now)) & " WHERE Mes = " & Mes
   Call ExecSQL(DbMain, Q1)
      
   AbrirMes = True
   
End Function
Public Function GetEstadoMes(ByVal Mes As Integer) As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Set Rs = OpenRs(DbMain, "SELECT Estado FROM EstadoMes WHERE Mes = " & Mes)
   If Rs.EOF = False Then
      GetEstadoMes = vFld(Rs("Estado"))
   Else
      GetEstadoMes = EM_NOEXISTE
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Sub CreatePrtFormats()
   Dim Nombres(7) As String
   Dim FntNombres(7) As FontDef_t
   Dim FntTitulos(0) As FontDef_t
   Dim FntEncabezados(0) As FontDef_t
   Dim i As Integer
               
   Set gPrtReportes = New ClsPrtFlxGrid
      
   gPrtReportes.PrtDemo = gAppCode.Demo
      
   FntNombres(0).FontName = "Arial"
   FntNombres(0).FontSize = 10
   FntNombres(0).FontBold = True
   
   For i = 1 To UBound(FntNombres)
      FntNombres(i).FontName = "Arial"
      FntNombres(i).FontSize = 9
      FntNombres(i).FontBold = False
   Next i
   
   Call gPrtReportes.FntNombres(FntNombres)
   
   FntTitulos(0).FontName = "Arial"
   FntTitulos(0).FontSize = 14
      
   Call gPrtReportes.FntTitulos(FntTitulos)
      
   FntEncabezados(0).FontName = "Arial"
   FntEncabezados(0).FontSize = 10
   
   Call gPrtReportes.FntEncabezados(FntEncabezados)
   
End Sub

Public Sub SetPrtData()
   Dim Nombres(7) As String
   Dim i As Integer

   For i = 0 To UBound(Nombres)
      Nombres(i) = ""
   Next i
   
   For i = 0 To UBound(Nombres)
      Nombres(i) = ""
   Next i
   
   Nombres(0) = gEmpresa.RazonSocial
   
   If gEmpresa.RutDisp = "" Then  'lo típico
      Nombres(1) = "RUT:" & vbTab & FmtCID(gEmpresa.Rut)
   Else
      Nombres(1) = "RUT:" & vbTab & FmtCID(gEmpresa.RutDisp)   'sólo para la Asoc. de AFP que tiene varias empresas con el mismo RUT
   End If
   
   If gEmpresa.Direccion <> "" And gEmpresa.Comuna <> "" Then
      Nombres(2) = "Dirección:" & vbTab & gEmpresa.Direccion & ", " & gEmpresa.Comuna
   ElseIf gEmpresa.Direccion <> "" Then
      Nombres(2) = "Dirección:" & vbTab & gEmpresa.Direccion
   Else
      Nombres(2) = "Dirección:" & vbTab & gEmpresa.Comuna
   End If
   
   Nombres(3) = "Teléfono:" & vbTab & gEmpresa.Telefono
   If gEmpresa.Fax <> "" Then
      Nombres(4) = "Fax:" & vbTab & gEmpresa.Fax
   End If
   
   gPrtReportes.Nombres = Nombres
   gPrtReportes.TabNombres = GetPrtTextWidth("Dirección:ww")
   
End Sub


Public Function CerrarMes(ByVal Mes As Integer) As Boolean
   Dim Q1 As String
   
   CerrarMes = False
   
   If Not ValidaCierreMes(Mes) Then
      Exit Function
   End If
   
   Q1 = "UPDATE EstadoMes SET Estado = " & EM_CERRADO & ", FechaCierre = " & CLng(Int(Now)) & " WHERE Mes = " & Mes
   Call ExecSQL(DbMain, Q1)
      
   CerrarMes = True

End Function

Public Function ValidaCierreMes(ByVal Mes As Integer) As Boolean

   ValidaCierreMes = True

End Function

Public Function GetUltimoMesConMovs(Optional ByVal Msg As Boolean = False) As Integer
   Dim Rs As Recordset
   Dim MaxFechaComp As Long
   Dim MaxFechaDoc As Long
   Dim MaxFecha As Long
   
   MaxFecha = 0
   MaxFechaComp = 0
   MaxFechaDoc = 0
   
   Set Rs = OpenRs(DbMain, "SELECT Max(Fecha) FROM DTE")
   If Rs.EOF = False Then
      MaxFecha = vFld(Rs(0))
   End If
      
   Call CloseRs(Rs)
         
   If MaxFecha > 0 Then
      GetUltimoMesConMovs = Month(MaxFecha)
   Else
      GetUltimoMesConMovs = 1  'partimos con enero
   End If
      
End Function


Public Sub ResetPrtBas(PrtCls As ClsPrtFlxGrid)
   Dim ColWi(0) As Integer
   Dim Total(0) As String
   Dim Titulos(0) As String
   Dim FntTitulos(0) As FontDef_t
   Dim FntEncabezados(0) As FontDef_t
   Dim Encabezados(0) As String
   Dim EncabezadosCont(0) As String

   PrtCls.CallEndDoc = True
   PrtCls.ColObligatoria = 1
   PrtCls.PrintHeader = True
   PrtCls.EsContinuacion = False
   
   PrtCls.GrFontName = ""
   PrtCls.GrFontSize = -1
   PrtCls.TotFntBold = True
   
   PrtCls.InitPag = -1
   PrtCls.CellHeight = 0
   
   PrtCls.FmtCol = -1
   
   PrtCls.ColWi = ColWi
   PrtCls.Titulos = Titulos
   PrtCls.Encabezados = Encabezados
   PrtCls.EncabezadosCont = EncabezadosCont
   Call PrtCls.FntTitulos(FntTitulos)
   Call PrtCls.FntEncabezados(FntEncabezados)

End Sub

Public Function FindTipoDoc(ByVal TipoLib As Integer, ByVal Diminutivo As String) As Integer
   Dim i As Integer
   
   FindTipoDoc = 0
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).Diminutivo = Diminutivo Then
         FindTipoDoc = gTipoDoc(i).TipoDoc
         Exit Function
      End If
      
   Next i
   
End Function

Public Sub FillTipoValLib(CbTipoValLib As ClsCombo, ByVal TipoLib As Integer, ByVal CallClear As Boolean, ByVal AddBlankItem As Boolean, Optional ByVal Atributo As String = "", Optional ByVal TipoDoc As Integer = 0, Optional ByVal IniTipoValLib As Integer = 0, Optional ByVal OcultarImpAdicDescontinuados As Boolean = False)
   Dim i As Integer
   Dim InitTipoLib As Boolean
   
   If CallClear Then
      CbTipoValLib.Clear
   End If
   
   If AddBlankItem Then
   
      Call CbTipoValLib.AddItem(" ", 0, "", "")
      
   End If
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib Then
      
         InitTipoLib = True
         
         If TipoDoc = 0 Or (TipoDoc <> 0 And (gTipoValLib(i).TipoDoc = "" Or InStr(gTipoValLib(i).TipoDoc, "," & TipoDoc & ",") <> 0)) Then
            If Atributo = "" Or (Atributo <> "" And gTipoValLib(i).Atributo = Atributo) Then
               If IniTipoValLib = 0 Or (IniTipoValLib > 0 And gTipoValLib(i).TipoValLib >= IniTipoValLib) Then
                  If Not OcultarImpAdicDescontinuados Or (OcultarImpAdicDescontinuados And Not gTipoValLib(i).Descontinuado) Then
                     Call CbTipoValLib.AddItem(gTipoValLib(i).Nombre, gTipoValLib(i).TipoValLib, gTipoValLib(i).CodSIIDTE, gTipoValLib(i).Tasa)
                  End If
               End If
            End If
         End If
      ElseIf InitTipoLib = True Then   'terminó el libro solicitado (están ordenados por TipoLib, TipoValLib (Codigo))
         Exit For
            
      End If
   
   Next i
   
   If CbTipoValLib.ListCount > 0 Then
      CbTipoValLib.ListIndex = 0
   End If
     
End Sub
Public Sub FillClsTipoValLib(Cb As ClsCombo, ByVal TipoLib As Integer, ByVal CallClear As Boolean, ByVal AddBlankItem As Boolean, Optional ByVal Atributo As String = "", Optional ByVal TipoDoc As Integer = 0, Optional ByVal OcultarImpAdicDescontinuados As Boolean = False)
   Dim i As Integer
   Dim InitTipoLib As Boolean
   Dim Tasa As Single, EsRecuperable As Boolean
   Dim IdCuenta As Long
   
   If CallClear Then
      Call Cb.Clear
   End If
   
   If AddBlankItem Then
   
      Call Cb.AddItem(" ", 0, 0, 0)
      
   End If
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib Then
      
         InitTipoLib = True
         
         If TipoDoc = 0 Or (TipoDoc <> 0 And (gTipoValLib(i).TipoDoc = "" Or InStr(gTipoValLib(i).TipoDoc, "," & TipoDoc & ",") <> 0)) Then
            If Atributo = "" Or (Atributo <> "" And gTipoValLib(i).Atributo = Atributo) Then
               If Not OcultarImpAdicDescontinuados Or (OcultarImpAdicDescontinuados And Not gTipoValLib(i).Descontinuado) Then
                  Call Cb.AddItem(gTipoValLib(i).Nombre, gTipoValLib(i).TipoValLib, gTipoValLib(i).CodSIIDTE, gTipoValLib(i).Tasa)
               End If
            End If
         End If
      ElseIf InitTipoLib = True Then   'terminó el libro solicitado (están ordenados por TipoLib, TipoValLib (Codigo))
         Exit For
            
      End If
   
   Next i
   
   If Cb.ListCount > 0 Then
      Cb.ListIndex = 0
   End If
     
End Sub
Public Function GetMaxTableId(ByVal IdName As String, ByVal TableName As String, ByVal Where As String) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Max(" & IdName & ") FROM " & TableName & " " & Where
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = True Then
      GetMaxTableId = 1
   ElseIf IsNull(Rs(0)) Then
      GetMaxTableId = 1
   Else
      GetMaxTableId = Rs(0) + 1
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function TienePrivilegio(Priv As Long, PrivSet As Long) As Boolean
   TienePrivilegio = ((Priv And PrivSet) <> 0)
End Function


Public Sub AbrirPDF(ByVal FName As String)
   Dim Pdf As ExtInfo_t, Cmd As String

   If GetExtInfo(".pdf", Pdf) = False Then
      MsgBox1 "No hay ninguna aplicación asociada a la extensión PDF.", vbExclamation
   Else
      If Pdf.OpenCmd <> "" Then
         Cmd = GenCmd(Pdf, "open", FName)
         Call ExecCmd(Cmd, vbMaximizedFocus, 10000)
      End If
   End If

End Sub
'Public Function SaveEstadoDTE(ByVal IdDTE As Long) As Integer
'   Dim Resultado As Integer, Rc As Integer
'   Dim Error As String, Respuesta As String, Glosa As String
'   Dim Q1 As String
'   Dim DTE As DTE_t
'
'   SaveEstadoDTE = 0
'
'   If IdDTE <= 0 Then
'      Exit Function
'   End If
'
'   Call FillDTEStruct(IdDTE, DTE)
'
'
'   Rc = LPConsultaEstadoDTE(DTE, Resultado, Respuesta, Glosa, Error)
'
'   If Rc = 0 Then
'      If Resultado = True Then
'
'         Q1 = "UPDATE DTE SET "
'         Q1 = Q1 & "  RespuestaSII = '" & ParaSQL(Respuesta) & "'"
'         Q1 = Q1 & ", GlosaSII = '" & ParaSQL(Glosa) & "'"
'         Q1 = Q1 & ", ErrorSII = ' '"
'         Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
'
'      Else
'         Q1 = "UPDATE DTE SET "
'         Q1 = Q1 & " ErrorSII = ' '"
'         Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
'
'      End If
'
'      Call ExecSQL(DbMain, Q1)
'   End If
'
'   SaveEstadoDTE = Rc
'End Function
'
'Public Function SaveEstadoEnviadoDTE(ByVal IdDTE As Long, ByVal TrackID As String) As Long
'   Dim Resultado As Integer
'   Dim Estado As String
'   Dim Error As String, Glosa As String
'   Dim Q1 As String, Rc As Long
'
'   SaveEstadoEnviadoDTE = 0
'
'   If TrackID = "" Then
'      Exit Function
'   End If
'
'
'   Rc = LPConsultaEstadoEnviadoDTE(TrackID, Resultado, Estado, Glosa, Error)
'
'   If Rc = 0 Then
'      If Resultado = True Then
'         Q1 = "UPDATE DTE SET "
'         Q1 = Q1 & "  IdEstadoSII = " & GetCodigoEstadoSII(Estado)
'         Q1 = Q1 & ", GlosaSII = '" & ParaSQL(Glosa) & "'"
'         Q1 = Q1 & ", ErrorSII = ' '"
'         Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
'      Else   'no se pudo obtener el estado del DTE
'         Q1 = "UPDATE DTE SET "
'         Q1 = Q1 & "  IdEstadoSII = " & EDTESII_DESCONOCIDO
'         Q1 = Q1 & ", GlosaSII = ' '"
'         Q1 = Q1 & ", ErrorSII = '" & ParaSQL(Error) & "'"
'         Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
'      End If
'
'      Call ExecSQL(DbMain, Q1)
'   End If
'
'   SaveEstadoEnviadoDTE = Rc
'End Function

Public Function FillDTEStruct(ByVal IdDTE As Long, DTE As DTE_t) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim iDTE As DTE_t
   
   FillDTEStruct = False
   
   If IdDTE = 0 Then
      Exit Function
   End If
   
   Q1 = "SELECT * FROM DTE WHERE IdDTE =" & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
    
   
   If Not Rs.EOF Then
      DTE.IdDTE = IdDTE
      DTE.IdEmpresa = gEmpresa.Id
      DTE.TipoDoc = vFld(Rs("TipoDoc"))
      DTE.TipoLib = vFld(Rs("TipoLib"))
      DTE.CodDocSII = vFld(Rs("CodDocSII"))
      DTE.Folio = vFld(Rs("Folio"))
      DTE.Fecha = vFld(Rs("Fecha"))
      DTE.IdEntidad = vFld(Rs("IdEntidad"))
      DTE.Rut = vFld(Rs("Rut"))
      DTE.Total = vFld(Rs("Total"))
      DTE.TrackID = vFld(Rs("TrackID"))
      
      FillDTEStruct = True
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function GetCodigoEstadoSII(ByVal EstadoSII As String) As Integer
   Dim i As Integer
   
   For i = 1 To UBound(gEstadoDTESII)
      If EstadoSII = gEstadoDTESII(i) Then
         GetCodigoEstadoSII = i
         Exit Function
      End If
   Next i
   GetCodigoEstadoSII = 0
         
End Function

Public Function ExitDemoFact() As Boolean
   Dim Rs As Recordset, Q1 As String
   
   Call AddDebug("En ExitDemoFact")
   
   If W.InDesign Then
      Exit Function
   End If
      
   Call AddDebug("ExitDemoFact: gAppCode.Demo=" & gAppCode.Demo)
   If gAppCode.Demo = False Then
      ExitDemoFact = False
      Exit Function
   End If

   Call AddDebug("ExitDemoFact: a Select RUT")
   ' Sólo deberían estar estos RUTs
   Q1 = "SELECT Rut FROM Empresas WHERE Rut NOT IN ('1','2','3')"
   Set Rs = OpenRs(DbMain, Q1)
   ExitDemoFact = Not Rs.EOF
   Call CloseRs(Rs)
      
End Function

Public Function GenDbZip(ByVal FnEmpr As String, ByVal ZipFile As String) As String
   'Dim ZipFile As String
   Dim zOpt As ZipOPT_t
   Dim zFiles As ZIPnames_t
   Dim zFnc As ZIPUSERFUNCTIONS_t
   Dim Rc As Long, KBytes As Long
   Dim Filename As String
   Dim i As Integer, nFiles As Integer

   nFiles = 1
   zFiles.zFiles(0) = gDbPath & "\" & BD_COMUNDTE

   zFiles.zFiles(1) = gDbPath & "\Empresas\" & FnEmpr
   nFiles = nFiles + 1
   
   zOpt.Date = vbNullString
   zOpt.flevel = Asc(9)  ' Compression Level (0 - 9)
   zOpt.szRootDir = gDbPath

   i = Len(FnEmpr)

   If Len(ZipFile) < 1 Then
      Filename = Left(FnEmpr, i - 4) & "_" & Format(Now, "yymmdd") & ".zip"
      ZipFile = W.TmpDir & "\" & Filename
   End If
   
   On Error Resume Next


   Rc = VBZip32(ZipFile, nFiles, zFiles, zOpt, zFnc)
   
   If Err Then
      MsgErr "No se puede generar el archivo " & ZipFile
      Exit Function
   End If
   
   If Rc Then
      Call AddLog("GenZip: Error " & Rc & " al generar el archivo " & ZipFile & ", DLL_Err=" & Err.LastDllError)
      MsgBox1 "Error " & Rc & " al generar el archivo " & ZipFile, vbCritical
   Else
      KBytes = FileLen(ZipFile) / 1024
'      Tx_XlsFile = Tx_XlsFile & vbCrLf & lZipFile & vbCrLf & "Tamaño: " & Format(KBytes, NUMFMT) & " KBytes"
      
      GenDbZip = ZipFile
   End If

End Function
''LLena una grilla como una fillcombo (Está en PamFGrid)
'Public Function FillGrLista(Qry As String, Gr As Control, ByVal iCol As Integer, wCol() As Byte)
'   Dim Rs As Recordset
'   Dim i As Integer, c As Integer
'
'   Set Rs = OpenRs(DbMain, Qry)
'
'   i = Gr.FixedRows
'   Gr.rows = i
'
'   For c = 0 To iCol
''      wCol(c) = Rs(c).Size
'      wCol(c) = FldSize(Rs(c))
'   Next c
'
'   Do While Rs.EOF = False
'      Gr.rows = i + 1
'
'      For c = 0 To iCol
'         Gr.TextMatrix(i, c) = vFld(Rs(c))
'      Next c
'
'      i = i + 1
'      Rs.MoveNext
'
'  Loop
'  Call CloseRs(Rs)
'  Call FGrVRows(Gr)
'
'End Function

Public Function GetIdEntidad(ByVal Rut As String, Nombre As String, NotValidRut As Boolean) As Long
   Dim Q1 As String
   Dim Rs As Recordset
   
   Rut = Trim(Rut)
   
   Q1 = "SELECT IdEntidad, Nombre, Rut, NotValidRut FROM Entidades WHERE Rut = '" & vFmtCID(Rut) & "' OR Rut = '" & Rut & "' ORDER BY Rut Desc"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      
      If vFld(Rs("NotValidRut")) <> 0 Then   'es RUT extranjero
         
         If vFld(Rs("Rut")) = Rut Then
            GetIdEntidad = vFld(Rs("IdEntidad"))
            Nombre = vFld(Rs("Nombre"))
            NotValidRut = vFld(Rs("NotValidRut"))
         Else
            GetIdEntidad = 0
            Nombre = ""
         End If
      
      Else
         
         If vFld(Rs("Rut")) = vFmtCID(Rut) Then
            GetIdEntidad = vFld(Rs("IdEntidad"))
            Nombre = vFld(Rs("Nombre"))
            NotValidRut = vFld(Rs("NotValidRut"))
         Else
            GetIdEntidad = 0
            Nombre = ""
         End If
      End If
      
   Else
      GetIdEntidad = 0
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function AddEntidad(ByVal Rut As String, ByVal RazonSocial As String, IdEntidad As Long) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim codigo As String

   Set Rs = DbMain.OpenRecordset("Entidades", dbOpenTable)
   Rs.AddNew
   
   IdEntidad = Rs("idEntidad")
   
   Rs("NotValidRut") = 0
   Rs("RUT") = vFmtCID(Rut)
   
   Rs.Update
   Rs.Close
   
   If Err Then
      IdEntidad = 0
      AddEntidad = False
      Exit Function
   End If
   
   codigo = vFmtRut(Rut)
      
   Q1 = "UPDATE Entidades SET "
   Q1 = Q1 & "  Nombre='" & ParaSQL(RazonSocial) & "'"
   Q1 = Q1 & ", Codigo='" & codigo & "'"
   Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
   
   Call ExecSQL(DbMain, Q1)
   
   AddEntidad = True
   
End Function

