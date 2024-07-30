VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LPFacturación - Fairware Ltda."
   ClientHeight    =   5085
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9810
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Download 
      Height          =   735
      Left            =   4200
      TabIndex        =   35
      Top             =   1260
      Visible         =   0   'False
      Width           =   4275
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Descargando archivo..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   36
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Fr_Invisible 
      Caption         =   "Invisibles"
      Height          =   2355
      Left            =   2580
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   5715
      Begin VB.PictureBox Pc_Lupa2 
         AutoSize        =   -1  'True
         Height          =   285
         Left            =   2220
         Picture         =   "FrmMain.frx":08CA
         ScaleHeight     =   225
         ScaleWidth      =   180
         TabIndex        =   34
         Top             =   660
         Visible         =   0   'False
         Width           =   240
      End
      Begin FlexEdGrid3.FEd3Grid FEd3Grid1 
         Height          =   555
         Left            =   480
         TabIndex        =   31
         Top             =   540
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   979
         Cols            =   2
         Rows            =   2
         FixedCols       =   1
         FixedRows       =   1
         ScrollBars      =   3
         AllowUserResizing=   0
         HighLight       =   1
         SelectionMode   =   0
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   -1  'True
         Locked          =   0   'False
      End
      Begin VB.PictureBox Pc_DocGreen 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1860
         Picture         =   "FrmMain.frx":0C75
         ScaleHeight     =   195
         ScaleWidth      =   180
         TabIndex        =   30
         Top             =   660
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Pc_Doc 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1500
         Picture         =   "FrmMain.frx":101D
         ScaleHeight     =   195
         ScaleWidth      =   180
         TabIndex        =   29
         Top             =   660
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Pc_Lupa 
         AutoSize        =   -1  'True
         Height          =   270
         Left            =   1500
         Picture         =   "FrmMain.frx":13C5
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   26
         Top             =   300
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox Pc_Flecha 
         AutoSize        =   -1  'True
         Height          =   150
         Left            =   1920
         Picture         =   "FrmMain.frx":173A
         ScaleHeight     =   90
         ScaleWidth      =   135
         TabIndex        =   25
         Top             =   300
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Timer Tmr_ChkDTE 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   1680
         Top             =   1200
      End
      Begin VB.Timer Tm_ChkUsr 
         Interval        =   60000
         Left            =   2220
         Top             =   1200
      End
      Begin VB.PictureBox Pc_Nota 
         AutoSize        =   -1  'True
         Height          =   135
         Left            =   2280
         Picture         =   "FrmMain.frx":17A8
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Timer Tmr_ChkActive 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   2700
         Top             =   1200
      End
      Begin MSComDlg.CommonDialog Cm_ComDlg 
         Left            =   420
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog Cm_PrtDlg 
         Left            =   1020
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FlexEdGrid2.FEd2Grid FEd2Grid1 
         Height          =   495
         Left            =   180
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   873
         Cols            =   2
         Rows            =   2
         FixedCols       =   1
         FixedRows       =   1
         ScrollBars      =   3
         AllowUserResizing=   0
         HighLight       =   1
         SelectionMode   =   0
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   -1  'True
         Locked          =   0   'False
      End
      Begin MSComDlg.CommonDialog Cm_FileDlg 
         Left            =   3360
         Top             =   1140
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Image Im_Down 
         BorderStyle     =   1  'Fixed Single
         Height          =   105
         Left            =   1200
         Picture         =   "FrmMain.frx":1810
         Top             =   300
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image Im_Orden 
         Height          =   270
         Left            =   180
         Picture         =   "FrmMain.frx":189E
         Top             =   1200
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   1320
      TabIndex        =   14
      Top             =   4500
      Width           =   7155
      Begin VB.CommandButton Bt_Equivalencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Picture         =   "FrmMain.frx":1BF5
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Equivalencias"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_ConvMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1140
         Picture         =   "FrmMain.frx":2045
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Convertir moneda"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calendar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Picture         =   "FrmMain.frx":24CD
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Calendario"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Indices 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         Picture         =   "FrmMain.frx":290B
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Valores e Índices"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Picture         =   "FrmMain.frx":2D13
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Calculadora"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label La_demo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "DEMO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   330
         Left            =   6180
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame Fr_Right 
      Height          =   3780
      Left            =   8520
      TabIndex        =   9
      Top             =   1260
      Width           =   1275
      Begin VB.CommandButton Bt_Exportar 
         Caption         =   "Exportar"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":3090
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2880
         Width           =   1155
      End
      Begin VB.CommandButton Bt_RepVentas 
         Caption         =   "Reporte de Ventas"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":3519
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Lista de Documentos Electrónicos recibidos"
         Top             =   1980
         Width           =   1155
      End
      Begin VB.CommandButton Bt_DTEEmitidos 
         Caption         =   "DTE Emitidos"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":3B11
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Lista de Documentos Electrónicos emitidos"
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton Bt_DTERecibidos 
         Caption         =   "DTERecibidos"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":40D6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Emitir Libros Electrónicos"
         Top             =   1080
         Width           =   1155
      End
   End
   Begin VB.Frame Fr_Left 
      Height          =   3780
      Left            =   60
      TabIndex        =   4
      Top             =   1260
      Width           =   1275
      Begin VB.CommandButton Bt_Emp 
         Caption         =   "Empresa"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":4667
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Datos Empresa"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Entidades 
         Caption         =   "Entidades"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":4D94
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Administración de Entidades Relacionadas"
         Top             =   1980
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Productos 
         Caption         =   "Prod. y Serv."
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":5497
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Administración de Lista de Productos"
         Top             =   2880
         Width           =   1155
      End
      Begin VB.CommandButton Bt_NewDTE 
         Caption         =   "Emitir DTE"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":5B0C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Emitir Documentos Electrónicos"
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Frame Fr_Empresa 
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.Label Lb_Tel 
         Alignment       =   1  'Right Justify
         Caption         =   "341 5788       205 4335"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   7500
         TabIndex        =   22
         Top             =   960
         Width           =   2025
      End
      Begin VB.Label Lb_Emp 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         ForeColor       =   &H00A67300&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Lb_Dir 
         AutoSize        =   -1  'True
         Caption         =   "El Belloto 3942, P1"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Lb_Emp 
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos:"
         ForeColor       =   &H00A67300&
         Height          =   195
         Index           =   2
         Left            =   6720
         TabIndex        =   19
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Lb_Emp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "RUT: "
         ForeColor       =   &H00A67300&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Lb_RUT 
         AutoSize        =   -1  'True
         Caption         =   "77.049.060-K"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Lb_Empresa 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fairware Ltda."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   9495
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   1320
      Picture         =   "FrmMain.frx":60E5
      ScaleHeight     =   3195
      ScaleWidth      =   7095
      TabIndex        =   28
      Top             =   1320
      Width           =   7155
   End
   Begin VB.Menu M_Empresa 
      Caption         =   "Empresa"
      Begin VB.Menu ME_SelEmp 
         Caption         =   "Abrir..."
      End
      Begin VB.Menu ME_EditEmp 
         Caption         =   "Modificar datos empresa..."
      End
      Begin VB.Menu ME_ConfigFactElect 
         Caption         =   "Configurar conexión Facturación Electrónica..."
      End
      Begin VB.Menu ME_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ME_MantEmp 
         Caption         =   "Mantención empresas..."
      End
      Begin VB.Menu ME_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu ME_LastOpen 
         Caption         =   "&1"
         Index           =   0
      End
      Begin VB.Menu ME_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu ME_Exit 
         Caption         =   "Salir           Alt+F4"
      End
   End
   Begin VB.Menu M_DatBas 
      Caption         =   "Datos Básicos"
      Begin VB.Menu MD_Entidades 
         Caption         =   "Entidades Relacionadas..."
      End
      Begin VB.Menu MD_Productos 
         Caption         =   "Productos y Servicios..."
      End
   End
   Begin VB.Menu M_DTE 
      Caption         =   "DTE"
      Begin VB.Menu MD_NewDTE 
         Caption         =   "Emitir Documentos Electrónicos..."
      End
      Begin VB.Menu MD_DTEEmitidos 
         Caption         =   "DTE Emitidos..."
      End
      Begin VB.Menu MD_DTERecibidos 
         Caption         =   "DTE Recibidos..."
      End
   End
   Begin VB.Menu M_Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu MR_RepVentas 
         Caption         =   "Reporte de Ventas..."
      End
   End
   Begin VB.Menu M_Procesos 
      Caption         =   "Procesos"
      Begin VB.Menu MP_ImportDTERec 
         Caption         =   "Importar DTE..."
      End
      Begin VB.Menu MP_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu MP_ExportarLPConta 
         Caption         =   "Exportar a LPContabilidad..."
      End
      Begin VB.Menu MP_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MP_ImportEnt 
         Caption         =   "Importar Entidades desde LPContabilidad..."
      End
      Begin VB.Menu MP_ImportEntTexto 
         Caption         =   "Importar Entidades desde Archivo de Texto..."
      End
      Begin VB.Menu MP_ImpProd 
         Caption         =   "Importar Productos y Servicios..."
      End
   End
   Begin VB.Menu M_Mant 
      Caption         =   "Mantención"
      Begin VB.Menu MM_DocRef 
         Caption         =   "Documentos de Referencia..."
      End
      Begin VB.Menu MM_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MM_Paises 
         Caption         =   "Paises..."
      End
      Begin VB.Menu MM_Puertos 
         Caption         =   "Puertos..."
      End
      Begin VB.Menu MM_ClauVenta 
         Caption         =   "Cláusulas de Compraventa..."
      End
      Begin VB.Menu MM_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MM_Transportistas 
         Caption         =   "Transportistas..."
         Begin VB.Menu MM_Vehiculos 
            Caption         =   "Vehículos..."
         End
         Begin VB.Menu MM_Conductores 
            Caption         =   "Conductores..."
         End
      End
      Begin VB.Menu MM_Vendedores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu MM_DetaFormaPago 
         Caption         =   "Detalle Forma de Pago"
      End
      Begin VB.Menu MM_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MC_Monedas 
         Caption         =   "Monedas"
         Begin VB.Menu MC_Equivalencias 
            Caption         =   "Equivalencias..."
         End
         Begin VB.Menu MC_CalcMonedas 
            Caption         =   "Calculadora de Monedas..."
         End
         Begin VB.Menu MM_Monedas 
            Caption         =   "Mantención de Monedas..."
         End
      End
      Begin VB.Menu MC_Indices 
         Caption         =   "Valores e Índices..."
      End
   End
   Begin VB.Menu M_Config 
      Caption         =   "Configuración"
      Begin VB.Menu MC_DatosOficina 
         Caption         =   "Datos Oficina..."
      End
      Begin VB.Menu MC_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MS_CambiarClave 
         Caption         =   "Cambiar Clave..."
      End
      Begin VB.Menu MC_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MS_Usuarios 
         Caption         =   "Usuarios..."
      End
      Begin VB.Menu MS_Perfiles 
         Caption         =   "Perfiles..."
      End
   End
   Begin VB.Menu M_Sistema 
      Caption         =   "Sistema"
      Begin VB.Menu MS_SetupPrt 
         Caption         =   "Preparar Impresora..."
      End
      Begin VB.Menu MS_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MS_Respaldo 
         Caption         =   "Respaldo..."
      End
      Begin VB.Menu MS_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MS_Desbloquear 
         Caption         =   "Desbloquear conexión..."
      End
      Begin VB.Menu MS_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MC_SolicCod 
         Caption         =   "Licenciar Producto..."
      End
      Begin VB.Menu MC_AutEquipos 
         Caption         =   "Ingresar Código de Licencia..."
      End
      Begin VB.Menu MS_Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu MS_DownLast 
         Caption         =   "Descargar actualización..."
      End
      Begin VB.Menu MS_Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu MS_MantDB 
         Caption         =   "Mantención Base de Datos"
         Begin VB.Menu MS_Reparar 
            Caption         =   "Reparar..."
         End
         Begin VB.Menu MS_Compactar 
            Caption         =   "Compactar..."
         End
      End
   End
   Begin VB.Menu MH_Help 
      Caption         =   "&Ayuda"
      Begin VB.Menu MH_HlpBackup 
         Caption         =   "Ayuda Respaldo..."
      End
      Begin VB.Menu MH_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MH_RepErr 
         Caption         =   "Reporte de Problema..."
      End
      Begin VB.Menu MH_Export 
         Caption         =   "Exportar Empresa..."
      End
      Begin VB.Menu MH_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MH_AcercaDe 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const nLAST = 5
Dim LastOpen(nLAST) As LastOpen_t
Dim lnChkDTE As Integer

Dim FrmActivate As Boolean

Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Call Frm.FView(Valor)
      
   Set Frm = Nothing

End Sub

Private Sub Bt_DTEEmitidos_Click()
   Call MD_DTEEmitidos_Click
End Sub

Private Sub Bt_DTERecibidos_Click()

   Call MD_DTERecibidos_Click
   
End Sub
Private Sub bt_Equivalencia_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub Bt_Exportar_Click()
   MP_ExportarLPConta_Click
End Sub

Private Sub Bt_Indices_Click()
   Dim Frm As FrmIPC
   
   Set Frm = New FrmIPC
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub Bt_NewDTE_Click()
   Dim Frm As FrmSelEmitirDTE
   
   If gConectData.Proveedor = 0 Or gConectData.Usuario = "" Or gConectData.Clave = "" Or gConectData.ClaveCert = "" Then
      MsgBox1 "Falta configurar la conexión de Facturación Electrónica. Vaya al Menú" & vbCrLf & vbCrLf & "Empresa>>Configurar conexión Facturación Electrónica" & vbCrLf & vbCrLf & "y complete todos los datos que ahí se solicitan.", vbExclamation
      Exit Sub
   End If
   
   
   If gEmpresa.RazonSocial = "" Or gEmpresa.Direccion = "" Or gEmpresa.Comuna = "" Or gEmpresa.Ciudad = "" Or gEmpresa.Giro = "" Or gEmpresa.CodActEcono = "" Then
      MsgBox1 "Falta ingresar algunos datos de la Empresa para poder emitir un DTE." & vbCrLf & vbCrLf & "Vaya al botón Empresa y complete los datos requeridos.", vbExclamation
      Exit Sub
   End If
   
   Set Frm = New FrmSelEmitirDTE
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub Bt_Emp_Click()
   Call ME_EditEmp_Click
   
End Sub

Private Sub Bt_Entidades_Click()
   Call MD_Entidades_Click
End Sub

Private Sub Bt_Productos_Click()
   Call MD_Productos_Click
End Sub

Private Sub Bt_RepVentas_Click()
   Call MR_RepVentas_Click
End Sub

Private Sub Form_Activate()
   Dim Rs As Recordset
   Dim Q1 As String
   
   Call AddDebug("FrmMain_Activate: Antes de ExitDemoFact")
   
'   If ExitDemoFact() Then se hace en el Timer
'      Unload Me
'   End If
   
   Call AddDebug("FrmMain_Activate: Antes de FrmActivate")
   
   If FrmActivate = True Then
      Exit Sub
   End If
         
   Call AddDebug("FrmMain_Activate: Después de FrmActivate")
   
   FrmActivate = True
         
   Call ShowMsgBackup
    
End Sub

Private Sub Form_Load()

   Call AddDebug("FrmMain_Load: Antes de gFrmMain = Me")

   Set gFrmMain = Me
   
   On Error Resume Next

   Call AddDebug("FrmMain_Load: Antes de FEd2Grid1")

   FEd2Grid1.TextMatrix(0, 0) = "$1#2¿P" ' No borrar
   FEd2Grid1.TextMatrix(0, 0) = "" ' No borrar

   FEd3Grid1.TextMatrix(0, 0) = "$7#3?F#" ' No borrar
   FEd3Grid1.TextMatrix(0, 0) = "" ' No borrar


   If Err Then
      MsgErr "La versión del objeto FlexEdGrid2 no corresponde."
   End If
   
   Call ReadLastOpen

   Set gPrtDlg = Me.Cm_PrtDlg

   Call AddDebug("FrmMain_Load: Antes de FillDatosEmp")

   Call FillDatosEmp
   Call WriteLastOpen(O_EDIT)

   Call SetPrtData

   Call SetupPriv

   La_demo.Visible = gAppCode.Demo

'   Tmr_Chk.Enabled = (gAppCode.Demo = False)

   Tmr_ChkActive.Enabled = Not gFwChkActive


#If Inscr2 = 1 Then
'   MH_Sep2.Visible = False
'   MH_Inscrip.Visible = False
'   MH_DesInscr.Visible = False
#End If

   Call AddDebug("FrmMain_Load: nos vamos OK")
   
   Call BloquearFormularios(False)
   Call FormulariosXPerfil
End Sub

Private Sub FormulariosXPerfil()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim a As Long
   'Call ClearPriv
   
'   If IdxPerfil < 0 Then
'      Exit Sub
'   End If
   
   Q1 = "SELECT p.Privilegios FROM Usuarios as u LEFT JOIN  Perfiles p ON u.idperfil = p.idperfil WHERE u.IdUsuario = " & gUsuario.IdUsuario & " AND p.Privilegios  IS NOT NULL "
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = False Then
      For i = 0 To 5
         If TienePrivilegio(2 ^ i, Rs(0)) = True Then
            Call MostrarFormularios(2 ^ i)
         End If
      Next i
   Else
    Call BloquearFormularios(True)
   End If
   
   Call CloseRs(Rs)
   
'   If Ls_Priv.ListCount > 0 Then
'      Ls_Priv.TopIndex = 0
'      Ls_Priv.ListIndex = 0
'   End If
End Sub

Private Sub MostrarFormularios(indice As Long)

    Select Case indice
    
    Case PRVF_ADM_SIS
        M_Config.Enabled = True
    
    Case PRVF_CFG_EMP
        M_Empresa.Enabled = True
        Bt_Emp.Enabled = True
    
    Case PRVF_ADM_EMPRESA
        M_DatBas.Enabled = True
        Bt_Entidades.Enabled = True
        Bt_Productos.Enabled = True
        
    Case PRVF_ADM_EXP
        M_Procesos.Enabled = True
        Bt_Exportar.Enabled = True
        
    Case PRVF_EMITIR_FACT
        MD_NewDTE.Enabled = True
        Bt_NewDTE.Enabled = True
        Bt_DTEEmitidos.Enabled = True
    
    Case PRVF_ADM_FACT
        MD_DTEEmitidos.Enabled = True
        MD_DTERecibidos.Enabled = True
        Bt_DTERecibidos.Enabled = True
        
    Case PRVF_MANT_DATOS
        M_Mant.Enabled = True
    Case Else
    End Select

End Sub
Private Sub BloquearFormularios(Habilitar As Boolean)

M_Config.Enabled = Habilitar

M_Empresa.Enabled = Habilitar
Bt_Emp.Enabled = Habilitar

M_DatBas.Enabled = Habilitar
Bt_Entidades.Enabled = Habilitar
Bt_Productos.Enabled = Habilitar

M_Procesos.Enabled = Habilitar
Bt_Exportar.Enabled = Habilitar

MD_NewDTE.Enabled = Habilitar
Bt_NewDTE.Enabled = Habilitar
Bt_DTEEmitidos.Enabled = Habilitar

MD_DTEEmitidos.Enabled = Habilitar
MD_DTERecibidos.Enabled = Habilitar
Bt_DTERecibidos.Enabled = Habilitar

M_Mant.Enabled = Habilitar

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Call ContUnregisterPc(1)
   
   Call CloseDb(DbMain)
   Call CheckRs(True)
   Cancel = True
   End
   
End Sub
Private Sub FillDatosEmp()
   Dim Cntrl As Control
   Dim i As Integer, Dt As Long ' , Hoy As Long
   
   If gEmpresa.Id > 0 Then     'hay empresa seleccionada
   
      For i = 0 To 2
         Lb_Emp(i).Visible = True
      Next i
      
      Lb_RUT = FmtCID(gEmpresa.Rut)
      Lb_Dir = gEmpresa.Direccion
      Lb_Tel = gEmpresa.Telefono
      Lb_Empresa = gEmpresa.NombreCorto
'      Lb_Año = gEmpresa.Ano
'      Lb_Mes = Left(gNomMes(GetMesActual), 3)
      
'      Me.Caption = gEmpresa.NombreCorto & " - " & gEmpresa.Ano & " - " & gLpFactura
      Me.Caption = gEmpresa.NombreCorto & " - " & gLPFactura
      
      If gAppCode.Demo Then
         Me.Caption = Me.Caption & " - D" & "E" & "M" & "O"
      End If
      
      If gNuevaInstancia Then
         Me.Caption = Me.Caption & " [R]"
      End If
      
'      Lb_Cierre.Visible = gEmpresa.FCierre <> 0
      
      Tmr_ChkDTE.Enabled = True
      lnChkDTE = 0
'      Hoy = Int(Now)
'      Dt = Val(GetIniString(gCfgFile, "Import-" & gEmpresa.Rut, "FDteRec")) ' 2 ago 2019: para saber cuando leyo por última vez
'      If Dt <= 0 Or (Dt > 0 And Dt < Hoy - 5) Then
'         If AcpCountArchFCompra(Val(gEmpresa.Rut)) > 0 Then
'            If Dt = 0 Then
'               Msg = "Hace más de 10"
'            Else
'               Msg = "Hace " & (Hoy - Dt)
'            End If
'
'            MsgBox1 "ATENCIÓN" & vbCrLf & Msg & " días que no obtiene los DTE Recibidos desde el menú Procesos." & vbCrLf & "Pasado un tiempo razonable estos datos van eliminando.", vbExclamation
'         End If
'      End If
   
   Else
      For i = 0 To 2
         Lb_Emp(i).Visible = False
      Next i
         
      Lb_RUT = ""
      Lb_Dir = ""
      Lb_Tel = ""
      Lb_Empresa = gLPFactura
'      Lb_Año = ""
'      Lb_Mes = ""
      
      Me.Caption = gLPFactura
      
      If gAppCode.Demo Then
         Me.Caption = Me.Caption & " - D" & "E" & "M" & "O"
      End If
      
      If gNuevaInstancia Then
         Me.Caption = Me.Caption & " [R]"
      End If
      
'      Lb_Cierre.Visible = False
      
   End If
     
End Sub

Private Sub SetupPriv()
   Dim bool As Boolean
   
   bool = True

   If gEmpresa.Id = 0 Then
      bool = False
   End If
   
   ME_EditEmp.Enabled = bool
   ME_ConfigFactElect.Enabled = bool
   Bt_Emp.Enabled = bool
   MD_Entidades.Enabled = bool
   Bt_Entidades.Enabled = bool
   MD_Productos.Enabled = bool
   Bt_Productos.Enabled = bool
   MD_NewDTE.Enabled = bool
   Bt_NewDTE.Enabled = bool
   Bt_DTEEmitidos.Enabled = bool
   MD_DTEEmitidos.Enabled = bool
   Bt_DTEEmitidos.Enabled = bool
   MR_RepVentas.Enabled = bool
   Bt_RepVentas.Enabled = bool
   MP_ExportarLPConta.Enabled = bool
   Bt_Exportar.Enabled = bool
   MP_ImportEnt.Enabled = bool
   MP_ImportEntTexto.Enabled = bool
   MP_ImpProd.Enabled = bool
   MH_Export.Enabled = bool
   
   If bool Then
      If Not ChkPriv(PRVF_ADM_FACT) Then
         MD_Entidades.Enabled = False
         MD_Productos.Enabled = False
      End If
      
      If Not ChkPriv(PRVF_EMITIR_FACT) Then
         MD_NewDTE.Enabled = False
      End If
   End If
End Sub

Private Sub MC_CalcMonedas_Click()
   Dim Frm As FrmConverMoneda
   
   Set Frm = New FrmConverMoneda
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MC_DatosOficina_Click()
   Dim Frm As FrmOficina
   
   Set Frm = New FrmOficina
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MC_Indices_Click()
   Call Bt_Indices_Click
End Sub

Private Sub MC_SolicCod_Click()
   Dim Frm As FrmEquiposAut
   
   If gOficina.Rut = "" Then
      MsgBox1 "Debe ingresar el RUT de su empresa en el menú Configuración >> Datos Oficina.", vbExclamation
      Exit Sub
   End If

   Set Frm = New FrmEquiposAut
   Call Frm.Solicitud
   Set Frm = Nothing

End Sub

Private Sub MC_AutEquipos_Click()
   Dim Frm As FrmEquiposAut
   
   If gOficina.Rut = "" Then
      MsgBox1 "Debe ingresar el RUT de su empresa en el menú Configuración >> Datos Oficina.", vbExclamation
      Exit Sub
   End If
   
   Set Frm = New FrmEquiposAut
   Call Frm.Admin
   Set Frm = Nothing

   Call SetCaption

End Sub
Private Sub SetCaption()

   Me.Caption = gLPFactura

   If gAppCode.Demo Then
      Me.Caption = Me.Caption & " - DEMO"
   End If


End Sub

Private Sub MD_DTERecibidos_Click()
   Dim Frm As FrmDTERecibidos
   
   Set Frm = New FrmDTERecibidos
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MD_Entidades_Click()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmEntidades
   Call Frm.FEdit(True)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub


Private Sub MD_NewDTE_Click()
   Call Bt_NewDTE_Click
End Sub

Private Sub MD_Productos_Click()
   Dim Frm As FrmProductos
   
   Set Frm = New FrmProductos
   Me.MousePointer = vbHourglass
   Call Frm.FEdit
   Set Frm = Nothing
   Me.MousePointer = vbDefault
   
End Sub

Private Sub ME_ConfigFactElect_Click()
   Dim Frm As FrmConfigConect
   
   Set Frm = New FrmConfigConect
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub ME_EditEmp_Click()
   Dim Frm As FrmEmpresa
   
   Set Frm = New FrmEmpresa
   MousePointer = vbHourglass
   
   If Frm.FEdit(gEmpresa.Id) = vbOK Then
      Lb_Dir = gEmpresa.Direccion
      Lb_Tel = gEmpresa.Telefono
   End If
   
   MousePointer = vbDefault
   Set Frm = Nothing
   
End Sub

Private Sub ME_Exit_Click()
   Unload Me
End Sub

Private Sub ME_MantEmp_Click()
   Dim Frm As FrmMantEmpresas
   
   If gEmpresa.Id > 0 Then
      If MsgBox1("ATENCIÖN: Antes de continuar se debe cerrar la empresa con la cual está trabajando actualmente." & vbCrLf & vbCrLf & "Desea continuar?", vbYesNoCancel + vbQuestion) <> vbYes Then
         Exit Sub
      End If
      Call CerrarEmp
   End If

   Set Frm = New FrmMantEmpresas
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub ME_SelEmp_Click()
   Dim Rut As String
   Dim Frm As FrmSelEmpresas
   Dim Rc As Integer
   Dim BoolIniEmpresa As Boolean
   Dim DbMainOld As Database
   Dim gCurEmp As Empresa_t
   Dim Q1 As String
   Dim Rs As Recordset
   Dim PlanVacio As Boolean
'   Dim FrmConfig As FrmConfig
   
   
   BoolIniEmpresa = False
      
   gCurEmp = gEmpresa
      
   Do While BoolIniEmpresa = False
      Set Frm = New FrmSelEmpresas
      Rc = Frm.FSelect
      Set Frm = Nothing
      
      If Rc = vbOK Then
      
         Set DbMainOld = DbMain     'db de la empresa actual
         Set DbMain = Nothing       'para que no la cierre
         
         BoolIniEmpresa = IniEmpresa()
         
         If BoolIniEmpresa = False Then   'falló, dejamos la Db actual
            Set DbMain = DbMainOld
            gEmpresa = gCurEmp
         Else
            Call CloseDb(DbMainOld)       'abrió otra db, cerramos la anterior
         End If
      
      Else
         BoolIniEmpresa = True
      End If
      
   Loop
   
   If Rc = vbOK Then
      Call FillDatosEmp
      Call WriteLastOpen(O_EDIT)
      Call SetPrtData
      Call SetupPriv
   End If
   
'   Q1 = "SELECT IdCuenta FROM Cuentas"
'   Set Rs = OpenRs(DbMain, Q1)
'   PlanVacio = (Rs.EOF = True)
'   Call CloseRs(Rs)
'
'   If PlanVacio Then    'no hay cuentas definidas
'      MsgBox1 "Se ha detectado que no está definido el Plan de Cuentas para esta empresa." & vbNewLine & vbNewLine & "La ventana de Configuración Inicial le permitirá definir el Plan de Cuentas y otros elementos básicos para trabajar con el sistema.", vbInformation + vbOKOnly
'      Set FrmConfig = New FrmConfig
'      FrmConfig.Show vbModal
'      Set FrmConfig = Nothing
'
'      MsgBox1 "Recuerde configurar las Razones Financieras para esta empresa, " & vbCrLf & vbCrLf & "utilizando la opción que provee el sistema, bajo el menú 'Configuración'", vbOKOnly + vbInformation
'
'   End If

End Sub

Private Sub MH_AcercaDe_Click()
   Dim Frm As FrmAbout
   
   Set Frm = New FrmAbout
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MH_HlpBackup_Click()
   Dim Frm As FrmHlpBackup
   
   Set Frm = New FrmHlpBackup
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MM_ClauVenta_Click()
   Dim Frm As FrmMantClauVenta
   
   Set Frm = New FrmMantClauVenta
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MM_Conductores_Click()
   Dim Frm As FrmMantConductores
   
   Set Frm = New FrmMantConductores
   Frm.Show vbModal
   Set Frm = Nothing


End Sub

Private Sub MM_DetaFormaPago_Click()
Dim Frm As FrmDetFormPago
   
   Set Frm = New FrmDetFormPago
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MM_DocRef_Click()
   Dim Frm As FrmMantTipoDocRef
   
   Set Frm = New FrmMantTipoDocRef
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MM_Monedas_Click()
   Dim Frm As FrmMantMonedas
   
   Set Frm = New FrmMantMonedas
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MM_Paises_Click()
   Dim Frm As FrmMantPaises
   
   Set Frm = New FrmMantPaises
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MM_Puertos_Click()
   Dim Frm As FrmMantPuertos
   
   Set Frm = New FrmMantPuertos
   Frm.Show vbModal
   Set Frm = Nothing


End Sub

Private Sub MM_Vehiculos_Click()
   Dim Frm As FrmMantVehiculos
   
   Set Frm = New FrmMantVehiculos
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MM_Vendedores_Click()
Dim Frm As FrmMantVendedor
   
   Set Frm = New FrmMantVendedor
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MP_ExportarLPConta_Click()
   Dim Frm As FrmExport
   
   Set Frm = New FrmExport
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MP_ImportDTERec_Click()
'   Dim Frm As FrmImpDTERecibidos
    Dim Frm As FrmImportDocs
'
'   If gConectData.Proveedor = 0 Or gConectData.Usuario = "" Or gConectData.Clave = "" Or gConectData.ClaveCert = "" Then
'      MsgBox1 "Falta configurar la conexión de Facturación Electrónica. Vaya al Menú" & vbCrLf & vbCrLf & "Empresa>>Configurar conexión Facturación Electrónica" & vbCrLf & vbCrLf & "y complete todos los datos que ahí se solicitan.", vbExclamation
'      Exit Sub
'   End If
'
'   Set Frm = New FrmImpDTERecibidos
'   Frm.Show vbModal
'   Set Frm = Nothing

   Set Frm = New FrmImportDocs
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MP_ImportEnt_Click()
   Dim Frm As FrmImportEnt
   
   Set Frm = New FrmImportEnt
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub MP_ImpProd_Click()
   Dim Frm As FrmImportProd
   
   Set Frm = New FrmImportProd
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MD_DTEEmitidos_Click()
   Dim Frm As FrmAdmDTE
   
   Set Frm = New FrmAdmDTE
   Call Frm.FView
   Set Frm = Nothing

End Sub

Private Sub MR_RepVentas_Click()
   Dim Frm As FrmRepVentasProd
   
   Set Frm = New FrmRepVentasProd
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MS_CambiarClave_Click()
   Dim Frm As FrmCambioClave
   
   Set Frm = New FrmCambioClave
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MS_Desbloquear_Click()
   Dim Frm As FrmDesbloquear
   
   Set Frm = New FrmDesbloquear
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MS_DownLast_Click()
   Static bDown As Boolean
   
   If bDown Then
      Exit Sub
   End If
   bDown = True

   MousePointer = vbHourglass
   Fr_Download.Visible = True
   DoEvents

'   If Trim(gAppCode.Rut) = "" Then
'      gAppCode.Rut = gLicInfo.RutEmpr
'   End If
   
   Call FwDownLast(Me, Cm_FileDlg, APP_DEMO)

   Fr_Download.Visible = False
   MousePointer = vbDefault

   bDown = False


End Sub

Private Sub MS_Perfiles_Click()
   Dim Frm As FrmPerfiles
   
   Set Frm = New FrmPerfiles
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MS_Respaldo_Click()
   Dim Frm As FrmBackup
   
   If gEmpresa.Id > 0 Then
      If MsgBox1("ATENCIÖN: Antes de continuar se debe cerrar la empresa con la cual está trabajando actualmente." & vbCrLf & vbCrLf & "Desea continuar?", vbYesNoCancel + vbQuestion) <> vbYes Then
         Exit Sub
      End If
      Call CerrarEmp
   End If
   
   Set Frm = New FrmBackup
   Frm.Show vbModal
   Set Frm = Nothing

End Sub
Private Sub CerrarEmp()
   
   If gEmpresa.Id = 0 Then
      Exit Sub
   End If
   
   Call CloseDb(DbMain)
   
   gEmpresa.Rut = ""
   gEmpresa.NombreCorto = ""
   gEmpresa.Id = 0
   gEmpresa.Ano = 0
   'gEmpresa.FCierre = vFmt(LsAno.ItemData(LsAno.ListIndex))
   gEmpresa.FCierre = 0
   gEmpresa.FApertura = 0
      
'   gUsuario.idPerfil = 0
'   gUsuario.Priv = 0

   If OpenDbAdmFact() = False Then
      End
   End If
   
   Call FillDatosEmp
   Call SetupPriv
   
End Sub

Private Sub MS_SetupPrt_Click()
   Dim CurrPrt As String
   Dim Rc As Integer
   
   If PrepararPrt(Cm_PrtDlg) Then
   
      Call SetIniString(gIniFile, "Config", "Printer", Printer.DeviceName)
   Else
      Call FindPrinter(GetIniString(gIniFile, "Config", "Printer"), True)
    
   End If

End Sub

Private Sub MS_Usuarios_Click()
   Dim Frm As FrmUsuarios
   
   Set Frm = New FrmUsuarios
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub
Private Sub Bt_Calc_Click()
   Call Calculadora
End Sub

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   Call Frm.SelDate(Fecha)
   Set Frm = Nothing
   
End Sub

Private Sub MP_ImportEntTexto_Click()
   Dim Frm As FrmImportEntTxt
   
   Set Frm = New FrmImportEntTxt
   Frm.Show vbModal
   Set Frm = Nothing
   

End Sub


'Se agrega opción para exportar empresa
Private Sub MH_Export_Click()
   Dim FnEmpr As String, i As Integer, FnZip As String, Fn As String
   Static bExporting As Boolean
   
   If bExporting Then
      Exit Sub
   End If
   bExporting = True
   
   If MsgBox1("¡ATENCIÓN!" & vbLf & "Antes de exportar esta empresa debe verificar que no hayan usuarios conectados al sistema." & vbLf & "¿Desea continuar?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
      bExporting = False
      Exit Sub
   End If
   
   If gEmpresa.Id > 0 Then
      If MsgBox1("ATENCIÖN: Antes de continuar se debe cerrar la empresa para exportarla correctamente." & vbCrLf & vbCrLf & "Desea continuar?", vbYesNoCancel + vbQuestion) <> vbYes Then
         bExporting = False
         Exit Sub
      End If
      FnEmpr = DbMain.Name
      Call CerrarEmp
   Else
      Exit Sub
   End If


   MousePointer = vbHourglass
   DoEvents
    
   i = rInStr(FnEmpr, "\")
   
   If i <= 0 Then
      MousePointer = vbDefault
      bExporting = False
      Exit Sub
   End If
      
   FnEmpr = Mid(FnEmpr, i + 1)
   
   i = Len(FnEmpr)
   FnZip = Left(FnEmpr, i - 4) & "_" & Format(Now, "yymmdd") & ".zip"
   
   On Error Resume Next
   Cm_ComDlg.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
   Cm_ComDlg.CancelError = True
   Cm_ComDlg.Filename = FnZip
   Cm_ComDlg.InitDir = W.TmpDir
   Cm_ComDlg.ShowSave
   
   If Err.Number Then
      MousePointer = vbDefault
      bExporting = False
      Exit Sub
   End If
   
   FnZip = Cm_ComDlg.Filename
   
   Fn = GenDbZip(FnEmpr, FnZip)
   i = rInStr(Fn, "\")
   
   If Len(Fn) > 0 Then
      If MsgBox1("Se generó el archivo" & vbCrLf & Fn & vbCrLf & vbCrLf & "¿ Desea abrir la carpeta del archivo ?", vbInformation Or vbYesNo) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", Left(Fn, i), "", "", SW_SHOW)
      End If
   End If
   
   If OpenDbAdmFact() = False Then      'antes de llamar al backup se cierra la empresa
      End
   End If

   MousePointer = vbDefault
   bExporting = False
   
End Sub

Private Sub MC_Equivalencias_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Frm.FEdit (0)
   Set Frm = Nothing
   
End Sub


Private Sub MS_Reparar_Click()
   Dim DbPath As String
   
   If MsgBox1("Antes de realizar esta operación, verifique que no haya ningún usuario trabajando en el sistema." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   DbPath = DbMain.Name
   
   Call CloseDb(DbMain)
   
   If RepairDb(DbPath) Then
      If OpenDbEmpFact() = False Then
         End
      End If
      Me.MousePointer = vbDefault
   Else
      Unload Me
      End
   End If
End Sub
Private Sub MS_Compactar_Click()
   Dim ConnStr As String

   If MsgBox1("Antes de realizar esta operación, verifique que no haya ningún usuario trabajando en el sistema." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   
   'ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
   If CompactDb2(DbMain, True, gEmpresa.ConnStr) = 0 Then 'no hubo error
      If OpenDbEmpFact() = False Then
         End
      End If
   Else
      MsgBox1 "Problemas al tratar de compactar la base de datos.", vbExclamation + vbOKOnly
   End If
   
   Me.MousePointer = vbDefault

End Sub


Private Sub ME_LastOpen_Click(Index As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim idEmpr As Long
   Dim BoolIniEmpresa As Boolean
      
   If ExitDemo() Then
      Exit Sub
   End If
   
 '  End If
      
   MousePointer = vbHourglass
   DoEvents
   
   If LastOpen(Index).Id <> 0 Then
      
      idEmpr = LastOpen(Index).Id
      Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE IdEmpresa = " & idEmpr
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
   
         gEmpresa.Rut = vFld(Rs("Rut"))
         gEmpresa.NombreCorto = vFld(Rs("NombreCorto"))
         gEmpresa.Id = idEmpr
      
      Else
         idEmpr = 0
         
      End If
         
      Call CloseRs(Rs)
      
      If idEmpr > 0 Then
      
         Call CloseDb(DbMain)
   
         Call AddLog("Main: FrmSelEmpresas OK. A IniEmpresa")
            
         'Se abre la base de datos de la empresa y se inicializan sus datos básicos
         BoolIniEmpresa = IniEmpresa()
      
         Call AddLog("Main: IniEmpresa RC=" & BoolIniEmpresa)
         If BoolIniEmpresa = False Then
            If OpenDbAdmFact() = False Then
               End
            End If
            
            'seteamos los datos de la empresa para clase de impresion de grillas
            Call SetPrtData
            
         End If
      
      Else
         MousePointer = vbDefault
         LastOpen(Index).Id = -1
         ME_LastOpen(Index).Visible = False
         Call SetIniString(gIniFile, "LastOpen", str(Index), "")
         Exit Sub
      End If
         
      Call FillDatosEmp
      Call WriteLastOpen(O_EDIT)
      Call SetPrtData
      Call SetupPriv
      
   Else
      MsgBox1 "No se encontró la empresa '" & LastOpen(Index).Nombre & "', pudo haber sido eliminada", vbExclamation
      LastOpen(Index).Id = -1
      ME_LastOpen(Index).Visible = False
      Call SetIniString(gIniFile, "LastOpen", str(Index), "")

   End If
   
   MousePointer = vbDefault
   
End Sub

Private Sub ReadLastOpen()
   Dim i As Integer, j As Integer, k As Integer
   Dim Buf As String, Q1 As String, Rs As Recordset
   Dim Rc As Long
   
   j = 0
   For i = 0 To nLAST - 1
      Buf = GetIniString(gIniFile, "LastOpen", str(i))
      
      k = InStr(Buf, "|")
      If k > 0 Then
      
         LastOpen(j).Id = Val(Left(Buf, k - 1))
         
         Q1 = "SELECT NombreCorto FROM Empresas WHERE idEmpresa=" & LastOpen(j).Id
         Set Rs = OpenRs(DbMain, Q1)
         If Rs.EOF = False Then
            LastOpen(j).Nombre = FCase(vFld(Rs("NombreCorto")))
            j = j + 1
'         Else
'            LastOpen(j).Nombre = Trim(Mid(Buf, k + 1))
         End If
         Call CloseRs(Rs)
         
      End If
      
   Next i
   
   On Error Resume Next
   ' Ahora ponemos el menu
   For i = 0 To nLAST - 1
   
'      If i >= gMaxEmpr Then
'         Exit For
'      End If
   
      Load ME_LastOpen(i)
      
      ME_LastOpen(i).Caption = "&0"
      ME_LastOpen(i).Enabled = False
      ME_LastOpen(i).Visible = (i = 0)
    
      If Trim(LastOpen(i).Nombre) <> "" Then
         ME_LastOpen(i).Caption = "&" & i + 1 & "  " & FCase(LastOpen(i).Nombre)
         ME_LastOpen(i).Enabled = True
         ME_LastOpen(i).Visible = True
         
      End If
            
   Next i
      
End Sub

Private Sub WriteLastOpen(Oper As Integer)
   Dim i As Integer, j As Integer
   Dim Rc As Long
   Dim nTot As Integer
   
   If LastOpen(0).Id = gEmpresa.Id And Oper <> O_EDIT Then
      Exit Sub
   End If
   
   ' Lo sacamos de la lista
   
   For i = 0 To nLAST - 1
      If LastOpen(i).Id = gEmpresa.Id Then
         For j = i + 1 To nLAST - 1
            LastOpen(j - 1).Nombre = LastOpen(j).Nombre
            LastOpen(j - 1).Id = LastOpen(j).Id
            LastOpen(j).Id = -1
         Next j
         Exit For
         
      End If
   Next i

   ' dejamos libre el primero
   For j = nLAST - 1 To 1 Step -1
      LastOpen(j).Nombre = LastOpen(j - 1).Nombre
      LastOpen(j).Id = LastOpen(j - 1).Id
      LastOpen(j - 1).Id = -1
      
   Next j
   
   ' lo ponemos primero
   LastOpen(0).Nombre = gEmpresa.RazonSocial
   LastOpen(0).Id = gEmpresa.Id
   
   j = 0
   For i = 0 To nLAST - 1
   
'      If i >= gMaxEmpr Then
'         Exit For
'      End If
      
      'If ME_LastOpen(i) = 0 Then
      
      If LastOpen(i).Id > 0 Then
         ME_LastOpen(i).Caption = "&0"
         ME_LastOpen(i).Enabled = False
         ME_LastOpen(i).Visible = (i = 0)
      End If
   
      If LastOpen(i).Id > 0 Then
'         ME_LastOpen(i).Caption = "&" & i + 1 & "  " & FCase(LastOpen(i).Nombre)
         ME_LastOpen(i).Caption = "&" & i + 1 & "  " & LastOpen(i).Nombre
         ME_LastOpen(i).Tag = LastOpen(i).Id
         ME_LastOpen(i).Enabled = True
         ME_LastOpen(i).Visible = True
         
         Call SetIniString(gIniFile, "LastOpen", str(j), LastOpen(i).Id & "|" & LastOpen(i).Nombre)
         j = j + 1
      Else
         Call SetIniString(gIniFile, "LastOpen", str(i), "")
         
      End If
     
   Next i
  
End Sub

Private Sub Tmr_ChkActive_Timer()
   Static n As Integer

   n = n + 1
   
   If n > 2 Then ' 10 mar 2021: para darle tiempo para pedir el código
      Tmr_ChkActive.Enabled = False
   
      If ExitDemoFact() Then
         Unload Me
      End If
   End If

End Sub
' 14 mar 2021: se agrega esta función
Private Sub Tmr_ChkDTE_Timer()
   Dim Dt As Long, Hoy As Long, Msg As String
   
   Tmr_ChkDTE.Enabled = (lnChkDTE < 7)
   lnChkDTE = lnChkDTE + 1
   Hoy = Int(Now)
   Dt = Val(GetIniString(gCfgFile, "Import-" & gEmpresa.Rut, "FDteRec")) ' 2 ago 2019: para saber cuando leyo por última vez
   If Hoy - Dt > 5 Then ' 5 días
      If AcpCountArchFCompra(Val(gEmpresa.Rut)) > 0 Then
         If Dt = 0 Then
            Msg = "Hace más de 10"
         Else
            Msg = "Hace " & (Hoy - Dt)
         End If
         
         Msg = "ATENCIÓN" & vbCrLf & Msg & " días que no importa los DTE Recibidos desde el menú Procesos."
         
         If Hoy - Dt > 15 Then ' 15 días
            Msg = Msg & vbCrLf & "Pasado un tiempo razonable estos datos se van eliminando y ya no pueden ser leídos."
         End If
         
         'MsgBox1 Msg, vbExclamation
      End If
   Else
      Tmr_ChkDTE.Enabled = False
   End If

End Sub
