VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmDTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Factura"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15030
   Icon            =   "FrmDTE.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   15030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_PieDTE 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4035
      Left            =   60
      TabIndex        =   52
      Top             =   6180
      Width           =   13635
      Begin VB.CheckBox Ch_DecEnPrecio 
         Caption         =   "Permitir decimales en el precio"
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   480
         Width           =   2835
      End
      Begin VB.Frame Fr_Totales 
         Height          =   2475
         Left            =   8640
         TabIndex        =   63
         Top             =   120
         Width           =   4815
         Begin VB.TextBox Tx_SubTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Tx_MontoDescto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Tx_PjeDescto 
            Height          =   315
            Left            =   1620
            MaxLength       =   5
            TabIndex        =   27
            Top             =   600
            Width           =   555
         End
         Begin VB.TextBox Tx_Neto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Tx_IVA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Tx_ImpAdic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox Tx_Total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Sub Total"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   78
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Descuento Global"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   77
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   76
            Top             =   660
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Index           =   3
            Left            =   2820
            TabIndex        =   75
            Top             =   660
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Monto Neto"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   74
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label Lb_IVA 
            AutoSize        =   -1  'True
            Caption         =   "IVA"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   1380
            Width           =   255
         End
         Begin VB.Label Lb_TasaIVA 
            AutoSize        =   -1  'True
            Caption         =   "19 %"
            Height          =   195
            Left            =   1920
            TabIndex        =   72
            Top             =   1380
            Width           =   345
         End
         Begin VB.Label Lb_ImpAdic 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto Adicional"
            Height          =   195
            Left            =   180
            TabIndex        =   71
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   70
            Top             =   2100
            Width           =   360
         End
      End
      Begin VB.CheckBox Ch_Referencias 
         Caption         =   "Ver/Agregar Referencias"
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Top             =   1140
         Width           =   2235
      End
      Begin VB.Frame Fr_Botones 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   8640
         TabIndex        =   62
         Top             =   2520
         Width           =   4875
         Begin VB.CommandButton Bt_Firmar 
            Caption         =   "Firmar y Enviar"
            Height          =   315
            Left            =   1740
            TabIndex        =   29
            Top             =   300
            Width           =   1515
         End
         Begin VB.CommandButton Bt_Limpiar 
            Caption         =   "Limpiar"
            Height          =   315
            Left            =   3360
            TabIndex        =   30
            Top             =   300
            Width           =   1395
         End
         Begin VB.CommandButton Bt_Visualizar 
            Caption         =   "Validar y Visualizar"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame Fr_VerColumnas 
         Caption         =   "Ver Columnas"
         Height          =   1035
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   4395
         Begin VB.CheckBox Ch_VerColDTE 
            Caption         =   "Impuestos Adicionales"
            Height          =   195
            Index           =   4
            Left            =   2340
            TabIndex        =   20
            Top             =   660
            Width           =   1875
         End
         Begin VB.CheckBox Ch_VerColDTE 
            Caption         =   "Descripción"
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   19
            Top             =   660
            Width           =   1875
         End
         Begin VB.CheckBox Ch_VerColDTE 
            Caption         =   "Código Producto"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   17
            Top             =   360
            Width           =   1875
         End
         Begin VB.CheckBox Ch_VerColDTE 
            Caption         =   "Unidad de Medida"
            Height          =   195
            Index           =   3
            Left            =   2340
            TabIndex        =   18
            Top             =   360
            Width           =   1875
         End
      End
      Begin VB.CheckBox Ch_DesactivarSelProd 
         Caption         =   "Desactivar Selección de Producto"
         Height          =   375
         Left            =   4560
         TabIndex        =   21
         Top             =   180
         Width           =   3375
      End
      Begin VB.Frame Fr_ObsDTE 
         BorderStyle     =   0  'None
         Caption         =   "Observaciones "
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   2580
         Width           =   7935
         Begin VB.TextBox Tx_ObsDTE 
            Height          =   315
            Left            =   0
            MaxLength       =   100
            TabIndex        =   55
            Top             =   240
            Width           =   7875
         End
         Begin VB.Label Label2 
            Caption         =   "Observaciones:"
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton Bt_DelRef 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         Picture         =   "FrmDTE.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Eliminar referencia seleccionada"
         Top             =   1080
         Width           =   315
      End
      Begin FlexEdGrid3.FEd3Grid Gr_Ref 
         Height          =   1035
         Left            =   0
         TabIndex        =   25
         Top             =   1440
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   1826
         Cols            =   2
         Rows            =   4
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
      Begin VB.Frame Fr_Despacho 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   60
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   8055
         Begin VB.ComboBox Cb_Traslado 
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   60
            Width           =   2535
         End
         Begin VB.ComboBox Cb_TipoDespacho 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   60
            Width           =   3135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Traslado:"
            Height          =   195
            Index           =   6
            Left            =   4620
            TabIndex        =   59
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Despacho:"
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   60
            Top             =   120
            Width           =   1140
         End
      End
      Begin VB.CheckBox Ch_ModPrecio 
         Caption         =   "Permitir modificar el precio"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   780
         Width           =   2835
      End
   End
   Begin VB.ListBox Ls_CopiarDTE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   10920
      TabIndex        =   51
      Top             =   3180
      Visible         =   0   'False
      Width           =   3675
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   2595
      Left            =   60
      TabIndex        =   16
      Top             =   3480
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   4577
      Cols            =   2
      Rows            =   2
      FixedCols       =   1
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   1
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   47
      Top             =   60
      Width           =   14775
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
         Left            =   1740
         Picture         =   "FrmDTE.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Convertir moneda"
         Top             =   180
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
         Left            =   2160
         Picture         =   "FrmDTE.frx":07A6
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Calculadora"
         Top             =   180
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
         Left            =   2580
         Picture         =   "FrmDTE.frx":0B07
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Del 
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
         Picture         =   "FrmDTE.frx":0F30
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Eliminar detalle de producto seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_SelProd 
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
         Left            =   720
         Picture         =   "FrmDTE.frx":132C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Seleccionar Producto"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_SelEnt 
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
         Index           =   0
         Left            =   120
         Picture         =   "FrmDTE.frx":17E7
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Seleccionar Entidad"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   13080
         TabIndex        =   37
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame Fr_Receptor 
      Caption         =   "Receptor"
      Height          =   2775
      Left            =   60
      TabIndex        =   38
      Top             =   720
      Width           =   14775
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   6120
         MaxLength       =   12
         TabIndex        =   84
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ComboBox Cb_Vendedor 
         Height          =   315
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   2280
         Width           =   4575
      End
      Begin VB.ComboBox Cb_DetFormaPago 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton Bt_DatosAdicionales 
         Caption         =   "Datos Factura  Exportación..."
         Height          =   495
         Left            =   13200
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Ch_Rut 
         Height          =   195
         Left            =   780
         TabIndex        =   79
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton Bt_CopiarDTE 
         Height          =   495
         Left            =   13080
         Picture         =   "FrmDTE.frx":1C85
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1980
         Width           =   1455
      End
      Begin VB.ComboBox Cb_FormaDePago 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Tx_FechaVenc 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Bt_FechaVenc 
         Height          =   315
         Left            =   1500
         Picture         =   "FrmDTE.frx":2459
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Tx_MailReceptor 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         Top             =   1620
         Width           =   3345
      End
      Begin VB.CommandButton Bt_SelEnt 
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
         Index           =   1
         Left            =   9120
         Picture         =   "FrmDTE.frx":24CE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Seleccionar Entidad"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Tx_Ciudad 
         Height          =   315
         Left            =   9660
         TabIndex        =   7
         Top             =   1080
         Width           =   3075
      End
      Begin VB.TextBox Tx_Contacto 
         Height          =   315
         Left            =   9660
         TabIndex        =   10
         Top             =   1620
         Width           =   3045
      End
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Left            =   12480
         Picture         =   "FrmDTE.frx":296C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   11220
         TabIndex        =   3
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox Tx_Giro 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1620
         Width           =   5685
      End
      Begin VB.TextBox Tx_Comuna 
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Tx_Direccion 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   5685
      End
      Begin VB.TextBox Tx_RazonSocial 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   540
         Width           =   7065
      End
      Begin VB.TextBox Tx_RUT 
         Height          =   315
         Left            =   240
         MaxLength       =   12
         TabIndex        =   0
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Vendedor:"
         Height          =   195
         Index           =   13
         Left            =   6120
         TabIndex        =   85
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Index           =   12
         Left            =   8160
         TabIndex        =   83
         Top             =   1995
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Detalle Forma de Pago:"
         Height          =   195
         Index           =   11
         Left            =   3720
         TabIndex        =   81
         Top             =   1995
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago:"
         Height          =   195
         Index           =   10
         Left            =   1920
         TabIndex        =   50
         Top             =   1995
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vencimiento:"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   49
         Top             =   2000
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mail Receptor:"
         Height          =   195
         Index           =   8
         Left            =   6120
         TabIndex        =   48
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contacto:"
         Height          =   195
         Index           =   7
         Left            =   9660
         TabIndex        =   46
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   1
         Left            =   11220
         TabIndex        =   45
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Giro:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   44
         Top             =   1380
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   5
         Left            =   9660
         TabIndex        =   43
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comuna:"
         Height          =   195
         Index           =   4
         Left            =   6180
         TabIndex        =   42
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social: "
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   41
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   39
         Top             =   300
         Width           =   390
      End
   End
End
Attribute VB_Name = "FrmDTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'columnas detalle documento
Const C_NUMLIN = 0
Const C_IDDETFACT = 1
Const C_TIPOCOD = 2
Const C_CODPROD = 3
Const C_PRODUCTO = 4
Const C_DESCRIP = 5
Const C_CANTIDAD = 6
Const C_UMEDIDA = 7
Const C_PRECIO = 8
Const C_ESEXENTO = 9
Const C_IDIMPADIC = 10
Const C_IMPADIC = 11
Const C_TASAIMPADIC = 12
Const C_TASAEDITABLE = 13
Const C_MONTOIMPADIC = 14
Const C_CODIMPADICSII = 15
Const C_PJEDESCTO = 16
Const C_MONTODESCTO = 17
Const C_SUBTOTAL = 18
Const C_IDPROD = 19
Const C_PRECIOORI = 20
Const C_UPDATE = 21
Const C_DESCIVA = 22


Const NCOLS = C_DESCIVA

'Columnas Detalle referencias

Const CR_NUMLIN = 0
Const CR_IDREFERENCIA = 1
Const CR_IDDTEREF = 2
Const CR_IDTIPODOCREF = 3
Const CR_TIPODOCREF = 4
Const CR_CODDOCREFSII = 5
Const CR_FOLIO = 6
Const CR_FECHA = 7
Const CR_LNGFECHA = 8
Const CR_CODREFSII = 9
Const CR_REFSII = 10
Const CR_RAZONREF = 11
Const CR_UPDATE = 12

Const R_NCOLS = CR_UPDATE

'columnas matriz de ClsCombo de impuestos adicionales
Const IA_CODSIIDTE = 2
Const IA_TASA = 3

'Columnas matriz de ClsCombo de codigos referencias
Const REF_CODREFSII = 2

'Copiar documento
Const COPY_ULTDTE = 1      'Ingresar documento similar al último emitido
Const COPY_DTEPREVIO = 2   'Ingresar documento similar a uno emitido previamente

Dim lInLoad As Boolean

Dim lFrmWidth As Integer


Dim lIdxTipoDoc As Integer
Dim lTipoDoc As Integer
Dim lCodDocSII As String
Dim lDiminutivoDoc As String
Dim lSoloExento As Boolean
Dim lTieneExento As Boolean
Dim lEsExport As Boolean

Dim lTotAfecto As Double
Dim lTotExento As Double
Dim lTotDescIva As Double
Dim rebajaIva As Boolean


Dim lIdEntidad As Long
Dim lEntCompleta As Boolean
Dim lFldLen(NCOLS) As Integer
Dim lFldLenRef(R_NCOLS) As Byte
Dim lNotSelProd As Boolean

Dim lcbImpAdic As ClsCombo
Dim lcbDocRef As ClsCombo
Dim lIdDTE As Long

Dim lEsGuiaDespacho As Boolean

Dim lDTE As DTE_t

Dim lFmtPrecio As String
Dim lFmtPrecio2 As String

'datos para el caso que es una nota de crédito o débito, en que se recibe el DTE de referencia
Dim lEsNotaCredDeb As Boolean
Dim lTipoRef As Integer          'tipo de referencia: anula, o corrige, válido sólo si lEsNotaCredDeb = true
Dim lIdDTERef As Long            'válido sólo si lEsNotaCredDeb = true
Dim lEsNotaCredDebFactCompra As Boolean     'válido sólo si es nota de crédito o débito.



'Caracteres que hay que escapar:
'  "   &quot;
'  '   &apos;
'  <   &lt;
'  >   &gt;
'  &   &amp;

Public Function FNew(ByVal IdxTipoDoc As Integer, Optional ByVal TipoRef As Integer = 0, Optional ByVal IdDTERef As Long = 0, Optional ByVal EsGuiaDespacho As Boolean = False, Optional ByVal DiminutivoDocRef As String = "", Optional ByVal EsNotaCredDebFactCompra As Boolean = False)

   lIdxTipoDoc = IdxTipoDoc    'si es guía de despacho, este idx corresponde a FAV, porque se asimila la guía de despacho a la factura de venta
   lEsGuiaDespacho = EsGuiaDespacho
   
   If DiminutivoDocRef <> "" Then
      lEsNotaCredDebFactCompra = IIf(DiminutivoDocRef = "FCV", True, False)    'el documento de referencia es una factura de compra del libro de ventas
   Else
      lEsNotaCredDebFactCompra = EsNotaCredDebFactCompra
   End If
   
   lTipoRef = TipoRef
   lIdDTERef = IdDTERef
   Me.Show vbModal
   
End Function

Private Sub Bt_Calc_Click()
   Call Calculadora
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
   
End Sub


Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Call Frm.FSelect(Valor)
      
   If Valor > 0 Then
      Grid.TextMatrix(Grid.Row, C_PRECIO) = Format(Valor, lFmtPrecio)
   End If
   
   Set Frm = Nothing
   'Call CalcTotal
   
End Sub

Private Sub Bt_CopiarDTE_Click()

   Ls_CopiarDTE.Visible = Not Ls_CopiarDTE.Visible
   
   If Ls_CopiarDTE.Visible Then
      Ls_CopiarDTE.SetFocus
   End If
   
End Sub

Private Sub DatosFactExp()
   Dim Frm As FrmDatosFactExp
   
   If Not lEsExport Then
      Exit Sub
   End If
   
   Set Frm = New FrmDatosFactExp
   Call Frm.FEdit(lDTE.FactExp)
   Set Frm = Nothing
   
End Sub
Private Sub DatosAdicionales()
   Dim Frm As FrmDatosAdicDTE
   
   Set Frm = New FrmDatosAdicDTE
   Call Frm.FEdit(lDTE.GuiaDesp)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_DatosAdicionales_Click()

   If lIdEntidad = 0 Or Not lEntCompleta Then
      MsgBox1 "Debe seleccionar la entidad receptora y completar todos los datos antes de continuar.", vbExclamation
      Exit Sub
   End If

   If lEsExport Then
      Call DatosFactExp
   
   Else
      Call DatosAdicionales

   End If
   
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Row <> Grid.RowSel Then
      MsgBox1 "Debe eliminar un registro a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_CODPROD) = "" And Grid.TextMatrix(Row, C_PRODUCTO) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If MsgBox1("¿Está seguro que desea eliminar este registro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_IDDETFACT, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
   If lTotDescIva > 0 Then
   
   Else
    Call CalcTotal
   End If
End Sub


Private Sub Bt_DelRef_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Gr_Ref.Row
   
   If Row < Gr_Ref.FixedRows Then
      Exit Sub
   End If
   
   If Gr_Ref.Row <> Gr_Ref.RowSel Then
      MsgBox1 "Debe eliminar un registro a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Gr_Ref.TextMatrix(Row, CR_TIPODOCREF) = "" And Gr_Ref.TextMatrix(Row, CR_FOLIO) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If MsgBox1("¿Está seguro que desea eliminar esta referencia?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Gr_Ref, Row, FGR_D, CR_IDREFERENCIA, C_UPDATE)
      
   Gr_Ref.rows = Gr_Ref.rows + 1
      
End Sub

Private Sub Bt_Firmar_Click()
   Dim XmlDTE As String, Folio As Long, TipoDTE As Integer, TrackID As String, Fn As String, FnCed As String
   Dim Rc As Long, bFirmar As Boolean
   Dim Q1 As String
   Dim Error As String
   Dim UrlDTE As String
   
   If Not Valida() Then
      Exit Sub
   End If
   
   bFirmar = Bt_Firmar.Enabled
   Me.MousePointer = vbHourglass
   
   Call AddDebug("FrmDTE_Firmar: Antes de SaveAll. idDTE=" & lDTE.IdDTE)
   
   If lDTE.IdDTE <= 0 Then ' 31 dic 2020: para que no genere otro id
      Call SaveAll ' descomentar para produccion
      
      Grid.Locked = True ' no se puede modificar nada, por si ya fue emitido pero no tenemos el folio
      Gr_Ref.Locked = True
      Fr_Receptor.Enabled = False
      Fr_Totales.Enabled = False
      Tx_ObsDTE.Locked = True
   End If
   
  If W.InDesign Then
      'gEmpresa.Rut = RUT_EMP_ACEPTA  ' -8  RUT de Acepta para Pruebas
      gEmpresa.Rut = RUT_EMP_TECNOBACK
      gEmpresa.RutFirma = "9199548-4"
      gConectData.RutFirma = "9199548"
   End If
   
   If gConectData.RutFirma = "" Then
     MsgBox1 "Falta ingresar el Rut Firmante Menu (Empresa -> Configurar conexión Facturación Electrónica -> Firma)", vbInformation
     Exit Sub
   Else
        gEmpresa.RutFirma = gConectData.RutFirma & "-" & DV_Rut(gConectData.RutFirma)
   End If

   Call AddDebug("FrmDTE_Firmar: Antes de GenXMLDTE")
   XmlDTE = GenXMLDTE(lDTE)
   'XmlDTE = GenJsonDTE(lDTE)
   
   Call AddDebug("FrmDTE_Firmar: Antes de LpProcesar: Proveedor: " & gConectData.Proveedor & ", idDTE=" & lDTE.IdDTE & ", RUT=" & Tx_RUT & ", Fecha=" & Tx_Fecha)
  
   If W.InDesign Then
      Debug.Print vbCrLf & "* * * En Diseño no se procesa el documento. * * *"
      Me.MousePointer = vbDefault
      'Exit Sub
   End If
  
   Bt_Firmar.Enabled = False
   DoEvents
  
   If gConectData.Proveedor = PROV_ACEPTA Then
      
      Rc = AcpProcesar(lDTE.IdDTE, XmlDTE, lDTE.MailReceptor, Trim(Tx_ObsDTE), Folio, UrlDTE) ' 31 dic 2020: se agrega idDTE
      
      If Rc = 0 And Folio > 0 Then  ' 9 feb 2021: se agrega And Folio > 0
         lDTE.Folio = Folio
         lDTE.idEstado = EDTE_EMITIDO
         TrackID = ""
         lDTE.TrackID = TrackID
         
         Q1 = "UPDATE DTE SET Folio = " & Folio & ", TrackID = '" & TrackID & "', IdEstado = " & lDTE.idEstado & ", UrlDTE = '" & ParaSQL(UrlDTE) & "'"
         Q1 = Q1 & " WHERE IdDTE=" & lIdDTE & " AND IdEmpresa = " & gEmpresa.Id
         Call ExecSQL(DbMain, Q1)
   
         bFirmar = False
         
         MsgBox1 "Documento firmado y enviado con Folio " & Folio, vbInformation
         
         Call AcpShowDTE(Me, UrlDTE, True)
         
         MsgBox1 "Para obtener el documento DEFINITIVO, vaya a los DTE Emitidos, a través del menú Reportes o el botón con el mismo nombre.", vbInformation
      
      ElseIf Rc = AERR_MOTOR Then   ' 31 dic 2020
         Call AddLog("FrmDTE_Firmar: MOTOR, Error=" & Rc & ", idDTE=" & lDTE.IdDTE & ", RUT=" & Tx_RUT & ", Fecha=" & Tx_Fecha)
         MsgBox "Error " & Rc & " en la generación del documento, no cierre la ventana, no modifique nada y vuelva a intentar en un momento.", vbExclamation
      
      Else
         Call AddLog("FrmDTE_Firmar: Error=" & Rc & ", idDTE=" & lDTE.IdDTE & ", RUT=" & Tx_RUT & ", Fecha=" & Tx_Fecha)
         MsgBox "Error " & Rc & " en la generación del documento.", vbExclamation
      
      End If
      
   End If
   
   Me.MousePointer = vbDefault
   
   Bt_Firmar.Enabled = bFirmar
   
End Sub

'Private Function GetPdfDTE(ByVal Folio As Long, ByVal TipoDTE As Integer, ByVal FechaDTE As Long) As Integer
'   Dim Rc As Integer
'   Dim Fn As String, FnCed As String
'
'   Rc = LpObtenerLink(LP_TC_NORMAL, Folio, TipoDTE, FechaDTE, Fn, False)
'
'   If Rc = 0 Then
'      Call LpObtenerLink(LP_TC_CEDIBLE, Folio, TipoDTE, lDTE.Fecha, FnCed, False)
'
'      Call AbrirPDF(Fn)
'
'   ElseIf Rc = LP_ERR_PDFNOTAVAILABLE Then
'      MsgBox1 "No es posible obtener el archivo PFD del documento. Por favor inténtelo más tarde.", vbInformation
''      Tmr_GetPdfDTE.Interval = 10000
''      Tmr_GetPdfDTE.Enabled = True
'   End If
'
'End Function

Private Sub Bt_Limpiar_Click()
   Call LimpiarForm
   Bt_Firmar.Enabled = True
   
   Grid.Locked = False ' 31 dic 2020: se vuelven a dejar libres
   Gr_Ref.Locked = False
   Fr_Receptor.Enabled = True
   Fr_Totales.Enabled = True
   Tx_ObsDTE.Locked = False

   
End Sub

Private Sub Bt_SelFecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha)
   Set Frm = Nothing
   
   Call SetTxDate(Tx_FechaVenc, DateAdd("d", 30, GetTxDate(Tx_Fecha)))

End Sub

Private Sub Bt_FechaVenc_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaVenc)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_SelProd_Click()
   Dim Frm As FrmProductos
   Dim IdProd As Long
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Precio As Double
   
   Set Frm = New FrmProductos
   IdProd = Frm.FSelect
   Set Frm = Nothing
   
   If IdProd > 0 Then
      Row = Grid.Row
      Q1 = "SELECT TipoCod, CodProd, Producto, Obs, UMedida, Precio FROM Productos WHERE IdProducto = " & IdProd & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         Grid.TextMatrix(Row, C_IDPROD) = IdProd
         Grid.TextMatrix(Row, C_TIPOCOD) = vFld(Rs("TipoCod"))
         Grid.TextMatrix(Row, C_CODPROD) = vFld(Rs("CodProd"))
         
         If lEsExport Then
            If lDTE.FactExp.TipoCambio = 0 Then
               MsgBox1 "Falta ingresar el tipo de cambio en los 'Datos Factura Exportación', con el fin de calcular el valor del producto en la moneda que corresponde a este documento.", vbExclamation
            Else
               Precio = vFld(Rs("Precio")) / lDTE.FactExp.TipoCambio
            End If
         Else
            Precio = vFld(Rs("Precio"))
         End If
            
         Grid.TextMatrix(Row, C_PRODUCTO) = vFld(Rs("Producto"))
         Grid.TextMatrix(Row, C_DESCRIP) = vFld(Rs("Obs"))
         Grid.TextMatrix(Row, C_UMEDIDA) = vFld(Rs("UMedida"))
         Grid.TextMatrix(Row, C_PRECIO) = Format(Precio, IIf(Ch_DecEnPrecio <> 0, lFmtPrecio2, lFmtPrecio))
         Grid.TextMatrix(Row, C_PRECIOORI) = vFld(Rs("Precio"))
                  
         If Row >= Grid.rows - 2 Then
            Grid.rows = Grid.rows + 1
         End If
         
      End If
      
      Call CloseRs(Rs)
      
   End If
   
End Sub
Private Sub Bt_Visualizar_Click()
   Dim XmlDTE As String, Fn As String
   Dim PreHtmlDTE As String
   Dim FName As String
   Dim Fd As Long
   Dim Rc As Long
      
   
   If Not Valida() Then
      Exit Sub
   End If
   
'   If gConectData.Proveedor = PROV_ACEPTA Then   'por ahora no tenemos visualizador
'      MsgBox1 "Información del DTE validada.", vbInformation
'      Exit Sub
'   End If

   If W.InDesign Then
      'gEmpresa.Rut = RUT_EMP_ACEPTA  ' -8  RUT de Acepta para Pruebas
      gEmpresa.Rut = RUT_EMP_TECNOBACK
      gEmpresa.RutFirma = "9199548-4"
      gConectData.RutFirma = "9199548"
   End If
   
   If gConectData.RutFirma = "" Then
     MsgBox1 "Falta ingresar el Rut Firmante Menu (Empresa -> Configurar conexión Facturación Electrónica -> Firma)", vbInformation
     Exit Sub
   Else
        gEmpresa.RutFirma = gConectData.RutFirma & "-" & DV_Rut(gConectData.RutFirma)
   End If
   
   Me.MousePointer = vbHourglass
   
   Call AddDebug("FrmDTE_Visualizar: Antes de GenDTEStruct")
   
   Call GenDTEStruct
   
   Call AddDebug("FrmDTE_Visualizar: Antes de GenXMLDTE")
   XmlDTE = GenXMLDTE(lDTE)
   
   Call AddDebug("FrmDTE_Visualizar: Antes de LpPrevisualizaPDF")
   
   If gConectData.Proveedor = PROV_ACEPTA Then
      Rc = AcpPrevisualizar(XmlDTE, lDTE.MailReceptor, Trim(Tx_ObsDTE), PreHtmlDTE, lDTE.EsExport)
      
      If Rc <> 0 Then    'error
         Me.MousePointer = vbDefault
         Exit Sub
      End If
      
      If Len(PreHtmlDTE) < 100 Then   ' 21 sep 2020 - pam: si es muy pequeño, no hay nada que mostrar
         Me.MousePointer = vbDefault
         Exit Sub
      End If
      
'      fname = W.TmpDir & "\PrevDTE_" & Format(Now, "ddmmyyyy_hhmmss") & ".htm"
'
'      Fd = FreeFile
'      Open fname For Output As #Fd
'      If err Then
'         MsgErr fname
'         Me.MousePointer = vbDefault
'         Exit Sub
'      End If
'
'      Print #Fd, PreHtmlDTE
'
'      Close #Fd
'
''      Rc = Shell(gHtmExt.OpenCmd & " " & FName, vbNormalFocus)
'      Rc = ShellExecute(Me.hWnd, "open", fname, fname, "", SW_SHOWNORMAL)
      Call AcpShowDTE(Me, Trim(Replace(PreHtmlDTE, Chr(34), "")), True)
'      If Rc < 32 Then
'         MsgBox1 "Error " & Rc & ", " & fname, vbExclamation
'      Else
'         Sleep (5000) ' milisegundos
'      End If
'
'      If W.InDesign = False Then
'         Call RemoveFile(fname)
'      End If
   
'   ElseIf LpPrevisualizaPDF(XmlDTE, Fn) = 0 Then
'
'      Call AbrirPDF(Fn)
                 
   End If
   
   
   Me.MousePointer = vbDefault
   
End Sub
Private Sub Cb_FormaDePago_Click()
   
   If CbItemData(Cb_FormaDePago) = FP_CONTADO Then
      Call SetTxDate(Tx_FechaVenc, GetTxDate(Tx_Fecha))
   ElseIf CbItemData(Cb_FormaDePago) = FP_CREDITO Then
      Call SetTxDate(Tx_FechaVenc, DateAdd("m", 1, GetTxDate(Tx_Fecha)))
   End If
   Call CargaCbDetFormaPago
      
End Sub

Private Sub Ch_DesactivarSelProd_Click()

   If Ch_DesactivarSelProd <> 0 Then
      lNotSelProd = True
   Else
      lNotSelProd = False
   End If
   
   gEmpConfig.OptEdFact(OPTEDFACT_NOTSELPROD) = Ch_DesactivarSelProd

End Sub

Private Sub Ch_ModPrecio_Click()

   gEmpConfig.OptEdFact(OPTEDFACT_MODPRECIO) = Ch_ModPrecio

End Sub

Private Sub Ch_Referencias_Click()

   If Ch_Referencias <> 0 Then
      Gr_Ref.Visible = True
   Else
      Gr_Ref.Visible = False
   End If
   
   gEmpConfig.OptEdFact(OPTEDFACT_VERREF) = Ch_Referencias

End Sub


Private Sub Ch_VerColDTE_Click(Index As Integer)
      
   Select Case Index
   
      Case OPTEDFACT_VERCOLCODPROD
         If Ch_VerColDTE(Index) <> 0 Then
            Grid.ColWidth(C_TIPOCOD) = 800
            Grid.ColWidth(C_CODPROD) = 1200
            Grid.TextMatrix(0, C_TIPOCOD) = "Tipo Cód."
            Grid.TextMatrix(0, C_CODPROD) = "Código"
         Else
            Grid.ColWidth(C_TIPOCOD) = 0
            Grid.ColWidth(C_CODPROD) = 0
            Grid.TextMatrix(0, C_TIPOCOD) = ""
            Grid.TextMatrix(0, C_CODPROD) = ""
         End If
      
      Case OPTEDFACT_VERCOLDESC
         If Ch_VerColDTE(Index) <> 0 Then
            Grid.ColWidth(C_DESCRIP) = 4000
            Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
         Else
            Grid.ColWidth(C_DESCRIP) = 0
            Grid.TextMatrix(0, C_DESCRIP) = ""
         End If
      
      Case OPTEDFACT_VERCOLUMED
         If Ch_VerColDTE(Index) <> 0 Then
            Grid.ColWidth(C_UMEDIDA) = 1000
            Grid.TextMatrix(0, C_UMEDIDA) = "U. Medida"
         Else
            Grid.ColWidth(C_UMEDIDA) = 0
            Grid.TextMatrix(0, C_UMEDIDA) = ""
         End If
         
      Case OPTEDFACT_VERCOLIMPADIC
         If Ch_VerColDTE(Index) <> 0 Then
            Grid.ColWidth(C_IMPADIC) = 3000
            Grid.TextMatrix(0, C_IMPADIC) = "Imp. Adicional"
            Grid.ColWidth(C_MONTOIMPADIC) = 1200
            Grid.TextMatrix(0, C_MONTOIMPADIC) = "Subtot.Imp.Adic"
         
            'factura de compra hay que mostrar la tasa
            If lDiminutivoDoc = "FCV" Or lEsNotaCredDebFactCompra Then
               Grid.ColWidth(C_TASAIMPADIC) = 660
               Grid.TextMatrix(0, C_TASAIMPADIC) = "Tasa"
            End If
         Else
            Grid.ColWidth(C_IMPADIC) = 0
            Grid.TextMatrix(0, C_IMPADIC) = ""
            Grid.ColWidth(C_MONTOIMPADIC) = 0
            Grid.TextMatrix(0, C_MONTOIMPADIC) = ""
            Grid.ColWidth(C_TASAIMPADIC) = 0
            Grid.TextMatrix(0, C_TASAIMPADIC) = ""
         End If
         
   End Select

   gEmpConfig.OptEdFact(Index) = Ch_VerColDTE(Index)
   
   Call LocateFrTotales
End Sub


Private Sub Form_Activate()
   
   If lIdDTERef > 0 Then
      Call FillDTERef
   End If
   
   If lEsExport Then
      Tx_Ciudad.SetFocus
   End If

End Sub


Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim CbRef As ComboBox
   Dim TipoDocFCC As Integer
   
   lInLoad = True
   
   Call SetTxDate(Tx_Fecha, Now)
   Call SetTxDate(Tx_FechaVenc, DateAdd("d", 30, Now))

   lTipoDoc = gTipoDoc(lIdxTipoDoc).TipoDoc
   lCodDocSII = gTipoDoc(lIdxTipoDoc).CodDocDTESII
   lDiminutivoDoc = gTipoDoc(lIdxTipoDoc).Diminutivo
   lSoloExento = IIf(gTipoDoc(lIdxTipoDoc).TieneExento And Not gTipoDoc(lIdxTipoDoc).TieneAfecto, True, False)
   lTieneExento = IIf(gTipoDoc(lIdxTipoDoc).TieneExento, True, False)
   Me.Caption = "Emitir " & gTipoDoc(lIdxTipoDoc).Nombre & " Electrónica"
   
   If lDiminutivoDoc = "EXP" Or lDiminutivoDoc = "NCE" Or lDiminutivoDoc = "NDE" Then
      lEsExport = True
   End If
   
   Ch_Rut = 1
   Ch_Rut.Enabled = False
   
   If lEsGuiaDespacho Then
      Me.Caption = "Emitir Guía de Despacho Electrónica"
      lCodDocSII = CODDOCDTESII_GUIADESPACHO
      Fr_Despacho.Visible = True
'      Fr_ObsDTE.Visible = False
      Tx_ObsDTE = ""
      Bt_DatosAdicionales.Caption = "Datos Adicionales..."
      lFmtPrecio = NUMFMT
      lFmtPrecio2 = DBLFMT2DO 'le pone decimales sólo si tiene
      
   ElseIf Not lEsExport Then
      Tx_ObsDTE = gEmpresa.ObsDTE
      Bt_DatosAdicionales.Visible = False
      lFmtPrecio = NUMFMT
      lFmtPrecio2 = DBLFMT2DO 'le pone decimales sólo si tiene
   
   Else    'lEsExport = True
   
      Call SetRO(Tx_RUT, True)   'no se permite otro RUT que el estándar
'      Ch_Rut.Enabled = True
      lFmtPrecio = DBLFMT2
      Ch_DecEnPrecio.Visible = False
      Bt_DatosAdicionales.Caption = "Datos Factura  Exportación..."
   End If
    
   Call SetupForm
      
   For i = 1 To MAX_OPTEDFACTVERCOL
      If Ch_VerColDTE(i).Enabled Then
         Ch_VerColDTE(i) = gEmpConfig.OptEdFact(i)
         If Ch_VerColDTE(i) = 0 Then    'si vale cero no llama al click en el load (por qué???)
            Call Ch_VerColDTE_Click(i)
         End If
      End If
   Next i

   Ch_Referencias = gEmpConfig.OptEdFact(OPTEDFACT_VERREF)
   Ch_DesactivarSelProd = gEmpConfig.OptEdFact(OPTEDFACT_NOTSELPROD)
   lNotSelProd = gEmpConfig.OptEdFact(OPTEDFACT_NOTSELPROD)
   Ch_ModPrecio = gEmpConfig.OptEdFact(OPTEDFACT_MODPRECIO)
   
   Call GetFLdLen
   
   Set lcbImpAdic = New ClsCombo
   Call lcbImpAdic.SetControl(Grid.CbList(C_IMPADIC))
   
   If lDiminutivoDoc = "FCV" Or lEsNotaCredDebFactCompra Then   'es factura de compra
      'los otros impuestos deben ser los que corresponden a la factura de compra del libro de compras (Joshua Catrin 30 ago 2018)
      TipoDocFCC = FindTipoDoc(LIB_COMPRAS, "FCC")
      Call FillTipoValLib(lcbImpAdic, LIB_COMPRAS, True, True, "", TipoDocFCC, LIBCOMPRAS_IVARETTOT, gOcultarImpAdicDescont)
   Else
      Call FillTipoValLib(lcbImpAdic, LIB_VENTAS, True, True, "", lTipoDoc, LIBVENTAS_REBAJA65, gOcultarImpAdicDescont)
   End If
   
   If lDiminutivoDoc = "NCV" Or lDiminutivoDoc = "NDV" Then
      Bt_CopiarDTE.Visible = False
   End If
   
   Set lcbDocRef = New ClsCombo
   Call lcbDocRef.SetControl(Gr_Ref.CbList(CR_TIPODOCREF))
   Call lcbDocRef.AddItem(" ", 0)
   Q1 = "SELECT Nombre, IdTipoDocRef, CodDocRefSII FROM TipoDocRef ORDER BY Nombre"
   Call lcbDocRef.FillCombo(DbMain, Q1, -1)
   
   Set CbRef = Gr_Ref.CbList(CR_REFSII)
   Call CbAddItem(CbRef, " ", 0, True)
   For i = 1 To MAX_TIPOREF
      Call CbAddItem(CbRef, gTipoRefSII(i), i)
   Next i
   
   Call CbAddItem(Cb_FormaDePago, "", 0)
   For i = 1 To UBound(gFormaDePago)
      Call CbAddItem(Cb_FormaDePago, gFormaDePago(i), i)
   Next i
   Cb_FormaDePago.ListIndex = 0
   
   Call CargaCbVendedor
   
   Call CbAddItem(Ls_CopiarDTE, "Ingresar DTE similar al último Enviado/Emitido ", COPY_ULTDTE)
   Call CbAddItem(Ls_CopiarDTE, "Ingresar DTE basado en uno emitido previamente ", COPY_DTEPREVIO)
   Ls_CopiarDTE.ListIndex = 0
   
'   If gConectData.Proveedor = PROV_ACEPTA Then
'      Bt_Visualizar.Caption = "Validar"
'   End If
      
   For i = 1 To MAX_TIPODESPACHO
      Call CbAddItem(Cb_TipoDespacho, gTipoDespacho(i), i)
   Next i

   For i = 0 To MAX_TIPOTRASLADO
      Call CbAddItem(Cb_Traslado, gTipoTraslado(i), i)
   Next i
   
   If lEsExport Then
      Call SelEntDefExport
   End If
   
   Me.Caption = Me.Caption & " - " & FmtRut(gEmpresa.Rut)
   
   lInLoad = False
   
   lFrmWidth = Me.Width
End Sub

Private Sub CargaCbDetFormaPago()
Dim Q1 As String
Dim Rs As Recordset

   Cb_DetFormaPago.Clear
   Q1 = "SELECT Id, Descripcion "
   Q1 = Q1 & " From DetFormaPago "
   Q1 = Q1 & " Where FormaPago = " & CbItemData(Me.Cb_FormaDePago)
   Q1 = Q1 & " AND Estado = 1 "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Call CbAddItem(Cb_DetFormaPago, vFld(Rs("Descripcion")), vFld(Rs("Id")))
   Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
End Sub

Private Function CargaCbVendedor(Optional ByVal codigo As Long = 0) As Boolean
Dim Q1 As String
Dim Rs As Recordset
   
   CargaCbVendedor = False
   
   Q1 = "SELECT Rut, Nombre "
   Q1 = Q1 & " From Vendedor "
   Q1 = Q1 & " Where Estado = 1"
   If Trim(codigo) <> 0 Then
        Q1 = Q1 & " AND Codigo = " & Trim(codigo)
   Else
        Cb_Vendedor.Clear
   End If
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      If Trim(codigo) <> 0 Then
        Call CbSelItem(Cb_Vendedor, vFld(Rs("Rut")))
        CargaCbVendedor = True
      Else
        Call CbAddItem(Cb_Vendedor, vFld(Rs("Nombre")), vFld(Rs("Rut")))
      End If
   Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
End Function

Private Sub SetupForm()

   If lIdDTERef <> 0 Then
      Call SetTxRO(Tx_RUT, True)
      Bt_SelEnt(0).Enabled = False
      Bt_SelEnt(1).Enabled = False
   ElseIf Not lEsExport Then
      Call SetTxRO(Tx_RUT, False)
   End If
   
   Call SetTxRO(Tx_RazonSocial, True)
   Call SetTxRO(Tx_Direccion, True)
   If Not lEsExport Then
      Call SetTxRO(Tx_Ciudad, True)
   End If
   Call SetTxRO(Tx_Comuna, True)
   Call SetTxRO(Tx_Giro, True)
   Call SetTxRO(Tx_MailReceptor, True)
   
   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid, True)
    
   Grid.ColWidth(C_NUMLIN) = 0
   Grid.ColWidth(C_TIPOCOD) = 900
   Grid.ColWidth(C_CODPROD) = 1200
   Grid.ColWidth(C_PRODUCTO) = 2500
   Grid.ColWidth(C_DESCRIP) = 4000
   Grid.ColWidth(C_CANTIDAD) = 1200
   Grid.ColWidth(C_UMEDIDA) = 1000
   Grid.ColWidth(C_PRECIO) = 1200
   Grid.ColWidth(C_PRECIOORI) = 0
   
   If lTieneExento = True And Not lSoloExento Then
      Grid.ColWidth(C_ESEXENTO) = 300
   Else
      Grid.ColWidth(C_ESEXENTO) = 0
   End If
   
   Grid.ColWidth(C_IDIMPADIC) = 0
   If lSoloExento Then
      Grid.ColWidth(C_IMPADIC) = 0
      Grid.ColWidth(C_MONTOIMPADIC) = 0
      Ch_VerColDTE(OPTEDFACT_VERCOLIMPADIC).Visible = False
      Ch_VerColDTE(OPTEDFACT_VERCOLIMPADIC).Enabled = False
      
      Lb_ImpAdic.Visible = False
      Tx_ImpAdic.Visible = False

   Else
      Grid.ColWidth(C_IMPADIC) = 3000
      Grid.ColWidth(C_MONTOIMPADIC) = 1200
      
   End If
   
   Grid.ColWidth(C_TASAIMPADIC) = 0
   Grid.ColWidth(C_TASAEDITABLE) = 0
   Grid.ColWidth(C_CODIMPADICSII) = 0
   Grid.ColWidth(C_PJEDESCTO) = 700
   Grid.ColWidth(C_MONTODESCTO) = 0
   Grid.ColWidth(C_SUBTOTAL) = 1200
   Grid.ColWidth(C_IDDETFACT) = 0
   Grid.ColWidth(C_IDPROD) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_TIPOCOD) = flexAlignRightCenter
   Grid.ColAlignment(C_CODPROD) = flexAlignRightCenter
   Grid.ColAlignment(C_CANTIDAD) = flexAlignRightCenter
   Grid.ColAlignment(C_PRECIO) = flexAlignRightCenter
   Grid.ColAlignment(C_ESEXENTO) = flexAlignCenterCenter
   Grid.ColAlignment(C_PJEDESCTO) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTODESCTO) = flexAlignRightCenter
   Grid.ColAlignment(C_SUBTOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_TASAIMPADIC) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOIMPADIC) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_TIPOCOD) = "Tipo Cód."
   Grid.TextMatrix(0, C_CODPROD) = "Código"
   Grid.TextMatrix(0, C_PRODUCTO) = "Producto"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(0, C_CANTIDAD) = "Cantidad"
   Grid.TextMatrix(0, C_UMEDIDA) = "U. Medida"
   Grid.TextMatrix(0, C_PRECIO) = "Precio"
   If Grid.ColWidth(C_ESEXENTO) > 0 Then
      Grid.TextMatrix(0, C_ESEXENTO) = "Ex"
   End If
   If Grid.ColWidth(C_IMPADIC) > 0 Then
      Grid.TextMatrix(0, C_IMPADIC) = "Imp. Adicional"
      Grid.TextMatrix(0, C_MONTOIMPADIC) = "Subtot.Imp.Adic"
   End If
   Grid.TextMatrix(0, C_PJEDESCTO) = "% Desc."
   Grid.TextMatrix(0, C_SUBTOTAL) = "Sub Total"
   
   Gr_Ref.Cols = R_NCOLS + 1

   Call FGrSetup(Gr_Ref, True)
    
   Gr_Ref.ColWidth(C_NUMLIN) = 0
   Gr_Ref.ColWidth(CR_IDREFERENCIA) = 0
   Gr_Ref.ColWidth(CR_IDDTEREF) = 0
   Gr_Ref.ColWidth(CR_IDTIPODOCREF) = 0
   Gr_Ref.ColWidth(CR_TIPODOCREF) = 1700
   Gr_Ref.ColWidth(CR_CODDOCREFSII) = 0
   Gr_Ref.ColWidth(CR_FOLIO) = 1100
   Gr_Ref.ColWidth(CR_FECHA) = 900
   Gr_Ref.ColWidth(CR_LNGFECHA) = 0
   Gr_Ref.ColWidth(CR_CODREFSII) = 0
   Gr_Ref.ColWidth(CR_REFSII) = 1100
   Gr_Ref.ColWidth(CR_RAZONREF) = 3000
   Gr_Ref.ColWidth(CR_UPDATE) = 0
   
   Gr_Ref.ColAlignment(CR_FOLIO) = flexAlignRightCenter
   
   Gr_Ref.TextMatrix(0, CR_TIPODOCREF) = "Tipo Doc."
   Gr_Ref.TextMatrix(0, CR_FOLIO) = "Folio Ref."
   Gr_Ref.TextMatrix(0, CR_FECHA) = "Fecha"
   Gr_Ref.TextMatrix(0, CR_REFSII) = "Tipo Ref. SII"
   Gr_Ref.TextMatrix(0, CR_RAZONREF) = "Razón Referencia"

   Call FGrVRows(Grid)
   Call FGrVRows(Gr_Ref)
   
   Lb_TasaIVA = gIVA * 100 & "%"
   If lSoloExento Then
      Lb_TasaIVA = 0
      Lb_IVA.Visible = False
      Lb_TasaIVA.Visible = False
      Tx_IVA.Visible = False
   End If
   
End Sub
Private Sub GetFLdLen()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim fld As Field
   Dim Wh As String
         
   
   Q1 = "SELECT TipoCod, CodProd, Producto, UMedida, Descrip "
   Q1 = Q1 & " FROM DetDTE"
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.Id & " AND IdDTE = 0 "    'no se obtienen registros, sólo para obtener el largo de los campos
   Set Rs = OpenRs(DbMain, Q1)

   If lFldLen(C_TIPOCOD) = 0 Then
      lFldLen(C_TIPOCOD) = FldSize(Rs("TipoCod"))
      lFldLen(C_CODPROD) = FldSize(Rs("CodProd"))
      lFldLen(C_PRODUCTO) = FldSize(Rs("Producto"))
      lFldLen(C_UMEDIDA) = FldSize(Rs("UMedida"))
      lFldLen(C_CANTIDAD) = MAX_DIGITOSCANT
      lFldLen(C_PRECIO) = MAX_DIGITOSVALOR
      lFldLen(C_PJEDESCTO) = 5 ' 99.99
      lFldLen(C_DESCRIP) = MAX_LEN_DESCRIPPROD       ' Rs("Descrip").Size esto no resulta porque es memo
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT FolioRef, RazonReferencia "
   Q1 = Q1 & " FROM Referencias"
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.Id & " AND IdDTE = 0 "    'no se obtienen registros, sólo para obtener el largo de los campos
   Set Rs = OpenRs(DbMain, Q1)

   If lFldLenRef(CR_FOLIO) = 0 Then
      lFldLenRef(CR_FOLIO) = Rs("FolioRef").Size
      lFldLenRef(CR_FECHA) = 10
      lFldLenRef(CR_RAZONREF) = Rs("RazonReferencia").Size
   End If
   
   Call CloseRs(Rs)
   
End Sub
Private Sub LimpiarEntidad()

   lIdEntidad = 0
   lEntCompleta = 0
   
   Tx_RazonSocial = ""
   Tx_Direccion = ""
   Tx_Direccion = ""
   Tx_Comuna = ""
   Tx_Giro = ""
   Tx_Contacto = ""

End Sub


Private Sub LimpiarForm()
   Dim i As Integer, j As Integer

   Call SetTxDate(Tx_Fecha, Now)
   
   lIdEntidad = 0
   lEntCompleta = 0
   
   Tx_RUT = ""
   Tx_RazonSocial = ""
   Tx_Direccion = ""
   Tx_Direccion = ""
   Tx_Comuna = ""
   Tx_Giro = ""
   Tx_Contacto = ""

   If Not lEsExport Then
      Call SetTxRO(Tx_RUT, False)
   End If
   Bt_SelEnt(0).Enabled = True
   Bt_SelEnt(1).Enabled = True

   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   Call FGrVRows(Grid, 1)
   Grid.Redraw = True
   
   Gr_Ref.Redraw = False
   Gr_Ref.rows = Gr_Ref.FixedRows
   Call FGrVRows(Gr_Ref, 1)
   Gr_Ref.Redraw = True
   
   Tx_SubTotal = ""
   Tx_MontoDescto = ""
   Tx_PjeDescto = ""
   Tx_Neto = ""
   Tx_IVA = ""
   Tx_Total = ""
   
   lIdDTE = 0

   'datos para el caso que es una nota de crédito o débito, en que se recibe el DTE de referencia
   lTipoRef = 0
   lIdDTERef = 0

   lDTE.IdDTE = 0
   
   lDTE.IdEmpresa = 0
'   lDTE.Ano = gEmpresa.Ano
   lDTE.TipoDoc = 0
   lDTE.TipoLib = 0
   lDTE.CodDocSII = 0
'   If lEsGuiaDespacho Then
'      lDTE.CodDocSII = 0
'   End If
   lDTE.Folio = 0
   lDTE.idEstado = 0
   lDTE.Fecha = 0
   lDTE.FechaVenc = 0
   lDTE.FormaDePago = 0
   lDTE.IdEntidad = 0
   lDTE.Rut = ""
   lDTE.NotValidRut = 0
   lDTE.RazonSocial = ""
   lDTE.Giro = ""
   lDTE.Direccion = ""
   lDTE.Comuna = ""
   lDTE.Ciudad = ""
   lDTE.Contacto = ""
   lDTE.MailReceptor = ""
   lDTE.SubTotal = 0
   lDTE.PjeDestoGlobal = 0
   lDTE.DesctoGlobal = 0
   lDTE.Neto = 0
   lDTE.Exento = 0
   lDTE.TasaIVA = 0
   lDTE.Iva = 0
   lDTE.ImpAdicional = 0
   lDTE.Total = 0
   lDTE.TipoDespacho = 0
   lDTE.Traslado = 0
   lDTE.EsExport = 0
   lDTE.TrackID = ""
   lDTE.EsGuiaDesp = 0
   

   'limpiamos los arreglos
   For j = 0 To MAX_ITEMDTE
      lDTE.DetDTE(j).IdDetDTE = 0
      lDTE.DetDTE(j).IdDTE = 0
      lDTE.DetDTE(j).IdEmpresa = 0
      lDTE.DetDTE(j).IdProducto = 0
      lDTE.DetDTE(j).TipoCod = ""
      lDTE.DetDTE(j).CodProd = ""
      lDTE.DetDTE(j).Producto = ""
      lDTE.DetDTE(j).Descrip = ""
      lDTE.DetDTE(j).UMedida = ""
      lDTE.DetDTE(j).Cantidad = 0
      lDTE.DetDTE(j).Precio = 0
      lDTE.DetDTE(j).EsExento = 0
      lDTE.DetDTE(j).IdImpAdic = 0
      lDTE.DetDTE(j).CodImpAdicSII = ""
      lDTE.DetDTE(j).TasaImpAdic = 0
      lDTE.DetDTE(j).MontoImpAdic = 0
      lDTE.DetDTE(j).DescImpAdic = ""
      lDTE.DetDTE(j).PjeDescto = 0
      lDTE.DetDTE(j).MontoDescto = 0
      lDTE.DetDTE(j).SubTotal = 0
   Next j
   
   For i = 0 To MAX_IMPADICDTE
      lDTE.ImpAdic(i).IdImpAdic = 0
      lDTE.ImpAdic(i).IdImpAdicSII = 0
      lDTE.ImpAdic(i).TasaImpAdic = 0
      lDTE.ImpAdic(i).MontoImpAdic = 0
      lDTE.ImpAdic(i).NetoImpAdic = 0
      lDTE.ImpAdic(i).DescImpAdic = ""
   Next i
   
   For j = 0 To MAX_REFDTE
      lDTE.Referencia(j).IdTipoDocRef = 0
      lDTE.Referencia(j).IdDTE = 0
      lDTE.Referencia(j).IdEmpresa = 0
'      lDTE.Referencia(j).Ano = gEmpresa.Ano
      lDTE.Referencia(j).CodDocRefSII = ""
      lDTE.Referencia(j).FolioRef = ""
      lDTE.Referencia(j).FechaRef = 0
      lDTE.Referencia(j).CodRefSII = 0
      lDTE.Referencia(j).RazonReferencia = 0
   Next j

   
End Sub

Private Sub Form_Resize()

   Grid.Height = Me.Height - Grid.Top - Fr_PieDTE.Height - 700
   Grid.Width = Me.Width - Grid.Left - 160
   Fr_PieDTE.Top = Grid.Top + Grid.Height + 100
   Fr_PieDTE.Width = Me.Width - 100
   
   Call LocateFrTotales
   
   Call FGrVRows(Grid, 1)
   
End Sub
Private Sub LocateFrTotales()
   Dim wFrTot As Integer
   
   wFrTot = Fr_Totales.Width
   If Me.WindowState = vbMaximized Then
      Call FGrLocateCntrl(Grid, Fr_Totales, C_SUBTOTAL, True)
      Fr_Totales.Width = wFrTot
      Fr_Totales.Left = Fr_Totales.Left + 60
      If Fr_Totales.Left + Fr_Totales.Width > Me.Width - 200 Or Fr_Totales.Left < lFrmWidth - Fr_Totales.Width - 200 Then
         Fr_Totales.Left = lFrmWidth - Fr_Totales.Width - 200
      End If
   Else
      Fr_Totales.Left = Me.Width - Fr_Totales.Width - 200
   End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
   Call UpdateConfig
End Sub

Private Sub Gr_Ref_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Fecha As Long
   Dim Firstday As Long, LastDay As Long
   Dim FRef As Long

   Fecha = GetTxDate(Tx_Fecha)
   
   Value = Trim(Value)
   Action = vbOK

   Select Case Col
      Case CR_FOLIO
         Value = Trim(Value)
         
      Case CR_FECHA
         Call FirstLastMonthDay(DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)), Firstday, LastDay)
         
         FRef = GetDate(Value, "dmy")
         
         If FRef > LastDay Then
            MsgBox1 "Fecha de referencia posterior a fecha de emisión del presente documento.", vbExclamation
         End If
         
         Gr_Ref.TextMatrix(Row, CR_LNGFECHA) = FRef
         If FRef > 0 Then
            Value = Format(Gr_Ref.TextMatrix(Row, CR_FECHA), SDATEFMT)
         Else
            Value = ""
         End If
         
                      
      Case CR_TIPODOCREF
         Gr_Ref.TextMatrix(Row, CR_IDTIPODOCREF) = lcbDocRef.ItemData
         Gr_Ref.TextMatrix(Row, CR_CODDOCREFSII) = lcbDocRef.Matrix(REF_CODREFSII)
                 
      Case CR_REFSII
         Gr_Ref.TextMatrix(Row, CR_CODREFSII) = CbItemData(Gr_Ref.CbList(CR_REFSII))
                 
   End Select
         
   If Action = vbOK Then
      Call FGrModRow(Gr_Ref, Row, FGR_U, CR_IDREFERENCIA, CR_UPDATE)
      Gr_Ref.TextMatrix(Row, CR_IDDTEREF) = 0    'eliminamos la referencia a un DTE emitido ya que cambiaron los datos, así que ya no corresponde al mismo seleccionado
   End If

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Aux As Single
   
   Value = Trim(Value)
   Action = vbOK

   Select Case Col
      Case C_PJEDESCTO
         Aux = vFmt(Value)
         If Aux < 0 Or Aux > 100 Then
            MsgBox1 "Porcentaje inválido.", vbExclamation
            Action = vbCancel
         End If
         Value = Format(Aux, DBLFMT2)
         Grid.TextMatrix(Row, Col) = Value
         Call CalcRow(Row)
         
      Case C_CANTIDAD
         Value = Format(vFmt(Value), DBLFMT2)
         Grid.TextMatrix(Row, Col) = Value
         Call CalcRow(Row)
             
      Case C_PRECIO
         Value = Format(vFmt(Value), IIf(Ch_DecEnPrecio <> 0, lFmtPrecio2, lFmtPrecio))
         Grid.TextMatrix(Row, Col) = Value
         Call CalcRow(Row)
             
      Case C_IMPADIC
         Grid.TextMatrix(Row, C_IDIMPADIC) = lcbImpAdic.ItemData
         Grid.TextMatrix(Row, C_CODIMPADICSII) = lcbImpAdic.Matrix(IA_CODSIIDTE)
         If vFmt(lcbImpAdic.Matrix(IA_TASA)) = 0 Then
            Grid.TextMatrix(Row, C_TASAIMPADIC) = ""
            Call FGrForeColor(Grid, Row, C_TASAIMPADIC, vbBlue)
            Grid.TextMatrix(Row, C_TASAEDITABLE) = 1
         Else
            Grid.TextMatrix(Row, C_TASAIMPADIC) = Format(vFmt(lcbImpAdic.Matrix(IA_TASA)), DBLFMT2)
            Call FGrForeColor(Grid, Row, C_TASAIMPADIC, vbBlack)
            Grid.TextMatrix(Row, C_TASAEDITABLE) = 0
         End If
         Grid.TextMatrix(Row, C_MONTOIMPADIC) = ""
            
         CalcRow (Row)

         
      Case C_TASAIMPADIC
         Value = Format(vFmt(Value), DBLFMT2)
         Grid.TextMatrix(Row, Col) = Value
         Call CalcRow(Row)
         
   End Select
         
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDDETFACT, C_UPDATE)
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lIdEntidad = 0 Or Not lEntCompleta Then
      MsgBox1 "Debe seleccionar la entidad receptora y completar todos los datos antes de continuar.", vbExclamation
      Exit Sub
   End If
   
   If Grid.Row > Grid.FixedRows And (Grid.TextMatrix(Row - 1, C_PRODUCTO) = "" Or vFmt(Grid.TextMatrix(Row - 1, C_CANTIDAD)) = 0 Or vFmt(Grid.TextMatrix(Row - 1, C_PRECIO)) = 0) Then
      MsgBox1 "Debe completar el registro anterior.", vbExclamation
      Exit Sub
   End If
   
   If Row - Grid.FixedRows + 1 >= MAX_ITEMDTE Then
      MsgBox1 "Ha superado la cantidad de ítems en un documento electrónico.", vbExclamation
      Exit Sub
   End If

   Grid.TextMatrix(Row, C_NUMLIN) = Row - Grid.FixedRows + 1
   If Grid.ColWidth(C_ESEXENTO) = 0 Then
      Grid.TextMatrix(Row, C_ESEXENTO) = "x"
   End If
   
   If Row >= Grid.rows - 2 Then
      Grid.rows = Grid.rows + 1
   End If

   Select Case Col
      Case C_TIPOCOD, C_CODPROD, C_PRODUCTO, C_UMEDIDA, C_PRECIO
         
         If Val(Grid.TextMatrix(Row, C_IDPROD)) <> 0 Then
            If (Col = C_PRECIO Or Col = C_PRODUCTO) And vFmt(Grid.TextMatrix(Row, C_PRECIOORI)) = 0 Then
               EdType = FEG_Edit
            ElseIf Col = C_PRECIO And Ch_ModPrecio <> 0 Then
               EdType = FEG_Edit
            Else
               Call Bt_SelProd_Click
            End If
         ElseIf lNotSelProd = False And Grid.TextMatrix(Row, C_PRODUCTO) = "" Then
            Call Bt_SelProd_Click
            If Val(Grid.TextMatrix(Row, C_IDPROD)) = 0 Then
               EdType = FEG_Edit
            End If
         Else
            EdType = FEG_Edit
         End If
         
         Grid.TxBox.MaxLength = lFldLen(Col)
      
      Case C_DESCRIP, C_CANTIDAD, C_PJEDESCTO
         If Grid.TextMatrix(Row, C_PRODUCTO) <> "" Then
            EdType = FEG_Edit
            Grid.TxBox.MaxLength = lFldLen(Col)
            If Col = C_DESCRIP Then
               Grid.TxBox.MaxLength = 1000         'estipulado por el SII
            End If
            
         Else
            MsgBox1 "Debe ingresar el producto.", vbExclamation
         End If
      
        
      Case C_IMPADIC
         If Grid.TextMatrix(Row, C_ESEXENTO) = "" Then
            If Grid.TextMatrix(Row, C_PRODUCTO) <> "" And vFmt(Grid.TextMatrix(Row, C_CANTIDAD)) > 0 Then
               EdType = FEG_List
            Else
               MsgBox1 "Debe ingresar el producto y la cantidad.", vbExclamation
            End If
         End If
         
      Case C_TASAIMPADIC
         If Val(Grid.TextMatrix(Row, C_TASAEDITABLE)) <> 0 Then
            EdType = FEG_Edit
         End If
         
      Case C_ESEXENTO
      
         If Trim(Grid.TextMatrix(Row, Col)) = "" Then
            Grid.TextMatrix(Row, Col) = "x"
            Grid.TextMatrix(Row, C_IMPADIC) = ""
            Grid.TextMatrix(Row, C_IDIMPADIC) = ""
            Grid.TextMatrix(Row, C_TASAIMPADIC) = ""
            Grid.TextMatrix(Row, C_CODIMPADICSII) = ""
         Else
            Grid.TextMatrix(Row, Col) = ""
            Lb_TasaIVA = gIVA * 100 & "%"     'por si lo hubieramos eliminado en el caso de una NCV o NDV exenta
         End If
         
         Call FGrModRow(Grid, Row, FGR_U, C_IDDETFACT, C_UPDATE)
         CalcRow (Row)
         
   End Select
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Dim Col As Integer
   
   Col = Grid.Col

   Select Case Col
      Case C_TIPOCOD, C_CODPROD
         Call KeyCod(KeyAscii)
         
      Case C_PRODUCTO, C_DESCRIP, C_UMEDIDA
         Call KeyName(KeyAscii)
         
      Case C_CANTIDAD
         Call KeyDecPos(KeyAscii)
         
      Case C_PRECIO
         Call KeyDecPos(KeyAscii)
         
      Case C_PJEDESCTO, C_TASAIMPADIC
         Call KeyDecPos(KeyAscii)
         
   End Select


End Sub

Private Sub Gr_Ref_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   Dim Fecha As Long

   Fecha = GetTxDate(Tx_Fecha)
   
   If Row < Gr_Ref.FixedRows Then
      Exit Sub
   End If
   
   If lIdEntidad = 0 Or Not lEntCompleta Then
      MsgBox1 "Debe seleccionar la entidad receptora y completar todos los datos antes de continuar.", vbExclamation
      Exit Sub
   End If
   
   'verificamos que haya al menos un registro en el detalle de fatura
   If Grid.TextMatrix(Grid.FixedRows, C_PRODUCTO) = "" Or vFmt(Grid.TextMatrix(Grid.FixedRows, C_CANTIDAD)) = 0 Or vFmt(Grid.TextMatrix(Grid.FixedRows, C_PRECIO)) = 0 Then
      If lDiminutivoDoc <> "NCV" Then
         MsgBox1 "Debe completar por lo menos un registro del detalle del documento electrónico.", vbExclamation
         Exit Sub
      End If
   End If
  
   If Gr_Ref.Row > Gr_Ref.FixedRows And (Gr_Ref.TextMatrix(Row - 1, CR_TIPODOCREF) = "" Or Gr_Ref.TextMatrix(Row - 1, CR_FOLIO) = "" Or Gr_Ref.TextMatrix(Row - 1, CR_FECHA) = "") Then
      MsgBox1 "Debe completar el registro anterior.", vbExclamation
      Exit Sub
   End If
   
   If Row - Gr_Ref.FixedRows >= MAX_REFDTE Then
      MsgBox1 "Ha superado la cantidad de ítems en las referencias de un documento electrónico.", vbExclamation
      Exit Sub
   End If

   Gr_Ref.TextMatrix(Row, CR_NUMLIN) = Row - Gr_Ref.FixedRows + 1
   
   If Row >= Gr_Ref.rows - 2 Then
      Gr_Ref.rows = Gr_Ref.rows + 1
   End If

   If Col <> CR_TIPODOCREF And Val(Gr_Ref.TextMatrix(Row, CR_IDTIPODOCREF)) = 0 Then
      MsgBox1 "Debe ingresar el tipo de documento de referencia.", vbExclamation
      Exit Sub
   End If

   Select Case Col
      Case CR_TIPODOCREF, CR_REFSII
         EdType = FEG_List
         
      Case CR_FOLIO, CR_RAZONREF
         
         Gr_Ref.TxBox.MaxLength = lFldLenRef(Col)
         EdType = FEG_Edit
         
      Case CR_FECHA
      
         If vFmt(Gr_Ref.TextMatrix(Row, CR_LNGFECHA)) > 0 Then
            Gr_Ref.TextMatrix(Row, CR_FECHA) = Format(Gr_Ref.TextMatrix(Row, CR_LNGFECHA), SDATEFMT)
         Else
            Gr_Ref.TextMatrix(Row, CR_LNGFECHA) = CLng(DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)))
            Gr_Ref.TextMatrix(Row, CR_FECHA) = Format(Val(Gr_Ref.TextMatrix(Row, CR_LNGFECHA)), SDATEFMT)
            
         End If
         
         Gr_Ref.TxBox.MaxLength = 8
         EdType = FEG_Edit
                 
   End Select
   
End Sub
Private Sub Ls_CopiarDTE_Click()
   Dim IdDTE As Long
   Dim FrmAdm As FrmAdmDTE
   Dim Rc As Integer

   If lInLoad Then
      Exit Sub
   End If
   
   If CbItemData(Ls_CopiarDTE) = COPY_ULTDTE Then
      Call LoadDTE(0)

   Else
      MsgBox1 "Seleccione el documento que desea copiar.", vbInformation
      Set FrmAdm = New FrmAdmDTE
      Rc = FrmAdm.FSelectCopy(lCodDocSII, IdDTE)
      Set FrmAdm = Nothing

      If Rc = vbOK Then
         LoadDTE (IdDTE)
      End If
   End If
   
   Call Bt_CopiarDTE_Click
   
End Sub

Private Sub Ls_CopiarDTE_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Ls_CopiarDTE.Visible = False
   End If
End Sub

Private Sub Tx_Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Tx_Codigo.Text) <> "" Then
            If Not CargaCbVendedor(Tx_Codigo.Text) Then
                MsgBox1 "Vendedor No existe o se encuentra Inactivo ", vbExclamation
                Tx_Codigo.Text = ""
            End If
        End If
    End If
End Sub

Private Sub Tx_Codigo_Validate(Cancel As Boolean)
If Trim(Tx_Codigo.Text) <> "" Then
    If Not CargaCbVendedor(Tx_Codigo.Text) Then
        MsgBox1 "Vendedor No existe o se encuentra Inactivo ", vbExclamation
        Tx_Codigo.Text = ""
    End If
End If
End Sub

Private Sub Tx_Fecha_GotFocus()
   Call DtGotFocus(Tx_Fecha)
End Sub

Private Sub Tx_Fecha_LostFocus()
   
   If Trim$(Tx_Fecha) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Fecha)
   
   Call SetTxDate(Tx_FechaVenc, DateAdd("d", 30, GetTxDate(Tx_Fecha)))
   
'   If Year(GetTxDate(Tx_Fecha)) <> gEmpresa.Ano Then
'      MsgBox1 "La fecha debe corresponder al año con col cual está trabajando (" & gEmpresa.Ano & ")", vbExclamation
'   End If
End Sub

Private Sub Tx_Fecha_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_FechaVenc_GotFocus()
   Call DtGotFocus(Tx_FechaVenc)
End Sub

Private Sub Tx_FechaVenc_LostFocus()
   
   If Trim$(Tx_FechaVenc) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FechaVenc)
   
End Sub

Private Sub Tx_FechaVenc_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub UpdateConfig()
   Dim i As Integer
   Dim Rs As Recordset
   
   For i = 1 To MAX_OPTEDFACT

      Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpDTE WHERE Tipo='OPTEDFACT' AND Codigo=" & i & " AND IdEmpresa = " & gEmpresa.Id)
      
         If Rs.EOF = False Then
            'actualizamos
            Call ExecSQL(DbMain, "UPDATE ParamEmpDTE SET Valor = '" & gEmpConfig.OptEdFact(i) & "' WHERE Tipo = 'OPTEDFACT' AND Codigo =" & i & " AND IdEmpresa = " & gEmpresa.Id)
         Else
            'insertamos
            Call ExecSQL(DbMain, "INSERT INTO ParamEmpDTE (IdEmpresa, Tipo, Codigo, Valor) VALUES (" & gEmpresa.Id & ", 'OPTEDFACT'," & i & ", '" & gEmpConfig.OptEdFact(i) & "')")
         End If
      
      Call CloseRs(Rs)

   Next i
End Sub

Private Sub Tx_PjeDescto_KeyPress(KeyAscii As Integer)
   Call KeyDecPos(KeyAscii)
End Sub

Private Sub Tx_PjeDescto_LostFocus()

   Tx_PjeDescto = Format(Min(vFmt(Tx_PjeDescto), 100), DBLFMT2)
   
   Call CalcTotal
End Sub


Private Sub Tx_Rut_LostFocus()
   
   If Tx_RUT = "" Then
      Call LimpiarEntidad
      Exit Sub
   End If
   
   If vFmtCID(Tx_RUT) = 0 Then
      Call LimpiarEntidad
      Tx_RUT = ""
      Tx_RUT.SetFocus
      Exit Sub
   End If
   
'   If Not MsgValidRut(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
'
   
   Tx_RUT = FmtCID(vFmtCID(Tx_RUT))
   
   Call FillDataRUT
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If Trim(Tx_RUT) = "0-0" Then
      MsgBox1 "RUT Inválido.", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   
   If Not MsgValidRut(Tx_RUT) Then
      Call LimpiarEntidad
      Tx_RUT.SetFocus
      lIdEntidad = 0
      Cancel = True
      Exit Sub
      
   End If
   
   
End Sub
Private Sub Bt_SelEnt_Click(Index As Integer)
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   Dim Row As Integer
   Dim TipoEnt As Integer
   Dim Col As Integer
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim Q1 As String
      
   TipoEnt = ENT_CLIENTE
   
   Set Frm = New FrmEntidades
   Rc = Frm.FSelEdit(Entidad, TipoEnt, IIf(lEsExport, True, False))
   Set Frm = Nothing
   
   If Rc <> vbOK Then
      If Not lEsExport Then
         If Trim(Tx_RUT) <> "" Then
            Call Tx_Rut_LostFocus
         End If
      End If
      Exit Sub
   End If
               
   If Entidad.NotValidRut <> 0 And Not lEsExport Then
      MsgBox1 "Rut inválido.", vbExclamation
      Exit Sub
   End If
   
   Call Bt_Limpiar_Click
   
   lIdEntidad = Entidad.Id
   Tx_RUT = FmtCID(Entidad.Rut, Not Entidad.NotValidRut)
   Ch_Rut = IIf(Not Entidad.NotValidRut, 1, 0)
   Tx_RazonSocial = Entidad.Nombre
         
   Q1 = "SELECT Direccion, Regiones.Comuna, Ciudad, Giro, EMail "
   Q1 = Q1 & " FROM Entidades LEFT JOIN Regiones ON Entidades.Comuna = Regiones.Id "
   Q1 = Q1 & " WHERE IdEntidad = " & lIdEntidad & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Tx_Direccion = vFld(Rs("Direccion"))
      Tx_Comuna = vFld(Rs("Comuna"))
      Tx_Ciudad = vFld(Rs("Ciudad"))
      Tx_Giro = vFld(Rs("Giro"))
      Tx_MailReceptor = vFld(Rs("EMail"))
   End If
   
   Call CloseRs(Rs)

   lEntCompleta = True
   
   If Not lEsExport Then
      If Tx_Direccion = "" Or Tx_Comuna = "" Or Tx_Ciudad = "" Or Tx_Giro = "" Then
         MsgBox1 "Debe completar los datos de la entidad antes de continuar.", vbExclamation
         lEntCompleta = False
      End If
   ElseIf Tx_Ciudad = "" Then
      MsgBox1 "Falta ingresar la ciudad de la entidad antes de continuar.", vbExclamation
      Tx_Ciudad.SetFocus
      lEntCompleta = False
   End If
   
End Sub

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing

End Sub

Private Sub CalcRow(ByVal Row As Integer)
   Dim Tot As Double
   Dim Descto As Double
   
   Tot = vFmt(Grid.TextMatrix(Row, C_PRECIO)) * vFmt(Grid.TextMatrix(Row, C_CANTIDAD))
   
   If vFmt(Grid.TextMatrix(Row, C_PJEDESCTO)) > 0 Then
      Descto = Round(Tot * vFmt(Grid.TextMatrix(Row, C_PJEDESCTO) / 100))
   End If
      
   Grid.TextMatrix(Row, C_MONTODESCTO) = Format(Descto, NUMFMT)
   Grid.TextMatrix(Row, C_SUBTOTAL) = Format(Tot - Descto, lFmtPrecio)
   
   Call CalcTotal
End Sub

Private Sub CalcRowIVA(ByVal Row As Integer)
   Dim Tot As Double
   Dim Descto As Double
   
   Tot = vFmt(Grid.TextMatrix(Row, C_PRECIO)) * vFmt(Grid.TextMatrix(Row, C_CANTIDAD))
   
   If vFmt(Grid.TextMatrix(Row, C_PJEDESCTO)) > 0 Then
      Descto = Round(Tot * vFmt(Grid.TextMatrix(Row, C_PJEDESCTO) / 100))
   End If
      
   'Grid.TextMatrix(Row, C_MONTODESCTO) = Format(Descto, NUMFMT)
   Grid.TextMatrix(Row, C_SUBTOTAL) = Format(Tot - Descto, lFmtPrecio)
   
   Call CalcTotal
End Sub

Private Sub CalcTotal()
   Dim Aux As Double, AuxImp As Double
   Dim i As Integer
   Dim Total As Double
   Dim ImpTotal As Double
   Dim PjeDescto As Single
   Dim MontoDesctoAfecto As Double
   Dim ImpAdic As Single
   Dim Iva As Single
   
   PjeDescto = vFmt(Tx_PjeDescto) / 100
   lTotAfecto = 0
   lTotExento = 0
   lTotDescIva = 0
   rebajaIva = False
   
   Grid.Refresh
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_PRODUCTO) = "" Then
         Exit For
      End If
      
      If Grid.RowHeight(i) > 0 Then
         Total = Total + vFmt(Grid.TextMatrix(i, C_SUBTOTAL))
         If Trim(Grid.TextMatrix(i, C_ESEXENTO)) = "" Then
            lTotAfecto = lTotAfecto + vFmt(Grid.TextMatrix(i, C_SUBTOTAL))
         Else
            lTotExento = lTotExento + vFmt(Grid.TextMatrix(i, C_SUBTOTAL))
         End If
         
         If vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)) > 0 And Grid.TextMatrix(i, C_IDIMPADIC) <> "8" Then
            Aux = vFmt(Grid.TextMatrix(i, C_SUBTOTAL))
            
            If vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)) = 100 Then
               ImpAdic = gIVA * 100
            Else
               ImpAdic = vFmt(Grid.TextMatrix(i, C_TASAIMPADIC))
            End If
            
            AuxImp = Round(Aux * ImpAdic / 100)
            If PjeDescto > 0 Then
               AuxImp = Round(AuxImp * (1 + PjeDescto / 100))
            End If
            Grid.TextMatrix(i, C_MONTOIMPADIC) = Format(AuxImp, NUMFMT)   'habiendo aplicado el descuento global
            ImpTotal = ImpTotal + AuxImp
         ElseIf vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)) > 0 And Grid.TextMatrix(i, C_IDIMPADIC) = "8" Then
            rebajaIva = True
            If vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)) = 100 Then
               ImpAdic = gIVA * 100
            Else
               ImpAdic = vFmt(Grid.TextMatrix(i, C_TASAIMPADIC))
            End If
            Iva = Round(Round(vFmt(Grid.TextMatrix(i, C_SUBTOTAL)) * vFmt(Lb_TasaIVA)) * (ImpAdic / 100))
            Grid.TextMatrix(i, C_DESCIVA) = Iva
            lTotDescIva = lTotDescIva + Iva
            ImpTotal = lTotDescIva
         End If

      End If
   Next i
         
   Tx_SubTotal = Format(Total, lFmtPrecio)
   Tx_MontoDescto = Format(PjeDescto * Total, lFmtPrecio)
   
   Tx_Neto = Format(Total - vFmt(Tx_MontoDescto), lFmtPrecio)
   
   MontoDesctoAfecto = vFmt(Format(lTotAfecto * PjeDescto, lFmtPrecio))   'para que haga el redondeo igual que los otros
   lTotAfecto = lTotAfecto - MontoDesctoAfecto
   lTotExento = vFmt(Tx_Neto) - lTotAfecto
   Tx_IVA = Format(lTotAfecto * vFmt(Lb_TasaIVA), NUMFMT)
   Tx_ImpAdic = Format(ImpTotal, NUMFMT)
   If lDiminutivoDoc = "FCV" Or lEsNotaCredDebFactCompra Or rebajaIva Then
      Tx_Total = Format(vFmt(Tx_Neto) + vFmt(Tx_IVA) - vFmt(Tx_ImpAdic), lFmtPrecio)
   Else
      Tx_Total = Format(vFmt(Tx_Neto) + vFmt(Tx_IVA) + vFmt(Tx_ImpAdic), lFmtPrecio)
      
   End If
   If rebajaIva Then
    Lb_ImpAdic.Caption = "Crédito 65% Empresa Constructora"
   Else
    Lb_ImpAdic.Caption = "Impuesto Adicional"
   End If
   
   

End Sub


Private Sub FillDataRUT()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Rut As String
   
   lIdEntidad = 0
   Rut = vFmtCID(Tx_RUT)
   If Rut = "" Then
      Exit Sub
   End If
   
   Q1 = "SELECT IdEntidad, Nombre, Direccion, Regiones.Comuna, Ciudad, Giro, EMail "
   Q1 = Q1 & " FROM Entidades LEFT JOIN Regiones ON Entidades.Comuna = Regiones.Id "
   Q1 = Q1 & " WHERE Rut = '" & Rut & "'" & " AND IdEmpresa = " & gEmpresa.Id
   
   lIdEntidad = 0
   Tx_RazonSocial = ""
   Tx_Direccion = ""
   Tx_Comuna = ""
   Tx_Ciudad = ""
   Tx_Giro = ""
   
   lEntCompleta = True
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      lIdEntidad = vFld(Rs("IdEntidad"))
      Tx_RazonSocial = vFld(Rs("Nombre"))
      Tx_Direccion = vFld(Rs("Direccion"))
      If Not lEsExport Then
         Tx_Comuna = vFld(Rs("Comuna"))
      End If
      Tx_Ciudad = vFld(Rs("Ciudad"))
      Tx_Giro = vFld(Rs("Giro"))
      Tx_MailReceptor = vFld(Rs("EMail"))
      
      If Not lInLoad Then
         If Not lEsExport Then
            If Tx_Direccion = "" Or Tx_Comuna = "" Or Tx_Ciudad = "" Or Tx_Giro = "" Then
               MsgBox1 "Debe completar los datos de la entidad antes de continuar.", vbExclamation
               lEntCompleta = False
            End If
         ElseIf Tx_Ciudad = "" Then
            MsgBox1 "Falta ingresar la ciudad de la entidad antes de continuar.", vbExclamation
            Tx_Ciudad.SetFocus
            lEntCompleta = False
         End If
      End If
   Else
      MsgBox1 "No existe una entidad creada con este RUT." & vbCrLf & vbCrLf & "Verifique los datos en la lista de entidades.", vbExclamation
      lEntCompleta = False
      
   End If
   
   
   Call CloseRs(Rs)
   
   
End Sub

Private Function Valida() As Boolean
   Dim i As Integer
   Dim EsExento As Integer
   Dim nReg As Integer
   Dim Fecha As Long, FechaVenc As Long
   Dim ValidaDatosExp As Boolean
   Dim ValidaDatosGuiaDesp As Boolean
   
   Valida = False
   
   If Not lEsExport Then
      If Trim(Tx_RUT) = "" Or Trim(Tx_RazonSocial) = "" Or Trim(Tx_Direccion) = "" Or Trim(Tx_Comuna) = "" Or Trim(Tx_Ciudad) = "" Or Trim(Tx_Giro) = "" Then
         MsgBox1 "Debe completar los datos de la entidad antes de continuar.", vbExclamation
         lEntCompleta = False
         Exit Function
      End If
      
   ElseIf Tx_RUT = "" Or Tx_RazonSocial = "" Or Trim(Tx_Ciudad) = "" Then
      MsgBox1 "Debe completar los datos de la entidad antes de continuar.", vbExclamation
      lEntCompleta = False
      Exit Function
   End If
   
   If gConectData.Proveedor = PROV_ACEPTA And Trim(Tx_MailReceptor) = "" And lEsExport = False Then ' 29 sep 2020: se agrega  And lEsExport=false
      MsgBox1 "Debe ingresar el mail del receptor en los datos de la entidad antes de continuar." & vbCrLf & vbCrLf & "Esto permitirá que le llegue automáticamente el documento electrónico, una vez que éste haya sido aceptado por el SII.", vbExclamation
      lEntCompleta = False
      Exit Function
   End If
   
   Fecha = GetTxDate(Tx_Fecha)
   FechaVenc = GetTxDate(Tx_FechaVenc)
   
   If Fecha > FechaVenc Then
      MsgBox1 "La fecha de vencimiento debe ser posterior a la fecha de emisión del documento.", vbExclamation
      Exit Function
   End If
   
   'Validaciones del SII para las fechas
   If Fecha < DateSerial(2003, 4, 1) Or Fecha > DateSerial(2050, 12, 31) Then
      MsgBox1 "La fecha de emisión es inválida.", vbExclamation
      Exit Function
   End If
   
   If FechaVenc < DateSerial(2002, 8, 1) Or FechaVenc > DateSerial(2050, 12, 31) Then
      MsgBox1 "La fecha de vencimiento es inválida.", vbExclamation
      Exit Function
   End If
   
   If lEsExport Then
      ValidaDatosExp = True

      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodIndServicio <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodPais <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodPuertoEmbarque <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodPuertoDesembarque <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodMoneda <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodModVenta <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodClausulaVenta <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.CodViaTransporte <> ""
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.TipoCambio <> 0
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.TotalBultos <> 0
      ValidaDatosExp = ValidaDatosExp And lDTE.FactExp.TotClausulaVenta <> 0
   
      If Not ValidaDatosExp Then
         MsgBox1 "Falta completar los datos de la factura de exportación." & vbCrLf & vbCrLf & "Utilice el botón 'Datos Factura de Exportación...'", vbExclamation
         Exit Function
      End If
   End If
   
   If lEsGuiaDespacho Then
      ValidaDatosGuiaDesp = True

      ValidaDatosGuiaDesp = IIf(CbItemData(Cb_TipoDespacho) = GD_DESPEMICLI Or CbItemData(Cb_TipoDespacho) = GD_DESPEMIOTRO, ValidaDatosGuiaDesp And lDTE.GuiaDesp.Patente <> "", ValidaDatosGuiaDesp)
      ValidaDatosGuiaDesp = IIf(CbItemData(Cb_TipoDespacho) = GD_DESPEMICLI Or CbItemData(Cb_TipoDespacho) = GD_DESPEMIOTRO, ValidaDatosGuiaDesp And lDTE.GuiaDesp.RutChofer <> "", ValidaDatosGuiaDesp)
      
      If Not ValidaDatosGuiaDesp Then
         MsgBox1 "Falta completar los datos de la guía de despacho." & vbCrLf & vbCrLf & "Utilice el botón 'Datos Adicionales...'", vbExclamation
         Exit Function
      End If
      
      ValidaDatosGuiaDesp = IIf(CbItemData(Cb_TipoDespacho) <> GD_DESPEMICLI And CbItemData(Cb_TipoDespacho) <> GD_DESPEMIOTRO, ValidaDatosGuiaDesp And lDTE.GuiaDesp.Patente = "", ValidaDatosGuiaDesp)
      ValidaDatosGuiaDesp = IIf(CbItemData(Cb_TipoDespacho) <> GD_DESPEMICLI And CbItemData(Cb_TipoDespacho) <> GD_DESPEMIOTRO, ValidaDatosGuiaDesp And lDTE.GuiaDesp.RutChofer = "", ValidaDatosGuiaDesp)
            
      If Not ValidaDatosGuiaDesp Then
         MsgBox1 "Debe eliminar los datos de Patente, RUT y nombre del chofer. Éstos no se requieren, dado el tipo de despacho seleccionado." & vbCrLf & vbCrLf & "Utilice el botón 'Datos Adicionales...'", vbExclamation
         Exit Function
      End If
      
   End If
      
   EsExento = 0
   nReg = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_PRODUCTO) = "" Then
         If i = Grid.FixedRows Then
            If lDiminutivoDoc <> "NCV" Or (lDiminutivoDoc = "NCV" And lTipoRef = REF_ANULA) Then
               MsgBox1 "Debe ingresar nombre del primer ítem del detalle.", vbExclamation
               Exit Function
            End If
         End If
         Exit For
      End If
      
      nReg = nReg + 1
      
      If vFmt(Grid.TextMatrix(i, C_CANTIDAD)) = 0 Then
         If lDiminutivoDoc <> "NCV" Or (lDiminutivoDoc = "NCV" And lTipoRef = REF_ANULA) Then
            MsgBox1 "Registro " & i & " incompleto. La cantidad debe ser mayor a cero.", vbExclamation
            Exit Function
         End If
      End If
      
      If vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)) > 100 Then
         MsgBox1 "Registro " & i & " inválido. La tasa es mayor a 100%.", vbExclamation
         Exit Function
      End If
      
      If ((lDiminutivoDoc = "FCV" Or (lDiminutivoDoc = "NCV" Or lDiminutivoDoc = "NDV") And lEsNotaCredDebFactCompra)) And Grid.TextMatrix(i, C_ESEXENTO) = "" Then    'es factura de compra o nota de crédito o débito de factura de compra y el producto no es exento
         If vFmt(Grid.TextMatrix(i, C_IDIMPADIC)) = 0 Then
            MsgBox1 "Registro " & i & " inválido. Falta seleccionar un impuesto adicional.", vbExclamation
            Exit Function
         End If
         If vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)) = 0 Then
            MsgBox1 "Registro " & i & " inválido. Falta ingresar la tasa del impuesto adicional.", vbExclamation
            Exit Function
         End If
      End If
         
      If vFmt(Grid.TextMatrix(i, C_PRECIO)) = 0 Then
         If lDiminutivoDoc <> "NCV" Or (lDiminutivoDoc = "NCV" And lTipoRef = REF_ANULA) Then
            MsgBox1 "Registro " & i & " incompleto. El precio debe ser mayor a cero.", vbExclamation
            Exit Function
         End If
      End If
      
      If Trim(Grid.TextMatrix(i, C_ESEXENTO)) <> "" Then
         EsExento = EsExento + 1
      End If
      
   Next i
   
   If lTieneExento = True And Not lSoloExento Then
      If EsExento > 0 And EsExento >= nReg Then
      
         If lDiminutivoDoc = "NCV" Or lDiminutivoDoc = "NDV" Then    'es nota de crédito o nota de débito => permitimos que sea exenta (todos los itemes exentos)
            If MsgBox1("ATENCIÓN: Todos los ítems del documento son exentos de impuesto." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            Else
               Lb_TasaIVA = 0
            End If
         ElseIf MsgBox1("ATENCIÓN: Dado que todos los ítems del documento son exentos de impuesto, es posible que el SII no acepte este documento afecto." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
            Exit Function
         End If
         
      End If
   End If
   
   If vFmt(Tx_Total) = 0 Then
      If lDiminutivoDoc <> "NCV" Or (lDiminutivoDoc = "NCV" And lTipoRef = REF_ANULA) Then
         MsgBox1 "Total DTE debe ser mayor a cero.", vbExclamation
         Exit Function
      End If
   End If
   
   If lDiminutivoDoc = "FAV" And vFmt(Tx_IVA) = 0 Then
      If MsgBox1("ATENCIÓN: Valor IVA debe ser mayor a cero." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
         Exit Function
      End If
   End If
   
   If lEsGuiaDespacho Then
      If Cb_TipoDespacho.ListIndex < 0 Then
         MsgBox1 "Falta ingresar el tipo de despacho.", vbExclamation
         Cb_TipoDespacho.SetFocus
         Exit Function
      ElseIf Cb_Traslado.ListIndex < 0 Then
         MsgBox1 "Falta ingresar el indicador de tipo de traslado.", vbExclamation
         Cb_Traslado.SetFocus
         Exit Function
      End If
   End If
   
   For i = Gr_Ref.FixedRows To Gr_Ref.rows - 1
   
      If Gr_Ref.TextMatrix(i, CR_TIPODOCREF) <> "" Then
         If Trim(Gr_Ref.TextMatrix(i, CR_FOLIO)) = "" Or Val(Gr_Ref.TextMatrix(i, CR_LNGFECHA)) = 0 Then
            MsgBox1 "Falta completar los datos del documento de referencia. Si no desea ingresarlo, debe dejar el registro de referencia en blanco.", vbExclamation
            Exit Function
         End If
         If Trim(Gr_Ref.TextMatrix(i, CR_REFSII)) = "" And (lDiminutivoDoc = "NCV" Or lDiminutivoDoc = "NCV") Then
            MsgBox1 "Falta ingresar el Tipo de Ref. SII para el caso de notas de crédito o débito.", vbExclamation
            Exit Function
         End If
         
      ElseIf Trim(Gr_Ref.TextMatrix(i, CR_FOLIO)) <> "" Or Val(Gr_Ref.TextMatrix(i, CR_LNGFECHA)) <> 0 Or Trim(Gr_Ref.TextMatrix(i, CR_REFSII)) <> "" Then
         MsgBox1 "Falta completar los datos del documento de referencia. Si no desea ingresarlo, debe dejar el registro de referencia en blanco.", vbExclamation
         Exit Function
      End If

   Next i
   
   Valida = True
   
End Function

Private Sub SaveAll()
   Dim i As Integer, j As Integer, k As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim lIdDTEFactExp As Long
   Dim lIdDTEGuiaDesp As Long
   
   Call GenDTEStruct
   
   Set Rs = DbMain.OpenRecordset("DTE")
   Rs.AddNew
   
   lIdDTE = Rs("IdDTE")
   
   lDTE.IdDTE = lIdDTE
   
   'Encabezado
   
   Rs.Fields("IdEmpresa") = gEmpresa.Id
'   Rs.Fields("Ano") = gEmpresa.Ano
   Rs.Fields("TipoDoc") = lDTE.TipoDoc
   Rs.Fields("TipoLib") = lDTE.TipoLib
   Rs.Fields("CodDocSII") = lDTE.CodDocSII
   Rs.Fields("Fecha") = lDTE.Fecha
   Rs.Fields("Folio") = lDTE.Folio
   Rs.Fields("IdEstado") = lDTE.idEstado
   Rs.Fields("IdEntidad") = lDTE.IdEntidad
   Rs.Fields("RUT") = lDTE.Rut
   Rs.Fields("NotValidRut") = lDTE.NotValidRut
   Rs.Fields("Contacto") = ParaSQL(lDTE.Contacto)
   Rs.Fields("Referencias") = 0
   Rs.Fields("InfoPago") = 0
   Rs.Fields("SubTotal") = lDTE.SubTotal
   Rs.Fields("PjeDesctoGlobal") = lDTE.PjeDestoGlobal
   Rs.Fields("DesctoGlobal") = vFmt(Tx_MontoDescto)
   Rs.Fields("Exento") = lDTE.Exento
   Rs.Fields("Neto") = lDTE.Neto
   Rs.Fields("IVA") = lDTE.Iva
   Rs.Fields("ImpAdicional") = lDTE.ImpAdicional
   Rs.Fields("Total") = lDTE.Total
   Rs.Fields("IdUsuario") = gUsuario.IdUsuario
   Rs.Fields("FechaCreacion") = CLng(Int(Now))
  
   If lEsGuiaDespacho Then
      Rs.Fields("TipoDespacho") = lDTE.TipoDespacho    'estos 4 se agregan el 10 de agosto 2018
      Rs.Fields("Traslado") = lDTE.Traslado
   End If
   
   Rs.Fields("FechaVenc") = lDTE.FechaVenc
   Rs.Fields("FormaDePago") = lDTE.FormaDePago
   Rs.Fields("ObsDTE") = ParaSQL(Tx_ObsDTE)
   Rs.Fields("DetFormaPago") = lDTE.DetFormaPago
   Rs.Fields("Vendedor") = lDTE.Vendedor
   
   Rs.Update
   Rs.Close
   Set Rs = Nothing
   
   
   'Antecedentes Factura de Exportación
   If lEsExport Then
   
      Set Rs = DbMain.OpenRecordset("DTEFactExp")
      Rs.AddNew
   
      lIdDTEFactExp = Rs("IdDTEFactExp")
   
      lDTE.FactExp.IdDTEFactExp = lIdDTEFactExp
      
      Rs.Fields("IdDTE") = lIdDTE
      Rs.Fields("IdEmpresa") = gEmpresa.Id
      
      Rs.Fields("CodIndServicio") = lDTE.FactExp.CodIndServicio
      Rs.Fields("CodPais") = lDTE.FactExp.CodPais
      Rs.Fields("CodPuertoEmbarque") = lDTE.FactExp.CodPuertoEmbarque
      Rs.Fields("CodPuertoDesembarque") = lDTE.FactExp.CodPuertoDesembarque
      Rs.Fields("CodMoneda") = lDTE.FactExp.CodMoneda
      Rs.Fields("TipoCambioPesos") = lDTE.FactExp.TipoCambio
      Rs.Fields("TotalBultos") = lDTE.FactExp.TotalBultos
      Rs.Fields("CodModVenta") = lDTE.FactExp.CodModVenta
      Rs.Fields("CodClauCompraVenta") = lDTE.FactExp.CodClausulaVenta
      Rs.Fields("TotalClauVenta") = lDTE.FactExp.TotClausulaVenta
      Rs.Fields("CodViaTransporte") = lDTE.FactExp.CodViaTransporte
                
      Rs.Update
      Rs.Close
      Set Rs = Nothing
      
   End If
   
   'Antecedentes Guia Despacho
   If lEsGuiaDespacho Then
   
      Set Rs = DbMain.OpenRecordset("DTEGuiaDesp")
      Rs.AddNew
   
      lIdDTEGuiaDesp = Rs("IdDTEGuiaDesp")
   
      lDTE.GuiaDesp.IdDTEGuiaDesp = lIdDTEGuiaDesp
      
      Rs.Fields("IdDTE") = lIdDTE
      Rs.Fields("IdEmpresa") = gEmpresa.Id
      
      Rs.Fields("Patente") = ParaSQL(lDTE.GuiaDesp.Patente)
      Rs.Fields("RutChofer") = ParaSQL(lDTE.GuiaDesp.RutChofer)
      Rs.Fields("NombreChofer") = ParaSQL(lDTE.GuiaDesp.NombreChofer)
                
      Rs.Update
      Rs.Close
      Set Rs = Nothing
      
   End If
   
   
   
   j = 0
   
   'Detalle de productos
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_PRODUCTO) = "" Then
         Exit For
      End If
               
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
         Q1 = "INSERT INTO DetDTE (IdDTE, IdEmpresa, IdProducto, TipoCod, CodProd, Producto, "
         Q1 = Q1 & " Descrip, UMedida, Cantidad, Precio, EsExento, IdImpAdic, IdImpAdicSII, TasaImpAdic, "
         Q1 = Q1 & " MontoImpAdic, PjeDescto, MontoDescto, SubTotal) "
         Q1 = Q1 & " VALUES (" & lIdDTE & ", " & gEmpresa.Id
         Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(i, C_IDPROD))
         Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_TIPOCOD)) & "'"
         Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_CODPROD)) & "'"
         Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_PRODUCTO)) & "'"
         Q1 = Q1 & ",'" & Left(ParaSQL(Grid.TextMatrix(i, C_DESCRIP)), MAX_LEN_DESCRIPPROD) & "'"
         Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_UMEDIDA)) & "'"
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_CANTIDAD)))
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_PRECIO)))
         Q1 = Q1 & "," & IIf(Trim(Grid.TextMatrix(i, C_ESEXENTO)) <> "", 1, 0)
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_IDIMPADIC))
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_CODIMPADICSII))
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)))
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC)))
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_PJEDESCTO)))
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_MONTODESCTO)))
         Q1 = Q1 & "," & Str0(vFmt(Grid.TextMatrix(i, C_SUBTOTAL))) & ")"
         Call ExecSQL(DbMain, Q1)
      
         lDTE.DetDTE(j).IdDetDTE = -1
         lDTE.DetDTE(j).IdDTE = lIdDTE
         
         j = j + 1
          
         If j >= MAX_ITEMDTE Then
            Exit For
         End If
      
      
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
         Q1 = "UPDATE DetDTE SET "
         Q1 = Q1 & "  IdProducto='" & Grid.TextMatrix(i, C_IDPROD) & "'"
         Q1 = Q1 & ", TipoCod='" & Grid.TextMatrix(i, C_TIPOCOD) & "'"
         Q1 = Q1 & ", CodProd='" & Grid.TextMatrix(i, C_CODPROD) & "'"
         Q1 = Q1 & ", Producto='" & ParaSQL(Grid.TextMatrix(i, C_PRODUCTO)) & "'"
         Q1 = Q1 & ", Descrip ='" & Left(ParaSQL(Grid.TextMatrix(i, C_DESCRIP)), MAX_LEN_DESCRIPPROD) & "'"
         Q1 = Q1 & ", UMedida='" & ParaSQL(Grid.TextMatrix(i, C_UMEDIDA)) & "'"
         Q1 = Q1 & ", Cantidad=" & Str0(vFmt(Grid.TextMatrix(i, C_CANTIDAD)))
         Q1 = Q1 & ", Precio=" & Str0(vFmt(Grid.TextMatrix(i, C_PRECIO)))
         Q1 = Q1 & ", EsExento=" & IIf(Trim(Grid.TextMatrix(i, C_ESEXENTO)) <> "", 1, 0)
         Q1 = Q1 & ", IdImpAdic=" & vFmt(Grid.TextMatrix(i, C_IDIMPADIC))
         Q1 = Q1 & ", CodImpAdicSII=" & Grid.TextMatrix(i, C_CODIMPADICSII)
         Q1 = Q1 & ", TasaImpAdic=" & str(vFmt(Grid.TextMatrix(i, C_TASAIMPADIC)))
         Q1 = Q1 & ", MontoImpAdic=" & str(vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC)))
         Q1 = Q1 & ", PjeDescto=" & str(vFmt(Grid.TextMatrix(i, C_PJEDESCTO)))
         Q1 = Q1 & ", MontoDescto=" & str(vFmt(Grid.TextMatrix(i, C_MONTODESCTO)))
         Q1 = Q1 & ", SubTotal=" & str(vFmt(Grid.TextMatrix(i, C_SUBTOTAL)))
         Q1 = Q1 & " WHERE IdDetDTE = " & Grid.TextMatrix(i, C_IDDETFACT) & " AND IdEmpresa = " & gEmpresa.Id
         Call ExecSQL(DbMain, Q1)
      
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
      
         Q1 = "DELETE * FROM DetDTE "
         Q1 = Q1 & " WHERE IdDetDTE=" & Grid.TextMatrix(i, C_IDDETFACT) & " AND IdEmpresa = " & gEmpresa.Id
         Call ExecSQL(DbMain, Q1)
      
      End If
            
   Next i

   
   'Referencias
   
   j = 0

   For i = Gr_Ref.FixedRows To Gr_Ref.rows - 1
      If Gr_Ref.TextMatrix(i, CR_TIPODOCREF) = "" Then
         Exit For
      End If
               
      If Gr_Ref.TextMatrix(i, CR_UPDATE) = FGR_I Then
         
         Q1 = "INSERT INTO Referencias (IdDTE, IdEmpresa, IdDTERef, IdTipoDocRef, CodDocRefSII, FolioRef, FechaRef, CodRefSII, RazonReferencia)"
         Q1 = Q1 & " VALUES (" & lIdDTE & ", " & gEmpresa.Id
         Q1 = Q1 & ", " & Val(Gr_Ref.TextMatrix(i, CR_IDDTEREF))
         Q1 = Q1 & ", " & vFmt(Gr_Ref.TextMatrix(i, CR_IDTIPODOCREF))
         Q1 = Q1 & ",'" & ParaSQL(Gr_Ref.TextMatrix(i, CR_CODDOCREFSII)) & "'"
         Q1 = Q1 & ",'" & ParaSQL(Gr_Ref.TextMatrix(i, CR_FOLIO)) & "'"
         Q1 = Q1 & ", " & Val(Gr_Ref.TextMatrix(i, CR_LNGFECHA))
         Q1 = Q1 & ", " & Val(Gr_Ref.TextMatrix(i, CR_CODREFSII))
         Q1 = Q1 & ",'" & ParaSQL(Gr_Ref.TextMatrix(i, CR_RAZONREF)) & "')"
         Call ExecSQL(DbMain, Q1)
      
      ElseIf Gr_Ref.TextMatrix(i, CR_UPDATE) = FGR_U Then
      
         If Val(Gr_Ref.TextMatrix(i, CR_IDREFERENCIA)) > 0 Then
         
            Q1 = "UPDATE Referencias SET "
            Q1 = Q1 & "  IdDTERef=" & Val(Gr_Ref.TextMatrix(i, CR_IDDTEREF))
            Q1 = Q1 & ", IdTipoDocRef=" & Val(Gr_Ref.TextMatrix(i, CR_IDTIPODOCREF))
            Q1 = Q1 & ", CodDocRefSII='" & Gr_Ref.TextMatrix(i, CR_CODDOCREFSII) & "'"
            Q1 = Q1 & ", FolioRef='" & Gr_Ref.TextMatrix(i, CR_FOLIO) & "'"
            Q1 = Q1 & ", FechaRef=" & Val(Gr_Ref.TextMatrix(i, CR_LNGFECHA))
            Q1 = Q1 & ", CodRefSII=" & vFmt(Gr_Ref.TextMatrix(i, CR_CODREFSII))
            Q1 = Q1 & ", RazonReferencia ='" & ParaSQL(Gr_Ref.TextMatrix(i, CR_RAZONREF)) & "'"
            Q1 = Q1 & " WHERE IdReferencia = " & Val(Gr_Ref.TextMatrix(i, CR_IDREFERENCIA)) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
            
         End If
      
      ElseIf Gr_Ref.TextMatrix(i, CR_UPDATE) = FGR_D Then
      
         If Val(Gr_Ref.TextMatrix(i, CR_IDREFERENCIA)) > 0 Then
         
            Q1 = "DELETE * FROM Referencias "
            Q1 = Q1 & " WHERE IdReferencia= " & Val(Gr_Ref.TextMatrix(i, CR_IDREFERENCIA)) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      End If
     
      lDTE.Referencia(j).IdReferencia = -1
      lDTE.Referencia(j).IdDTE = lIdDTE
      lDTE.Referencia(j).IdEmpresa = gEmpresa.Id
'      lDTE.Referencia(j).Ano = gEmpresa.Ano
      lDTE.Referencia(j).IdTipoDocRef = vFmt(Gr_Ref.TextMatrix(i, CR_IDTIPODOCREF))
      lDTE.Referencia(j).CodDocRefSII = Gr_Ref.TextMatrix(i, CR_CODDOCREFSII)
      lDTE.Referencia(j).FolioRef = Gr_Ref.TextMatrix(i, CR_FOLIO)
      lDTE.Referencia(j).FechaRef = Val(Gr_Ref.TextMatrix(i, CR_LNGFECHA))
      lDTE.Referencia(j).CodRefSII = vFmt(Gr_Ref.TextMatrix(i, CR_CODREFSII))
      lDTE.Referencia(j).RazonReferencia = Gr_Ref.TextMatrix(i, CR_RAZONREF)
      
      j = j + 1
       
      If j >= MAX_REFDTE Then
         Exit For
      End If
      
   Next i
     
   

End Sub

Private Sub GenDTEStruct()
   Dim i As Integer, j As Integer, k As Integer, n As Integer
   Dim Repetido As Boolean
   Dim TotDesIva As Long
      
   lDTE.IdDTE = 0
   
   lDTE.IdEmpresa = gEmpresa.Id
'   lDTE.Ano = gEmpresa.Ano
   lDTE.TipoDoc = lTipoDoc
   
   lDTE.TipoLib = LIB_VENTAS
   lDTE.CodDocSII = lCodDocSII
   
   If lEsGuiaDespacho Then
      lDTE.TipoDoc = TIPODOC_GUIADESPACHO
      lDTE.TipoLib = LIB_OTROS
      lDTE.CodDocSII = CODDOCDTESII_GUIADESPACHO
   End If
   
   lDTE.Folio = 0
   lDTE.idEstado = EDTE_ENVIADO
   lDTE.Fecha = GetTxDate(Tx_Fecha)
   lDTE.FechaVenc = GetTxDate(Tx_FechaVenc)
   lDTE.FormaDePago = IIf(CbItemData(Cb_FormaDePago) < 0, 0, CbItemData(Cb_FormaDePago))
   lDTE.IdEntidad = lIdEntidad
   lDTE.Rut = vFmtCID(Tx_RUT)
   lDTE.NotValidRut = IIf(Ch_Rut <> 0, False, True)
   lDTE.RazonSocial = Trim(Tx_RazonSocial)
   lDTE.Giro = Trim(Tx_Giro)
   lDTE.Direccion = Trim(Tx_Direccion)
   lDTE.Comuna = Trim(Tx_Comuna)
   lDTE.Ciudad = Trim(Tx_Ciudad)
   lDTE.Contacto = Trim(Tx_Contacto)
   lDTE.SubTotal = vFmt(Tx_SubTotal)
   lDTE.PjeDestoGlobal = vFmt(Tx_PjeDescto)
   lDTE.DesctoGlobal = vFmt(Tx_MontoDescto)
'   If lSoloExento Then
'      lDTE.Exento = vFmt(Tx_Neto)
'   Else
'      lDTE.Neto = vFmt(Tx_Neto)
'   End If
   lDTE.Exento = lTotExento
   lDTE.Neto = lTotAfecto
   lDTE.TasaIVA = vFmt(Lb_TasaIVA)
   lDTE.Iva = vFmt(Tx_IVA)
   lDTE.ImpAdicional = vFmt(Tx_ImpAdic)
   lDTE.Total = vFmt(Tx_Total)
   lDTE.MailReceptor = Trim(Tx_MailReceptor)
   lDTE.Observaciones = Me.Tx_ObsDTE.Text
   
   lDTE.TipoDespacho = 0
   lDTE.Traslado = 0
   If lEsGuiaDespacho Then
      lDTE.TipoDespacho = IIf(CbItemData(Cb_TipoDespacho) < 0, 0, CbItemData(Cb_TipoDespacho))
      lDTE.Traslado = IIf(CbItemData(Cb_Traslado) < 0, 0, CbItemData(Cb_Traslado))
   End If
   
   lDTE.DetFormaPago = IIf(CbItemData(Cb_DetFormaPago) < 0, 0, CbItemData(Cb_DetFormaPago))
   lDTE.TextDetFormaPago = IIf(CbItemData(Cb_DetFormaPago) < 0, "", cbItemText(Cb_DetFormaPago, CbItemData(Cb_DetFormaPago)))
   lDTE.Vendedor = IIf(CbItemData(Cb_Vendedor) < 0, 0, CbItemData(Cb_Vendedor))
   lDTE.TextVendedor = IIf(CbItemData(Cb_Vendedor) < 0, "", cbItemText(Cb_Vendedor, CbItemData(Cb_Vendedor)))
   
   lDTE.EsExport = lEsExport
   lDTE.EsGuiaDesp = lEsGuiaDespacho
      
   'limpiamos los arreglos
   For j = 0 To MAX_ITEMDTE
      lDTE.DetDTE(j).IdDetDTE = 0
      lDTE.DetDTE(j).IdDTE = 0
      lDTE.DetDTE(j).IdEmpresa = 0
      lDTE.DetDTE(j).IdProducto = 0
      lDTE.DetDTE(j).TipoCod = ""
      lDTE.DetDTE(j).CodProd = ""
      lDTE.DetDTE(j).Producto = ""
      lDTE.DetDTE(j).Descrip = ""
      lDTE.DetDTE(j).UMedida = ""
      lDTE.DetDTE(j).Cantidad = 0
      lDTE.DetDTE(j).Precio = 0
      lDTE.DetDTE(j).EsExento = 0
      lDTE.DetDTE(j).IdImpAdic = 0
      lDTE.DetDTE(j).CodImpAdicSII = ""
      lDTE.DetDTE(j).TasaImpAdic = 0
      lDTE.DetDTE(j).MontoImpAdic = 0
      lDTE.DetDTE(j).DescImpAdic = ""
      lDTE.DetDTE(j).PjeDescto = 0
      lDTE.DetDTE(j).MontoDescto = 0
      lDTE.DetDTE(j).SubTotal = 0
   Next j
   
   For i = 0 To MAX_IMPADICDTE
      lDTE.ImpAdic(i).IdImpAdic = 0
      lDTE.ImpAdic(i).IdImpAdicSII = 0
      lDTE.ImpAdic(i).TasaImpAdic = 0
      lDTE.ImpAdic(i).MontoImpAdic = 0
      lDTE.ImpAdic(i).NetoImpAdic = 0
      lDTE.ImpAdic(i).DescImpAdic = ""
   Next i
   
   For j = 0 To MAX_REFDTE
      lDTE.Referencia(j).IdTipoDocRef = 0
      lDTE.Referencia(j).IdDTE = 0
      lDTE.Referencia(j).IdEmpresa = 0
'      lDTE.Referencia(j).Ano = gEmpresa.Ano
      lDTE.Referencia(j).IdTipoDocRef = 0
      lDTE.Referencia(j).CodDocRefSII = ""
      lDTE.Referencia(j).FolioRef = ""
      lDTE.Referencia(j).FechaRef = 0
      lDTE.Referencia(j).CodRefSII = 0
      lDTE.Referencia(j).RazonReferencia = ""
   Next j
   
   j = 0
   k = 0
   
   'ahora los llenamos
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_PRODUCTO) = "" Then
         Exit For
      End If
      
      If Grid.RowHeight(i) > 0 Then    ' no está borrado
                    
         lDTE.DetDTE(j).IdDetDTE = -1
         lDTE.DetDTE(j).IdDTE = lIdDTE
         lDTE.DetDTE(j).IdEmpresa = gEmpresa.Id
   '      lDTE.DetDTE(j).Ano = gEmpresa.Ano
         lDTE.DetDTE(j).IdProducto = vFmt(Grid.TextMatrix(i, C_IDPROD))
         lDTE.DetDTE(j).TipoCod = Grid.TextMatrix(i, C_TIPOCOD)
         lDTE.DetDTE(j).CodProd = Grid.TextMatrix(i, C_CODPROD)
         lDTE.DetDTE(j).Producto = Grid.TextMatrix(i, C_PRODUCTO)
         lDTE.DetDTE(j).Descrip = Left(Grid.TextMatrix(i, C_DESCRIP), MAX_LEN_DESCRIPPROD)
         lDTE.DetDTE(j).UMedida = Grid.TextMatrix(i, C_UMEDIDA)
         lDTE.DetDTE(j).Cantidad = vFmt(Grid.TextMatrix(i, C_CANTIDAD))
         lDTE.DetDTE(j).Precio = vFmt(Grid.TextMatrix(i, C_PRECIO))
         lDTE.DetDTE(j).EsExento = IIf(Trim(Grid.TextMatrix(i, C_ESEXENTO)) <> "", True, False)
         lDTE.DetDTE(j).IdImpAdic = vFmt(Grid.TextMatrix(i, C_IDIMPADIC))
         lDTE.DetDTE(j).CodImpAdicSII = Grid.TextMatrix(i, C_CODIMPADICSII)
         lDTE.DetDTE(j).TasaImpAdic = vFmt(Grid.TextMatrix(i, C_TASAIMPADIC))
         lDTE.DetDTE(j).MontoImpAdic = vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC))
         lDTE.DetDTE(j).DescImpAdic = IIf(lDTE.DetDTE(j).IdImpAdic <> 8, ParaSQL(Grid.TextMatrix(i, C_IMPADIC)), "Credito 65% Empresa Constructora")
         lDTE.DetDTE(j).PjeDescto = vFmt(Grid.TextMatrix(i, C_PJEDESCTO))
         lDTE.DetDTE(j).MontoDescto = vFmt(Grid.TextMatrix(i, C_MONTODESCTO))
         lDTE.DetDTE(j).SubTotal = vFmt(Grid.TextMatrix(i, C_SUBTOTAL))
         
         If lDTE.DetDTE(j).IdImpAdic > 0 Then 'And lDTE.DetDTE(j).IdImpAdic <> 8 Then
         
            'primero buscamos si ya hay un impuesto igual a este para no repetirlo
            Repetido = False
            For n = 0 To k - 1
               If lDTE.ImpAdic(n).IdImpAdic = lDTE.DetDTE(j).IdImpAdic Then   'hay uno igual, sumamos los montos
                  'lDTE.ImpAdic(n).MontoImpAdic = lDTE.ImpAdic(n).MontoImpAdic + vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC))
                  lDTE.ImpAdic(n).MontoImpAdic = lDTE.ImpAdic(n).MontoImpAdic + IIf(lDTE.DetDTE(j).IdImpAdic <> 8, vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC)), vFmt(Grid.TextMatrix(i, C_DESCIVA)))
                  TotDesIva = lDTE.ImpAdic(n).MontoImpAdic
                  Repetido = True
                  Exit For
               End If
            Next n
                  
            If Not Repetido Then          'no hay, es nuevo, lo agregamos
               lDTE.ImpAdic(k).IdImpAdic = lDTE.DetDTE(j).IdImpAdic
               lDTE.ImpAdic(k).IdImpAdicSII = Val(lDTE.DetDTE(j).CodImpAdicSII)
               lDTE.ImpAdic(k).TasaImpAdic = lDTE.DetDTE(j).TasaImpAdic
               'lDTE.ImpAdic(k).MontoImpAdic = vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC))
               lDTE.ImpAdic(k).MontoImpAdic = IIf(lDTE.DetDTE(j).IdImpAdic <> 8, vFmt(Grid.TextMatrix(i, C_MONTOIMPADIC)), vFmt(Grid.TextMatrix(i, C_DESCIVA)))
               lDTE.ImpAdic(k).NetoImpAdic = lDTE.DetDTE(j).SubTotal
               lDTE.ImpAdic(k).DescImpAdic = lDTE.DetDTE(j).DescImpAdic
               k = k + 1
               If k > MAX_IMPADICDTE Then
                  Exit For
               End If
            End If
         End If
            
         j = j + 1
          
         If j >= MAX_ITEMDTE Then
            Exit For
         End If
      End If
      
   Next i
   
   'referencias
   j = 0

   For i = Gr_Ref.FixedRows To Gr_Ref.rows - 1
   
      If Gr_Ref.TextMatrix(i, CR_TIPODOCREF) = "" Then
         Exit For
      End If
                    
      lDTE.Referencia(j).IdReferencia = -1
      lDTE.Referencia(j).IdDTE = lIdDTE
      lDTE.Referencia(j).IdEmpresa = gEmpresa.Id
'      lDTE.Referencia(j).Ano = gEmpresa.Ano
      lDTE.Referencia(j).IdTipoDocRef = vFmt(Gr_Ref.TextMatrix(i, CR_IDTIPODOCREF))
      lDTE.Referencia(j).CodDocRefSII = Gr_Ref.TextMatrix(i, CR_CODDOCREFSII)
      lDTE.Referencia(j).FolioRef = Gr_Ref.TextMatrix(i, CR_FOLIO)
      lDTE.Referencia(j).FechaRef = vFmt(Gr_Ref.TextMatrix(i, CR_LNGFECHA))
      lDTE.Referencia(j).CodRefSII = vFmt(Gr_Ref.TextMatrix(i, CR_CODREFSII))
      lDTE.Referencia(j).RazonReferencia = Gr_Ref.TextMatrix(i, CR_RAZONREF)
      
      j = j + 1
       
      If j >= MAX_REFDTE Then
         Exit For
      End If
      
   Next i


End Sub

Private Sub FillDTERef()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Idx As Integer
   Dim CbIdxImpAdic As Long
   
   If lIdDTERef = 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT RUT, Folio, Fecha, CodDocSII, PjeDesctoGlobal, TipoDocRef.IdTipoDocRef, TipoDocRef.CodDocRefSII, TipoDocRef.Nombre as NTipoDocRef, TipoDoc, TipoLib "
   Q1 = Q1 & " FROM DTE INNER JOIN TipoDocRef ON DTE.CodDocSII = TipoDocRef.CodDocRefSII "
   Q1 = Q1 & " WHERE IdDTE = " & lIdDTERef & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Tx_RUT = FmtCID(vFld(Rs("RUT")))
      Call FillDataRUT
    
      Tx_PjeDescto = Format(vFld(Rs("PjeDesctoGlobal")), DBLFMT2)
    
      Gr_Ref.rows = Gr_Ref.FixedRows
      i = Gr_Ref.rows
      Gr_Ref.rows = Gr_Ref.rows + 1
      Gr_Ref.TextMatrix(i, CR_NUMLIN) = i - Gr_Ref.FixedRows + 1
      Gr_Ref.TextMatrix(i, CR_IDDTEREF) = lIdDTERef
      Gr_Ref.TextMatrix(i, CR_IDTIPODOCREF) = vFld(Rs("IdTipoDocRef"))
      Gr_Ref.TextMatrix(i, CR_TIPODOCREF) = vFld(Rs("NTipoDocRef"))
      Gr_Ref.TextMatrix(i, CR_CODDOCREFSII) = vFld(Rs("CodDocRefSII"))
      Gr_Ref.TextMatrix(i, CR_FOLIO) = vFld(Rs("Folio"))
      Gr_Ref.TextMatrix(i, CR_FECHA) = Format(vFld(Rs("Fecha")), SDATEFMT)
      Gr_Ref.TextMatrix(i, CR_LNGFECHA) = vFld(Rs("Fecha"))
      Gr_Ref.TextMatrix(i, CR_CODREFSII) = lTipoRef
      Gr_Ref.TextMatrix(i, CR_REFSII) = gTipoRefSII(lTipoRef)
      Gr_Ref.TextMatrix(i, CR_UPDATE) = FGR_I
      
   End If
   
   Idx = GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))

   Call CloseRs(Rs)
   Call FGrVRows(Gr_Ref, 1)
   
   lSoloExento = IIf(gTipoDoc(Idx).TieneExento And Not gTipoDoc(Idx).TieneAfecto, True, False)
   lTieneExento = IIf(gTipoDoc(Idx).TieneExento, True, False)
   
   If lTieneExento = True And Not lSoloExento Then
      Grid.ColWidth(C_ESEXENTO) = 300
   Else
      Grid.ColWidth(C_ESEXENTO) = 0
      Grid.TextMatrix(0, C_ESEXENTO) = ""
   End If

   Q1 = "SELECT TipoCod, CodProd, Producto, Descrip, UMedida, Cantidad, Precio, IdImpAdic, IdImpAdicSII, TasaImpAdic, PjeDescto, MontoDescto, SubTotal, EsExento "
   Q1 = Q1 & " FROM DetDTE WHERE IdDTE = " & lIdDTERef & " AND IdEmpresa = " & gEmpresa.Id
   Q1 = Q1 & " ORDER BY IdDetDTE"

   Set Rs = OpenRs(DbMain, Q1)
     
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1
      Grid.TextMatrix(i, C_TIPOCOD) = vFld(Rs("TipoCod"))
      Grid.TextMatrix(i, C_CODPROD) = vFld(Rs("CodProd"))
      Grid.TextMatrix(i, C_PRODUCTO) = vFld(Rs("Producto"))
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"))
      Grid.TextMatrix(i, C_UMEDIDA) = vFld(Rs("UMedida"))
      Grid.TextMatrix(i, C_CANTIDAD) = Format(vFld(Rs("Cantidad")), DBLFMT2)
      Grid.TextMatrix(i, C_PRECIO) = Format(vFld(Rs("Precio")), IIf(Ch_DecEnPrecio <> 0, lFmtPrecio2, lFmtPrecio))
      Grid.TextMatrix(i, C_ESEXENTO) = IIf(vFld(Rs("EsExento")) <> 0, "x", "")
      Grid.TextMatrix(i, C_IDIMPADIC) = vFld(Rs("IdImpAdic"))
      
      CbIdxImpAdic = lcbImpAdic.FindItem(vFld(Rs("IdImpAdic")))
      Grid.TextMatrix(i, C_IMPADIC) = lcbImpAdic.List(CbIdxImpAdic)
      
      Grid.TextMatrix(i, C_CODIMPADICSII) = vFld(Rs("IdImpAdicSII"))
      Grid.TextMatrix(i, C_TASAIMPADIC) = Format(vFld(Rs("TasaImpAdic")), DBLFMT2)
      Grid.TextMatrix(i, C_PJEDESCTO) = Format(vFld(Rs("PjeDescto")), DBLFMT2)
      Grid.TextMatrix(i, C_MONTODESCTO) = Format(vFld(Rs("MontoDescto")), NUMFMT)
      Grid.TextMatrix(i, C_SUBTOTAL) = Format(vFld(Rs("SubTotal")), lFmtPrecio)
      Grid.TextMatrix(i, C_UPDATE) = FGR_I
      
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   Call FGrVRows(Grid, 1)
   
   Call CalcTotal
End Sub

Private Sub LoadDTE(ByVal IdDTE As Long)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Idx As Integer
   Dim CbIdxImpAdic As Long
   
   If IdDTE = 0 Then
      Q1 = "SELECT IdDTE FROM DTE "
      Q1 = Q1 & " WHERE TipoLib IN ( " & LIB_VENTAS & "," & LIB_OTROS & ")" & " AND CodDocSII = '" & lCodDocSII & "' AND IdEstado IN( " & EDTE_ENVIADO & "," & EDTE_PROCESADO & "," & EDTE_EMITIDO & ")"
      Q1 = Q1 & " ORDER BY IdDTE desc"
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdDTE = vFld(Rs(0))
      Else
         MsgBox1 "No se encontró un DTE de este tipo en estado Enviado, Procesado o Emitido para obtener los datos." & vbCrLf & vbCrLf & "Utilice la opción" & vbCrLf & vbCrLf & "  'Ingresar DTE basado en uno emitido previamente'" & vbCrLf & vbCrLf & "para seleccionar un DTE directamente.", vbOKOnly + vbExclamation
         Call CloseRs(Rs)
         Exit Sub
      End If
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT RUT, CodDocSII, TipoDoc, TipoLib, Contacto, SubTotal, PjeDesctoGlobal, TipoDespacho, Traslado, FormaDePago, ObsDTE"
   Q1 = Q1 & " FROM DTE "
   Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Tx_RUT = FmtCID(vFld(Rs("RUT")))
      Call FillDataRUT
      
      Tx_Contacto = vFld(Rs("Contacto"))
      
      If lEsGuiaDespacho Then
         Call CbSelItem(Cb_TipoDespacho, vFld(Rs("TipoDespacho")))
         Call CbSelItem(Cb_Traslado, vFld(Rs("Traslado")))
      End If
      
      Call CbSelItem(Cb_FormaDePago, vFld(Rs("FormaDePago")))
      
      Tx_PjeDescto = Format(vFld(Rs("PjeDesctoGlobal")), DBLFMT2)
      Tx_ObsDTE = vFld(Rs("ObsDTE"))
            
   End If
   
   Idx = GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))

   Call CloseRs(Rs)
   
   'cargamos detalle de productos
   Q1 = "SELECT TipoCod, CodProd, Producto, Descrip, UMedida, Cantidad, Precio, IdImpAdic, IdImpAdicSII, TasaImpAdic, MontoImpAdic, PjeDescto, MontoDescto, SubTotal, EsExento "
   Q1 = Q1 & " FROM DetDTE WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Q1 = Q1 & " ORDER BY IdDetDTE"

   Set Rs = OpenRs(DbMain, Q1)
     
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1
      Grid.TextMatrix(i, C_TIPOCOD) = vFld(Rs("TipoCod"))
      Grid.TextMatrix(i, C_CODPROD) = vFld(Rs("CodProd"))
      Grid.TextMatrix(i, C_PRODUCTO) = vFld(Rs("Producto"))
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"))
      Grid.TextMatrix(i, C_UMEDIDA) = vFld(Rs("UMedida"))
      Grid.TextMatrix(i, C_CANTIDAD) = Format(vFld(Rs("Cantidad")), DBLFMT2)
      Grid.TextMatrix(i, C_PRECIO) = Format(vFld(Rs("Precio")), IIf(Ch_DecEnPrecio <> 0, lFmtPrecio2, lFmtPrecio))
      Grid.TextMatrix(i, C_ESEXENTO) = IIf(vFld(Rs("EsExento")) <> 0, "x", "")
      Grid.TextMatrix(i, C_IDIMPADIC) = vFld(Rs("IdImpAdic"))
      
      CbIdxImpAdic = lcbImpAdic.FindItem(vFld(Rs("IdImpAdic")))
      Grid.TextMatrix(i, C_IMPADIC) = lcbImpAdic.List(CbIdxImpAdic)
      Grid.TextMatrix(i, C_TASAIMPADIC) = Format(vFld(Rs("TasaImpAdic")), DBLFMT2)
      
      Grid.TextMatrix(i, C_TASAEDITABLE) = IIf(Val(lcbImpAdic.Matrix(IA_TASA, CbIdxImpAdic)) = 0, 1, 0) ' 19 nov 2019 - pam: se agrega Val() porque venía ""
      If Val(Grid.TextMatrix(i, C_TASAEDITABLE)) <> 0 Then  'es editable
         Call FGrForeColor(Grid, i, C_TASAIMPADIC, vbBlue)
      Else
         Call FGrForeColor(Grid, i, C_TASAIMPADIC, vbBlack)
      End If

      Grid.TextMatrix(i, C_MONTOIMPADIC) = Format(vFld(Rs("MontoImpAdic")), NUMFMT)
      Grid.TextMatrix(i, C_CODIMPADICSII) = vFld(Rs("IdImpAdicSII"))
      Grid.TextMatrix(i, C_PJEDESCTO) = Format(vFld(Rs("PjeDescto")), DBLFMT2)
      Grid.TextMatrix(i, C_MONTODESCTO) = Format(vFld(Rs("MontoDescto")), NUMFMT)
      Grid.TextMatrix(i, C_SUBTOTAL) = Format(vFld(Rs("SubTotal")), lFmtPrecio)
      Grid.TextMatrix(i, C_UPDATE) = FGR_I
      
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   Call FGrVRows(Grid, 1)
   
   Call CalcTotal
   
   'cargamos las referencias
   Q1 = "SELECT Referencias.IdTipoDocRef, Referencias.CodDocRefSII, Nombre as NTipoDocRef, FolioRef, FechaRef, CodRefSII, RazonReferencia "
   Q1 = Q1 & " FROM Referencias INNER JOIN TipoDocRef ON Referencias.IdTipoDocRef = TipoDocRef.IdTipoDocRef "
   Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   Gr_Ref.rows = Gr_Ref.FixedRows
   i = Gr_Ref.rows
   Do While Not Rs.EOF
    
      Gr_Ref.rows = Gr_Ref.rows + 1
      Gr_Ref.TextMatrix(i, CR_NUMLIN) = i - Gr_Ref.FixedRows + 1
      Gr_Ref.TextMatrix(i, CR_IDDTEREF) = ""
      Gr_Ref.TextMatrix(i, CR_IDTIPODOCREF) = vFld(Rs("IdTipoDocRef"))
      Gr_Ref.TextMatrix(i, CR_TIPODOCREF) = vFld(Rs("NTipoDocRef"))
      Gr_Ref.TextMatrix(i, CR_CODDOCREFSII) = vFld(Rs("CodDocRefSII"))
      Gr_Ref.TextMatrix(i, CR_FOLIO) = vFld(Rs("FolioRef"))
      Gr_Ref.TextMatrix(i, CR_FECHA) = Format(vFld(Rs("FechaRef")), SDATEFMT)
      Gr_Ref.TextMatrix(i, CR_LNGFECHA) = vFld(Rs("FechaRef"))
      Gr_Ref.TextMatrix(i, CR_CODREFSII) = vFld(Rs("CodRefSII"))
      Gr_Ref.TextMatrix(i, CR_REFSII) = gTipoRefSII(vFld(Rs("CodRefSII")))
      Gr_Ref.TextMatrix(i, CR_RAZONREF) = vFld(Rs("RazonReferencia"))
      Gr_Ref.TextMatrix(i, CR_UPDATE) = ""

      i = i + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Gr_Ref, 1)
   
   'Cargamos los datos de factura de exportación si corresponde
   If lEsExport Then
      
      Q1 = "SELECT CodIndServicio, CodPais, CodPuertoEmbarque, CodPuertoDesembarque, CodMoneda, TipoCambioPesos, CodModVenta, CodViaTransporte, TotalBultos, TotalClauVenta"
      Q1 = Q1 & " FROM DTEFactExp "
      Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         lDTE.FactExp.CodIndServicio = vFld(Rs("CodIndServicio"))
         lDTE.FactExp.CodPais = vFld(Rs("CodPais"))
         lDTE.FactExp.CodPuertoEmbarque = vFld(Rs("CodPuertoEmbarque"))
         lDTE.FactExp.CodPuertoDesembarque = vFld(Rs("CodPuertoDesembarque"))
         lDTE.FactExp.CodMoneda = vFld(Rs("CodMoneda"))
         lDTE.FactExp.TipoCambio = vFld(Rs("TipoCambioPesos"))
         lDTE.FactExp.CodModVenta = vFld(Rs("CodModVenta"))
         lDTE.FactExp.CodViaTransporte = vFld(Rs("CodViaTransporte"))
         lDTE.FactExp.TotalBultos = vFld(Rs("TotalBultos"))
         lDTE.FactExp.TotClausulaVenta = vFld(Rs("TotalClauVenta"))

      End If
      
      Call CloseRs(Rs)
      
   End If
   
   'Cargamos los datos de guia de despacho  si corresponde
   If lEsGuiaDespacho Then
      
      Q1 = "SELECT Patente, RutChofer, NombreChofer "
      Q1 = Q1 & " FROM DTEGuiaDesp "
      Q1 = Q1 & " WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         lDTE.GuiaDesp.Patente = vFld(Rs("Patente"))
         lDTE.GuiaDesp.RutChofer = vFld(Rs("RutChofer"))
         lDTE.GuiaDesp.NombreChofer = vFld(Rs("NombreChofer"))

      End If
      
      Call CloseRs(Rs)
      
   End If
   
End Sub

Private Sub SelEntDefExport()
   Dim Rs As Recordset
   Dim Q1 As String
   
   If Not lEsExport Then
      Exit Sub
   End If
   
   Tx_RUT = FmtCID(RUT_DEFEXPORT)
   Ch_Rut = 1
   
         
   Q1 = "SELECT IdEntidad, Nombre, Regiones.Comuna, Ciudad, Giro, EMail "
   Q1 = Q1 & " FROM Entidades LEFT JOIN Regiones ON Entidades.Comuna = Regiones.Id "
   Q1 = Q1 & " WHERE RUT = '" & RUT_DEFEXPORT & "' AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      lIdEntidad = vFld(Rs("IdEntidad"))
      Tx_RazonSocial = vFld(Rs("Nombre"))
'      Tx_Comuna = vFld(Rs("Comuna"))     'Debe ir en blanco
      Tx_Ciudad = vFld(Rs("Ciudad"))
      Tx_Giro = vFld(Rs("Giro"))
      Tx_MailReceptor = vFld(Rs("EMail"))
   End If
   
   Call CloseRs(Rs)

   lEntCompleta = True
   
End Sub
