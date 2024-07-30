VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFmtImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato de Importación de Productos"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "FrmFmtImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_VerEjemplo 
      Caption         =   "Ver ejemplo...."
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   9120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   2115
      Left            =   240
      TabIndex        =   6
      Top             =   3900
      Width           =   10215
      Begin VB.Label Label1 
         Caption         =   "El formato del archivo es posicional, por lo que se deben incluir  TODOS los campos, aunque vayan en blanco. "
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   8955
      End
      Begin VB.Label Label3 
         Caption         =   "NOTAS:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lb_NotaImp 
         Caption         =   $"FrmFmtImport.frx":000C
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   8955
      End
      Begin VB.Label Label1 
         Caption         =   "* Indica los campos que deben tener un valor válido (distinto de blanco)"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   1620
         Width           =   8955
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   10275
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8880
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "FrmFmtImport.frx":0120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1260
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4471
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   1380
      TabIndex        =   5
      Top             =   7140
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Columnas o campos del archivo:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2475
   End
End
Attribute VB_Name = "FrmFmtImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const C_CAMPO = 0
Const C_FORMATO = 1

Dim lFmtArray() As FmtImp_t
Dim lFmtCaption As String
Dim lEjemplo As Boolean
Dim lNulos As Boolean


Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)
End Sub

Private Sub Form_Load()
   
   Me.Caption = lFmtCaption
   
   
   Call SetUpGrid
   Call LoadGrid
   
   If lEjemplo Then
      Bt_VerEjemplo.Visible = True
      Me.Height = Me.Height + Bt_VerEjemplo.Height + 200
   End If


End Sub

Private Sub SetUpGrid()

   Call FGrSetup(Grid)

   Grid.ColWidth(C_CAMPO) = 2400
   Grid.ColWidth(C_FORMATO) = 7400
   
   Grid.ColAlignment(C_CAMPO) = flexAlignLeftCenter
   Grid.ColAlignment(C_FORMATO) = flexAlignLeftCenter
   
   
   Grid.TextMatrix(0, C_CAMPO) = "Campo de Información"
   Grid.TextMatrix(0, C_FORMATO) = "Formato"
   
End Sub

Private Sub LoadGrid()
   Dim i As Integer
   Dim j As Integer

   Grid.rows = Grid.FixedRows
   i = Grid.rows - 1
   
   For j = 0 To UBound(lFmtArray)
      Grid.rows = Grid.rows + 1
      i = i + 1
      Grid.TextMatrix(i, C_CAMPO) = lFmtArray(j).Campo
      Grid.TextMatrix(i, C_FORMATO) = lFmtArray(j).Formato
   Next j
   
   Call FGrVRows(Grid)
End Sub

Friend Sub FView(ByVal FmtCaption As String, FmtArray() As FmtImp_t)

   lFmtCaption = FmtCaption
   lFmtArray = FmtArray
   Me.Show vbModal

End Sub
Public Sub FViewProducto()

   Call FillProducto
   Me.Show vbModal

End Sub
Public Sub FViewEntidad()

   Call FillEntidad
   Me.Show vbModal

End Sub

Private Sub FillProducto()
   Dim i As Integer

   lFmtCaption = "Formato Importación de Productos"
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Tipo Código"
   lFmtArray(i).Formato = "Texto largo 10 (letras no acentuadas y números)"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código"
   lFmtArray(i).Formato = "Texto largo 35 (letras no acentuadas y números)"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Producto *"
   lFmtArray(i).Formato = "Texto largo 80"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "UMedida"
   lFmtArray(i).Formato = "Texto largo 4 (letras no acentuadas y números)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Precio *"
   lFmtArray(i).Formato = "Valor numérico sin puntos, con coma para decimales"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Producto *"
   lFmtArray(i).Formato = "Valor Si o No"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Observaciones"
   lFmtArray(i).Formato = "Texto largo 255"


End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Me.Caption)
   End If
      
End Sub

Private Sub FillEntidad()
   Dim i As Integer

   lFmtCaption = "Formato Importación de Entidades"
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "RUT *"
   lFmtArray(i).Formato = "Con puntos (opcionalmente) y dígito verificador"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Codigo *"
   lFmtArray(i).Formato = "Nombre corto de la entidad, en mayúscula y sin blancos, largo 15"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre o Razón Social *"
   lFmtArray(i).Formato = "Texto largo 80"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Dirección"
   lFmtArray(i).Formato = "Texto largo 100"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Comuna"
   lFmtArray(i).Formato = "Texto largo 20"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Ciudad"
   lFmtArray(i).Formato = "Texto largo 20"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Teléfonos"
   lFmtArray(i).Formato = "Texto largo 30"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fax"
   lFmtArray(i).Formato = "Texto largo 15"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Giro"
   lFmtArray(i).Formato = "Texto largo 50"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Domicilio Postal"
   lFmtArray(i).Formato = "Texto largo 35"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Comuna Postal"
   lFmtArray(i).Formato = "Texto largo 20"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Email"
   lFmtArray(i).Formato = "Texto largo 50"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Sitio Web"
   lFmtArray(i).Formato = "Texto largo 50 "
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Observaciones"
   lFmtArray(i).Formato = "Texto largo 255"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Cliente"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Cliente"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Proveedor"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Proveedor"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Empleado"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Empleado"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Socio"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Socio"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Distribuidor"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Distribuidor"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Otro"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es Otros"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Supermercado"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es Supermercado"


End Sub

