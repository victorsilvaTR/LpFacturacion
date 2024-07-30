VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmEntidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Entidad"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   Icon            =   "FrmEntidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   300
      Picture         =   "FrmEntidad.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   690
      TabIndex        =   49
      Top             =   360
      Width           =   690
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   1320
      TabIndex        =   29
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Datos Básicos"
      TabPicture(0)   =   "FrmEntidad.frx":0687
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(21)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(22)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Tx_Fax"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Tx_Tel"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Tx_Ciudad"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Tx_Dir"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Tx_Nombre"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Tx_RUT"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Cb_ComPostal"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Cb_Comuna"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Cb_Region"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Tx_Giro"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Tx_DomPostal"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Bt_Web"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Tx_Web"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Tx_EMail"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Bt_Email"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Frame1(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Frame1(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Tx_Codigo"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Frame2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Ch_Rut"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Cb_EsSupermercado"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "Contactos"
      TabPicture(1)   =   "FrmEntidad.frx":06A3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   6915
         Left            =   -74520
         TabIndex        =   51
         Top             =   660
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   12197
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
      Begin VB.ComboBox Cb_EsSupermercado 
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   780
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   600
         Width           =   195
      End
      Begin VB.Frame Frame2 
         Caption         =   "Observaciones"
         ForeColor       =   &H00FF0000&
         Height          =   1515
         Left            =   300
         TabIndex        =   46
         Top             =   6180
         Width           =   8115
         Begin VB.TextBox Tx_Obs 
            Height          =   1035
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   300
            Width           =   7815
         End
      End
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         ForeColor       =   &H00FF0000&
         Height          =   555
         Index           =   1
         Left            =   5220
         TabIndex        =   44
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton Op_Estado 
            Caption         =   "Bloqueado"
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   4
            Top             =   240
            Width           =   1155
         End
         Begin VB.OptionButton Op_Estado 
            Caption         =   "Inactivo"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Op_Estado 
            Caption         =   "Activo"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clasificación"
         ForeColor       =   &H00FF0000&
         Height          =   915
         Index           =   0
         Left            =   300
         TabIndex        =   43
         Top             =   5160
         Width           =   8115
         Begin VB.CheckBox Ch_Clas 
            Caption         =   "Otro"
            Height          =   195
            Index           =   5
            Left            =   6420
            TabIndex        =   25
            Top             =   600
            Width           =   1155
         End
         Begin VB.CheckBox Ch_Clas 
            Caption         =   "Distribuidor"
            Height          =   195
            Index           =   4
            Left            =   6420
            TabIndex        =   24
            Top             =   240
            Width           =   1155
         End
         Begin VB.CheckBox Ch_Clas 
            Caption         =   "Socio"
            Height          =   195
            Index           =   3
            Left            =   3120
            TabIndex        =   23
            Top             =   600
            Width           =   1155
         End
         Begin VB.CheckBox Ch_Clas 
            Caption         =   "Empleado"
            Height          =   195
            Index           =   2
            Left            =   3120
            TabIndex        =   22
            Top             =   240
            Width           =   1155
         End
         Begin VB.CheckBox Ch_Clas 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   240
            Width           =   1155
         End
         Begin VB.CheckBox Ch_Clas 
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   600
            Width           =   1155
         End
      End
      Begin VB.CommandButton Bt_Email 
         Height          =   375
         Left            =   3900
         Picture         =   "FrmEntidad.frx":06BF
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Tx_EMail 
         Height          =   315
         Left            =   300
         MaxLength       =   100
         TabIndex        =   16
         ToolTipText     =   "Se puede indicar más de un mail separándolo por "";"""
         Top             =   4680
         Width           =   3615
      End
      Begin VB.TextBox Tx_Web 
         Height          =   315
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   18
         Top             =   4680
         Width           =   3735
      End
      Begin VB.CommandButton Bt_Web 
         Height          =   375
         Left            =   8040
         Picture         =   "FrmEntidad.frx":0ACA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4620
         Width           =   375
      End
      Begin VB.TextBox Tx_DomPostal 
         Height          =   315
         Left            =   300
         MaxLength       =   35
         TabIndex        =   14
         Top             =   4080
         Width           =   5415
      End
      Begin VB.TextBox Tx_Giro 
         Height          =   315
         Left            =   300
         MaxLength       =   80
         TabIndex        =   12
         Top             =   3480
         Width           =   5415
      End
      Begin VB.ComboBox Cb_Region 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox Cb_Comuna 
         Height          =   315
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2280
         Width           =   2775
      End
      Begin VB.ComboBox Cb_ComPostal 
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox Tx_RUT 
         Height          =   315
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   0
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   300
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1140
         Width           =   8115
      End
      Begin VB.TextBox Tx_Dir 
         Height          =   315
         Left            =   300
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1740
         Width           =   8115
      End
      Begin VB.TextBox Tx_Ciudad 
         Height          =   315
         Left            =   5760
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Tx_Tel 
         Height          =   315
         Left            =   300
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2880
         Width           =   5415
      End
      Begin VB.TextBox Tx_Fax 
         Height          =   315
         Left            =   5760
         MaxLength       =   15
         TabIndex        =   11
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Es Supermercado o Comercio similar:"
         Height          =   195
         Index           =   8
         Left            =   5760
         TabIndex        =   50
         Top             =   3300
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   360
         TabIndex        =   47
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Corto:"
         Height          =   195
         Index           =   6
         Left            =   2640
         TabIndex        =   45
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail receptor de Doc. Electrónico:"
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   42
         Top             =   4500
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sitio Web:"
         Height          =   195
         Index           =   22
         Left            =   4320
         TabIndex        =   41
         Top             =   4500
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio Postal:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   40
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comuna Postal:"
         Height          =   195
         Index           =   21
         Left            =   5760
         TabIndex        =   39
         Top             =   3900
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Giro:"
         Height          =   195
         Index           =   12
         Left            =   300
         TabIndex        =   38
         Top             =   3300
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT o Ref.:"
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   37
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre o Razón Social:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   36
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   35
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comuna:"
         Height          =   195
         Index           =   4
         Left            =   2940
         TabIndex        =   34
         Top             =   2100
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   5
         Left            =   5760
         TabIndex        =   33
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Región:"
         Height          =   195
         Index           =   7
         Left            =   300
         TabIndex        =   32
         Top             =   2100
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos:"
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   31
         Top             =   2700
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Index           =   10
         Left            =   5760
         TabIndex        =   30
         Top             =   2700
         Width           =   300
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   10440
      TabIndex        =   27
      Top             =   360
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   10440
      TabIndex        =   28
      Top             =   780
      Width           =   1155
   End
End
Attribute VB_Name = "FrmEntidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NOMBRE = 0
Const C_FONO = 1
Const C_CARGO = 2
Const C_ID = 3
Const C_ESTADO = 4

Const NCOLS = C_ESTADO

Dim lRc As Integer
Dim Oper As Integer
Dim lEntidad As Entidad_t

Private Sub bt_Cancel_Click()
   lRc = vbCancel
   Unload Me

End Sub

Private Sub Bt_Email_Click()
   Dim Buf As String
   Dim Rc As Long
   Dim Pos As Integer
   
   Pos = InStr(Tx_EMail, "@")
   If Trim(Tx_EMail) <> "" And Trim(Tx_Nombre) <> "" And Pos <> 0 Then
     Buf = "mailto:" & Trim(Tx_Nombre) & "<" & Trim(Tx_EMail) & ">"
     Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
     
   End If
   
End Sub
Private Sub bt_OK_Click()
      
   If Valida() = False Then
      Exit Sub
   End If
   
   lRc = vbOK
   
   Call SaveAll
   
   Unload Me
   
End Sub

Friend Function FView(Entidad As Entidad_t) As Integer
   Oper = O_VIEW
   lEntidad = Entidad
   Me.Show vbModal
   
   FView = lRc
   Entidad = lEntidad
   
End Function
Friend Function FEdit(Entidad As Entidad_t) As Integer
   Oper = O_EDIT
   
   lEntidad = Entidad
   Me.Show vbModal
   
   FEdit = lRc
   Entidad = lEntidad
   
End Function

Friend Function FNew(Entidad As Entidad_t, Optional ByVal Rut As String = "", Optional ByVal Nombre As String = "") As Integer
   Oper = O_NEW
      
   lEntidad.Id = 0
   lEntidad.Clasif = Entidad.Clasif
   lEntidad.Rut = Rut
   lEntidad.Nombre = Nombre
   
   Me.Show vbModal
   
   FNew = lRc
   Entidad = lEntidad
   
End Function

Private Sub Bt_Web_Click()
   Dim Rc As Long
   
   If Trim(Tx_Web) <> "" Then
      Rc = ShellExecute(Me.hWnd, "open", Tx_Web, "", "", 1)
   End If
   
End Sub

Private Sub Cb_Comuna_Click()

   Call SelItem(Cb_ComPostal, ItemData(Cb_Comuna))
   
End Sub

Private Sub Cb_Region_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Cod As String
   
   Cod = Right("00" & ItemData(Cb_Region), 2)
   
   Q1 = "SELECT Comuna, id FROM Regiones"
   Q1 = Q1 & " WHERE Codigo='" & Cod & "'"
   Q1 = Q1 & " ORDER BY Comuna"
   Cb_Comuna.Clear
   Cb_Comuna.AddItem "<Ninguna>"
   Cb_Comuna.ItemData(Cb_Comuna.NewIndex) = 0
   Call FillComboSinRepetir(Cb_Comuna, DbMain, Q1, -1)
   
End Sub

Private Sub Form_Activate()
   If Oper = O_EDIT Then
      Tx_Codigo.SetFocus
   End If
End Sub

Private Sub Form_Load()

   lRc = vbCancel
   SSTab1.Tab = 0
   Call FillCombosFrm
   Ch_Rut = 1

   Op_Estado(EE_ACTIVO).Value = True
   
   Call SetUpGrid
   
   If Oper = O_VIEW Then
      Call EnableForm(Me, False)
'   Else
'      Call EnableForm(Me, gEmpresa.FCierre = 0)
   End If
   
   If Oper = O_NEW Then
      Caption = "Nueva Entidad"
      
      If lEntidad.Clasif >= 0 Then
         Ch_Clas(lEntidad.Clasif).Value = 1
      End If
      
   ElseIf Oper = O_EDIT Then
      Caption = "Modificar Entidad"
     ' Call SettxRO(Tx_Rut, True)
     ' Ch_Rut.Enabled = False
      
   Else
      Caption = "Ver Entidad"
      
   End If
   
   Call LoadAll
   
   Call SetupPriv
   
End Sub
Private Sub FillCombosFrm()
   Dim Q1 As String
   Dim Rs As Recordset
      
   'COMBO REGION
   Call FillRegion(Cb_Region)
   
   Cb_Region.ListIndex = 0
   
   'COMUNA POSTAL, SE MUESTRAN TODAS LAS COMUNAS QU EXISTEN
   Q1 = "SELECT Comuna,id FROM Regiones"
   Q1 = Q1 & " ORDER BY Comuna"
   Cb_ComPostal.AddItem "< Ninguna >"
   Cb_ComPostal.ItemData(Cb_ComPostal.NewIndex) = 0
   Call FillCombo(Cb_ComPostal, DbMain, Q1, -1)
   
   'Es Supermercado
   Call CbAddItem(Cb_EsSupermercado, "No", 0)
   Call CbAddItem(Cb_EsSupermercado, "Si", 1)
   Call CbSelItem(Cb_EsSupermercado, 0)
   
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   Grid.FixedCols = 0
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_NOMBRE) = 3200
   Grid.ColWidth(C_FONO) = 1500
   Grid.ColWidth(C_CARGO) = 2650
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ESTADO) = 0
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
      Grid.ColAlignment(i) = flexAlignLeftCenter
      
   Next i
   
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_FONO) = "Teléfono"
   Grid.TextMatrix(0, C_CARGO) = "Cargo"
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   If lEntidad.Id <= 0 Then
      Tx_RUT = lEntidad.Rut
      Tx_Nombre = lEntidad.Nombre
      
      Exit Sub
   End If
   
   'DATOS BASICOS Y OBSERVACION
   Q1 = "SELECT Rut, NotValidRut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad,"
   Q1 = Q1 & "Telefonos, Fax, Giro, EsSupermercado, DomPostal, ComPostal, EMail,"
   Q1 = Q1 & "Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5"
   Q1 = Q1 & " FROM Entidades "
   Q1 = Q1 & " WHERE idEntidad=" & lEntidad.Id & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Ch_Rut.Value = IIf(vFld(Rs("NotValidRut")) <> 0, 0, 1)
      Tx_RUT = FmtCID(vFld(Rs("Rut")), Ch_Rut <> 0)
      Tx_Codigo = vFld(Rs("Codigo"))
      Tx_Nombre = vFld(Rs("Nombre"), True)
      Tx_Dir = vFld(Rs("Direccion"), True)
      Tx_Ciudad = vFld(Rs("Ciudad"), True)
      Tx_Tel = vFld(Rs("Telefonos"), True)
      Tx_Fax = vFld(Rs("Fax"), True)
      Tx_DomPostal = vFld(Rs("DomPostal"), True)
      Tx_EMail = vFld(Rs("Email"), True)
      Tx_Web = vFld(Rs("Web"), True)
      Tx_Giro = vFld(Rs("Giro"), True)
      Call CbSelItem(Cb_EsSupermercado, vFld(Rs("EsSupermercado")))
      Op_Estado(vFld(Rs("Estado"))).Value = True
      Tx_Obs = vFld(Rs("Obs"), True)
      
      For i = ENT_CLIENTE To ENT_OTRO
         Ch_Clas(i).Value = vFld(Rs("Clasif" & i))
      Next i
      
      Call SelItem(Cb_Region, vFld(Rs("Region")))
      Call SelItem(Cb_Comuna, vFld(Rs("Comuna")))
      Call SelItem(Cb_ComPostal, vFld(Rs("ComPostal")))
   End If
   Call CloseRs(Rs)
   
   'CONTACTOS
   Q1 = "SELECT Nombre,Telefono,Cargo,idContacto FROM Contactos"
   Q1 = Q1 & " WHERE idEntidad=" & lEntidad.Id & " AND IdEmpresa = " & gEmpresa.Id
   Q1 = Q1 & " ORDER BY Nombre"
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   Grid.rows = i
   Do While Rs.EOF = False
      Grid.rows = i + 1
        
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
      Grid.TextMatrix(i, C_CARGO) = vFld(Rs("Cargo"), True)
      Grid.TextMatrix(i, C_FONO) = vFld(Rs("Telefono"), True)
      Grid.TextMatrix(i, C_ID) = vFld(Rs("idContacto"), True)
      
      Rs.MoveNext
      i = i + 1
      
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   Call FGrModRow(Grid, Row, FGR_U, C_ID, C_ESTADO)
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG3_EdType)
   If Col = C_NOMBRE And Grid.TextMatrix(Row - 1, C_NOMBRE) = "" Then
      Exit Sub
   End If
   
   If Col <> C_NOMBRE And Grid.TextMatrix(Row, C_NOMBRE) = "" Then
      Exit Sub
   End If
   
   EdType = FEG_Edit
   
   If Grid.rows = Row + 1 Then
      Grid.rows = Grid.rows + 1
      Grid.FlxGrid.TopRow = Row
      
   End If
   
End Sub

Private Sub Tx_CodActEcon_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_Codigo_KeyPress(KeyAscii As Integer)
   Call KeyUpper(KeyAscii)
End Sub
Private Sub Tx_RUT_KeyPress(KeyAscii As Integer)
   If Ch_Rut <> 0 Then
      Call KeyCID(KeyAscii)
   
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   
   End If
End Sub

Private Sub Tx_RUT_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim AuxRut As String
   
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
'   If Ch_Rut.Value <> 0 And vFmtCID(Tx_RUT) = 0 Then     'FCA (3 jun 2009) para que no borre RUT si se desea indicar que no debe validar RUT con Checkbox adyacente
'      Tx_RUT = ""
'      Tx_RUT.SetFocus
'      Exit Sub
'   End If
      
   Q1 = "SELECT IdEntidad,Rut FROM Entidades WHERE Rut = '" & vFmtCID(Tx_RUT, Ch_Rut <> 0) & "'" & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   lEntidad.Id = 0
   If Rs.EOF = False Then   'ya existe
      lEntidad.Id = vFld(Rs(0))
      
      If (Oper = O_NEW) Or (Oper = O_EDIT And vFmtCID(Tx_RUT, Ch_Rut <> 0) <> vFmtCID(lEntidad.Rut, Ch_Rut <> 0)) Then
         MsgBox1 "¡ADVERTENCIA!" & vbNewLine & " Ha ingresado un RUT que ya existe y no es el RUT con el que estaba trabajando inicialmente, sólo podrá consultar sus datos y no grabar.", vbExclamation
      End If
      
      Call LoadAll
   ElseIf lEntidad.Id = 0 Then
      Call ClearAll
   End If
      
   Call CloseRs(Rs)
   
   If Ch_Rut <> 0 Then
      AuxRut = FmtCID(vFmtCID(Tx_RUT))
      If AuxRut <> "0-0" Then
         Tx_RUT = AuxRut
      End If
   End If
   
'   If lEntidad.id = 0 Then         'FCA (3 jun 2009) ya no existe la tabla NContrib de HR
'      Call GetEntidadFromNContrib
'   End If
   
   
End Sub
Private Function Valida() As Boolean
   Dim ChCont As Byte
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
      
   Valida = False
   
   If Tx_RUT = "" Or Trim(Tx_RUT) = "0-0" Then
      MsgBox1 "Debe ingresar RUT.", vbExclamation
      Tx_RUT.SetFocus
      Exit Function
   End If
   
   If Oper = O_EDIT And vFmtCID(lEntidad.Rut) <> vFmtCID(Tx_RUT) Then
      If MsgBox1("¡ATENCION!" & vbNewLine & " Ha modificado el RUT " & lEntidad.Rut & " de la entidad. ¿Desea continuar?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
         Exit Function
      End If
   End If
   
   If Not MsgValidCID(Tx_RUT, Ch_Rut <> 0) Then
      Tx_RUT.SetFocus
      Exit Function
   End If
   
'   If Tx_Codigo = "" Then
'      MsgBox1 "Debe ingresar CODIGO.", vbExclamation
'      Tx_Codigo.SetFocus
'      Exit Function
'   End If
'
   If Trim(Tx_Nombre) = "" Then
      MsgBox1 "Debe ingresar Nombre o Razón Social.", vbExclamation
      Tx_Codigo.SetFocus
      Exit Function
   End If
   
   For i = ENT_CLIENTE To ENT_OTRO
      If Ch_Clas(i).Value = 1 Then
         ChCont = ChCont + 1
      End If
   Next i
   
   If ChCont = 0 Then
      MsgBox1 "Debe darle una clasificación a la entidad.", vbExclamation
      Ch_Clas(ENT_CLIENTE).SetFocus
      Exit Function
   End If
   
   If (Oper = O_NEW) Or (Oper = O_EDIT And vFmtCID(lEntidad.Rut) <> vFmtCID(Tx_RUT)) Then
      Q1 = "SELECT Rut, Codigo FROM Entidades WHERE Rut='" & vFmtCID(Tx_RUT, Ch_Rut <> 0) & "' OR Codigo='" & Trim(Tx_Codigo) & "'" & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         MsgBox1 "Rut o Nombre Corto de esta entidad ya existe.", vbExclamation
         Tx_Codigo.SetFocus
         Call CloseRs(Rs)
         Exit Function
      End If
      Call CloseRs(Rs)
      
      Q1 = "SELECT Rut FROM Entidades WHERE Nombre='" & ParaSQL(Tx_Nombre) & "'" & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         MsgBox1 "El Nombre de esta entidad ya existe.", vbExclamation
         Tx_Codigo.SetFocus
         Call CloseRs(Rs)
         Exit Function
      End If
      Call CloseRs(Rs)
   
      
   End If
   
   If gConectData.Proveedor = PROV_ACEPTA And Trim(Tx_EMail) = "" Then
      MsgBox1 "Debe ingresar el E-Mail del receptor del Documento Electrónico.", vbExclamation
      Tx_EMail.SetFocus
      Exit Function
   End If
   
   If Not ValidEmail(Trim(Tx_EMail)) Then
      MsgBox1 "E-Mail del receptor del Documento Electrónico en inválido.", vbExclamation
      Tx_EMail.SetFocus
      Exit Function
   End If
   
   Valida = True
   
End Function
Private Sub SaveGrid()
   Dim i As Integer
   Dim Q1 As String
   
   For i = 1 To Grid.rows - 1
      If Grid.TextMatrix(i, C_NOMBRE) <> "" Then
         If Grid.TextMatrix(i, C_ESTADO) = FGR_I Then
            Q1 = "INSERT INTO Contactos (idEntidad,Nombre,Cargo,Telefono,IdEmpresa)"
            Q1 = Q1 & " VALUES (" & lEntidad.Id
            Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_CARGO)) & "'"
            Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_FONO)) & "'"
            Q1 = Q1 & ",'" & gEmpresa.Id & "')"
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(i, C_ESTADO) = FGR_U Then
            Q1 = "UPDATE Contactos SET Nombre='" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ", Cargo='" & ParaSQL(Grid.TextMatrix(i, C_CARGO)) & "'"
            Q1 = Q1 & ", Telefono='" & ParaSQL(Grid.TextMatrix(i, C_FONO)) & "'"
            Q1 = Q1 & " WHERE idEntidad=" & lEntidad.Id & " AND IdEmpresa = " & gEmpresa.Id
            Q1 = Q1 & " AND idContacto=" & Grid.TextMatrix(i, C_ID)
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      End If
      
   Next i
   
End Sub
Private Sub SaveAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Estado As Byte
   Dim i As Integer
   Dim OriClasif As Boolean
   Dim Codigo As String
   
   For i = EE_ACTIVO To EE_BLOQUEADO
      If Op_Estado(i).Value = True Then
         Estado = i
         Exit For
      End If
   Next i
   
   'verificamos que el usuario no haya eliminado la clasificación que viene desde la invocación, sino que sólo haya agregado una clasificación
   'si la eliminó, mandamos un mensaje de advertencia para el caso del New
   
   If lEntidad.Clasif <> SIN_CLASLST Then    'viene con clasificación en la invocación
   
      For i = ENT_CLIENTE To ENT_OTRO
         If Ch_Clas(i).Value = 1 And i = lEntidad.Clasif Then
            OriClasif = True
            Exit For
         End If
      Next i
         
   End If
   
   If Oper = O_NEW And lEntidad.Id = 0 Then
   
      If lEntidad.Clasif <> SIN_CLASLST And OriClasif = False Then  'venía una clasificación en la invocación y no está la clasificación que venía desde la invocación
         MsgBox1 "¡ADVERTENCIA!, la nueva entidad se dejó en la(s) clasificación(es) que usted le asignó.", vbExclamation
         lEntidad.Clasif = SIN_CLASLST
      End If
   
      Set Rs = DbMain.OpenRecordset("Entidades", dbOpenTable)
      Rs.AddNew
      
      lEntidad.Id = Rs("idEntidad")
      
      Rs("NotValidRut") = (Ch_Rut = 0)
      Rs("RUT") = vFmtCID(Tx_RUT, Ch_Rut <> 0)
      Rs("IdEmpresa") = gEmpresa.Id
      
      Rs.Update
      Rs.Close
   
   Else
      lRc = vbRetry  ' ya existe
      
   End If
   
   Codigo = IIf(Trim(Tx_Codigo) = "", vFmtRut(Tx_RUT), Tx_Codigo)
      
   Q1 = "UPDATE Entidades SET "
   Q1 = Q1 & "  Nombre='" & ParaSQL(Tx_Nombre) & "'"
   Q1 = Q1 & ", Codigo='" & ParaSQL(Codigo) & "'"
   Q1 = Q1 & ", Direccion='" & ParaSQL(Tx_Dir) & "'"
   Q1 = Q1 & ", Region=" & CbItemData(Cb_Region)
   Q1 = Q1 & ", Comuna=" & CbItemData(Cb_Comuna)
   Q1 = Q1 & ", Ciudad='" & ParaSQL(Tx_Ciudad) & "'"
   Q1 = Q1 & ", Telefonos='" & ParaSQL(Tx_Tel) & "'"
   Q1 = Q1 & ", Fax='" & ParaSQL(Tx_Fax) & "'"
   Q1 = Q1 & ", Giro='" & ParaSQL(Tx_Giro) & "'"
   Q1 = Q1 & ", EsSupermercado = " & CbItemData(Cb_EsSupermercado)
   Q1 = Q1 & ", EMail='" & ParaSQL(Tx_EMail) & "'"
   Q1 = Q1 & ", Web='" & ParaSQL(Tx_Web) & "'"
   Q1 = Q1 & ", Estado=" & Estado
   Q1 = Q1 & ", Obs='" & ParaSQL(Tx_Obs) & "'"
   Q1 = Q1 & ", DomPostal='" & ParaSQL(Tx_DomPostal) & "'"
   Q1 = Q1 & ", ComPostal=" & CbItemData(Cb_ComPostal)
   
   For i = ENT_CLIENTE To ENT_OTRO
      Q1 = Q1 & ",Clasif" & i & "=" & Ch_Clas(i).Value
      
      If lEntidad.Clasif = SIN_CLASLST And Ch_Clas(i).Value <> 0 Then
         lEntidad.Clasif = i
      End If

   Next i
   
   If Oper = O_EDIT Then
      Q1 = Q1 & ", NotValidRut=" & CInt(Ch_Rut = 0)
      Q1 = Q1 & ", RUT='" & vFmtCID(Tx_RUT, Ch_Rut <> 0) & "'"
   End If
   
   Q1 = Q1 & " WHERE idEntidad=" & lEntidad.Id & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
      
   Call SaveGrid
   
  ' lEntidad.Nombre = Tx_Nombre
   lEntidad.Rut = Tx_RUT
   lEntidad.NotValidRut = CInt(Ch_Rut = 0)
   lEntidad.Estado = Estado
   lEntidad.Nombre = Tx_Nombre
   lEntidad.Codigo = Codigo

End Sub

Private Sub ClearAll()
   Dim i As Integer

   Tx_Nombre = ""
   Tx_Dir = ""
   Tx_Ciudad = ""
   Tx_Tel = ""
   Tx_Fax = ""
   Tx_DomPostal = ""
   Tx_EMail = ""
   Tx_Web = ""
   Tx_Codigo = ""
   Op_Estado(0).Value = True
   Tx_Obs = ""
   
   For i = ENT_CLIENTE To ENT_OTRO
      Ch_Clas(i).Value = 0
   Next i
   
   If lEntidad.Clasif >= 0 Then
      Ch_Clas(lEntidad.Clasif).Value = 1
   End If
   
   Call SelItem(Cb_Region, -1)
   Call SelItem(Cb_Comuna, -1)
   Call SelItem(Cb_ComPostal, -1)
   
   Grid.rows = Grid.FixedRows
   Call FGrVRows(Grid)

End Sub
Private Function SetupPriv()
   
'   If Not ChkPriv(PRV_ADM_DEF) Then
'      Call EnableForm(Me, False)
'   End If
   
End Function
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If Trim(Tx_RUT) = "0-0" Then
      MsgBox1 "RUT Inválido.", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RUT, Ch_Rut <> 0) Then
      Cancel = True
      Exit Sub
   End If
End Sub
