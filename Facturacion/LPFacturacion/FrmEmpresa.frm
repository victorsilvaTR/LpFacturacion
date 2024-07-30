VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Empresa"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13800
   Icon            =   "FrmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   10260
      TabIndex        =   32
      Top             =   420
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   10260
      TabIndex        =   33
      Top             =   780
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   1260
      TabIndex        =   34
      Top             =   420
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Antecedentes Empresa"
      TabPicture(0)   =   "FrmEmpresa.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(22)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(21)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(20)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(19)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(18)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(17)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(13)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(16)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(28)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Bt_Web"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Tx_Web"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Cb_ComPostal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Cb_Comuna"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Cb_Region"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Tx_Nombre"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Tx_ApMaterno"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Tx_Dpto"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Tx_Numero"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Tx_RUT"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Tx_RazonSocial"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Tx_DirPostal"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Tx_Calle"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Tx_Ciudad"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Tx_Telefonos"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Tx_Fax"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "tx_NombreCorto"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Tx_EMail"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Bt_Email"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Tx_ObsDTE"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Antecedentes  Legales"
      TabPicture(1)   =   "FrmEmpresa.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Im_Exc(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Im_Exc(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(27)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(26)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(25)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(8)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lb_MsgCodAct"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "La_Url"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame4"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Fr_TrBolsa"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Fr_LibCaja"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Bt_FInicioAct"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Bt_FConstit"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Tx_FInicioAct"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Tx_FConstit"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Tx_Giro"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Frame2(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Cb_ActEcon"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Tx_CodActEcon"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      Begin VB.TextBox Tx_ObsDTE 
         Height          =   315
         Left            =   180
         MaxLength       =   100
         TabIndex        =   17
         Top             =   5880
         Width           =   8055
      End
      Begin VB.CommandButton Bt_Email 
         Height          =   375
         Left            =   3720
         Picture         =   "FrmEmpresa.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5220
         Width           =   375
      End
      Begin VB.TextBox Tx_EMail 
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   5220
         Width           =   3555
      End
      Begin VB.TextBox tx_NombreCorto 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2475
      End
      Begin VB.TextBox Tx_Fax 
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox Tx_Telefonos 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   4080
         Width           =   5535
      End
      Begin VB.TextBox Tx_Ciudad 
         Height          =   315
         Left            =   5760
         TabIndex        =   8
         Top             =   3540
         Width           =   2535
      End
      Begin VB.TextBox Tx_Calle 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   3000
         Width           =   5715
      End
      Begin VB.TextBox Tx_DirPostal 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   4620
         Width           =   5535
      End
      Begin VB.TextBox Tx_RazonSocial 
         Height          =   315
         Left            =   180
         MaxLength       =   200
         TabIndex        =   0
         Top             =   1920
         Width           =   8115
      End
      Begin VB.TextBox Tx_RUT 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1395
      End
      Begin VB.TextBox Tx_Numero 
         Height          =   315
         Left            =   5940
         TabIndex        =   4
         Top             =   3000
         Width           =   1155
      End
      Begin VB.TextBox Tx_Dpto 
         Height          =   315
         Left            =   7140
         TabIndex        =   5
         Top             =   3000
         Width           =   1155
      End
      Begin VB.TextBox Tx_ApMaterno 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   2460
         Width           =   4035
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   4260
         TabIndex        =   2
         Top             =   2460
         Width           =   4035
      End
      Begin VB.ComboBox Cb_Region 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3540
         Width           =   2775
      End
      Begin VB.ComboBox Cb_Comuna 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3540
         Width           =   2715
      End
      Begin VB.ComboBox Cb_ComPostal 
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4620
         Width           =   2535
      End
      Begin VB.TextBox Tx_Web 
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   5220
         Width           =   3675
      End
      Begin VB.CommandButton Bt_Web 
         Height          =   375
         Left            =   7860
         Picture         =   "FrmEmpresa.frx":044F
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5220
         Width           =   375
      End
      Begin VB.TextBox Tx_CodActEcon 
         Height          =   315
         Left            =   -68040
         MaxLength       =   6
         TabIndex        =   24
         Top             =   2280
         Width           =   1275
      End
      Begin VB.ComboBox Cb_ActEcon 
         Height          =   315
         Left            =   -74820
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2280
         Width           =   6735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Contador"
         ForeColor       =   &H00FF0000&
         Height          =   975
         Index           =   2
         Left            =   -74820
         TabIndex        =   61
         Top             =   2760
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox Tx_RutContador 
            Height          =   315
            Left            =   180
            TabIndex        =   25
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox Tx_Contador 
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   480
            Width           =   6420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RUT:"
            Height          =   195
            Index           =   23
            Left            =   180
            TabIndex        =   63
            Top             =   290
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   24
            Left            =   1440
            TabIndex        =   62
            Top             =   290
            Width           =   600
         End
      End
      Begin VB.TextBox Tx_Giro 
         Height          =   315
         Left            =   -74820
         MaxLength       =   80
         TabIndex        =   18
         Top             =   1080
         Width           =   8040
      End
      Begin VB.TextBox Tx_FConstit 
         Height          =   315
         Left            =   -72300
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Tx_FInicioAct 
         Height          =   315
         Left            =   -68040
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         Caption         =   "Representantes Legales"
         ForeColor       =   &H00FF0000&
         Height          =   1755
         Left            =   -74820
         TabIndex        =   56
         Top             =   3840
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox Tx_RUTRep 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox Tx_NombreRep 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   28
            Top             =   600
            Width           =   6420
         End
         Begin VB.TextBox Tx_NombreRep 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   31
            Top             =   1200
            Width           =   6420
         End
         Begin VB.TextBox Tx_RUTRep 
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   30
            Top             =   1200
            Width           =   1155
         End
         Begin VB.CheckBox Ch_RepConjunta 
            Caption         =   "Representación conjunta"
            Height          =   255
            Left            =   5760
            TabIndex        =   29
            Top             =   300
            Width           =   2115
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RUT:"
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   60
            Top             =   405
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   12
            Left            =   1440
            TabIndex        =   59
            Top             =   405
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   14
            Left            =   1440
            TabIndex        =   58
            Top             =   1005
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RUT:"
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   57
            Top             =   1005
            Width           =   390
         End
      End
      Begin VB.CommandButton Bt_FConstit 
         Height          =   315
         Left            =   -71280
         Picture         =   "FrmEmpresa.frx":0759
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton Bt_FInicioAct 
         Height          =   315
         Left            =   -67020
         Picture         =   "FrmEmpresa.frx":0A63
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Frame Fr_LibCaja 
         Caption         =   "Libro de Caja"
         Height          =   1275
         Left            =   -70560
         TabIndex        =   53
         Top             =   9180
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CheckBox Ch_ObligaLibComprasVentas 
            Caption         =   "Se encuentra obligado a llevar"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Libro Compras Ventas según la        Ley de IVA"
            Height          =   435
            Left            =   540
            TabIndex        =   55
            Top             =   600
            Width           =   2775
         End
      End
      Begin VB.Frame Fr_TrBolsa 
         Caption         =   "Transa en la Bolsa"
         Height          =   1275
         Left            =   -74580
         TabIndex        =   50
         Top             =   9180
         Visible         =   0   'False
         Width           =   3495
         Begin VB.OptionButton Op_TrBolsaSi 
            Caption         =   "Si"
            Height          =   255
            Left            =   720
            TabIndex        =   52
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton Op_TrBolsaNo 
            Caption         =   "No"
            Height          =   255
            Left            =   1980
            TabIndex        =   51
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Franquicias Tributarias"
         Height          =   3015
         Left            =   -70560
         TabIndex        =   42
         Top             =   6000
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "Régimen Artículo 14 bis"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   3015
         End
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "Ley 18.392 / 19.149"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   48
            Top             =   720
            Width           =   3015
         End
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "D. L. 600"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   3015
         End
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "D. L. 701"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   46
            Top             =   1440
            Width           =   3015
         End
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "D. S. 341"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   3015
         End
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "Régimen Artículo 14 ter"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   44
            Top             =   2160
            Width           =   3015
         End
         Begin VB.CheckBox Ch_Franquicia 
            Caption         =   "Régimen Artículo 14 quater"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   43
            Top             =   2520
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Contribuyente"
         Height          =   3015
         Left            =   -74280
         TabIndex        =   35
         Top             =   6000
         Visible         =   0   'False
         Width           =   3495
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad Empresarial"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   92
            Top             =   2520
            Width           =   2895
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad Anónima Abierta"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad Anónima Cerrada"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad por Acción"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   2895
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Soc. Personas 1ª Categoría"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   2895
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Empresario Individual (EIRL)"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   37
            Top             =   1800
            Width           =   2895
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Empresario Individual"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   36
            Top             =   2160
            Width           =   2895
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones pie de documento electrónico:"
         Height          =   195
         Index           =   28
         Left            =   180
         TabIndex        =   91
         Top             =   5640
         Width           =   3255
      End
      Begin VB.Label La_Url 
         Caption         =   "www.sii.cl/catastro/codigos_economica.htm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   -70020
         MouseIcon       =   "FrmEmpresa.frx":0D6D
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   5700
         Width           =   3315
      End
      Begin VB.Label lb_MsgCodAct 
         Caption         =   "Código Act. Economica descontinuado, verifique su código en"
         Height          =   255
         Left            =   -74520
         TabIndex        =   89
         Top             =   5700
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail:"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   88
         Top             =   5030
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre corto:"
         Height          =   195
         Index           =   13
         Left            =   1800
         TabIndex        =   87
         Top             =   885
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Index           =   10
         Left            =   5760
         TabIndex        =   86
         Top             =   3885
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos:"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   85
         Top             =   3890
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Región:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   84
         Top             =   3345
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio Postal:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   83
         Top             =   4430
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   5
         Left            =   5760
         TabIndex        =   82
         Top             =   3345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comuna:"
         Height          =   195
         Index           =   4
         Left            =   3000
         TabIndex        =   81
         Top             =   3360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calle:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   80
         Top             =   2810
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social/Apellido Paterno:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   79
         Top             =   1730
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   78
         Top             =   885
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   17
         Left            =   5940
         TabIndex        =   77
         Top             =   2810
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Of. Dpto:"
         Height          =   195
         Index           =   18
         Left            =   7140
         TabIndex        =   76
         Top             =   2810
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno:"
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   75
         Top             =   2270
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
         Height          =   195
         Index           =   20
         Left            =   4260
         TabIndex        =   74
         Top             =   2270
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comuna Postal:"
         Height          =   195
         Index           =   21
         Left            =   5760
         TabIndex        =   73
         Top             =   4425
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sitio Web:"
         Height          =   195
         Index           =   22
         Left            =   4200
         TabIndex        =   72
         Top             =   5030
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód. act. ecón.:"
         Height          =   195
         Index           =   8
         Left            =   -68040
         TabIndex        =   71
         Top             =   2090
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clasificador de Actividades Económicas:"
         Height          =   195
         Index           =   2
         Left            =   -74820
         TabIndex        =   70
         Top             =   2090
         Width           =   2865
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Giro:"
         Height          =   195
         Index           =   25
         Left            =   -74820
         TabIndex        =   69
         Top             =   885
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio Actividades:"
         Height          =   195
         Index           =   26
         Left            =   -69900
         TabIndex        =   68
         Top             =   1620
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Constitución de la Empresa:"
         Height          =   195
         Index           =   27
         Left            =   -74820
         TabIndex        =   67
         Top             =   1620
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Image Im_Exc 
         Height          =   330
         Index           =   0
         Left            =   -66720
         Picture         =   "FrmEmpresa.frx":0EBF
         Top             =   2280
         Width           =   300
      End
      Begin VB.Image Im_Exc 
         Height          =   330
         Index           =   1
         Left            =   -74820
         Picture         =   "FrmEmpresa.frx":1289
         Top             =   5640
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(*) Comuna descontinuada por el SII"
         Height          =   195
         Left            =   180
         TabIndex        =   66
         Top             =   6300
         Width           =   2550
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   570
      Index           =   0
      Left            =   300
      Picture         =   "FrmEmpresa.frx":1653
      Top             =   420
      Width           =   570
   End
End
Attribute VB_Name = "FrmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const VERSION_1 = 1

Dim lRc As Integer
Dim lOper As Integer
Dim lcbCodActiv As ClsCombo
Dim lInLoad As Boolean

Private Sub bt_Cancel_Click()
   lRc = vbCancel
   Unload Me

End Sub

Private Sub Bt_Email_Click()
   Dim Buf As String
   Dim Rc As Long
   Dim Pos As Integer
   
   Pos = InStr(Tx_EMail, "@")
   If Trim(Tx_EMail) <> "" And Trim(Tx_RazonSocial) <> "" And Pos <> 0 Then
     Buf = "mailto:" & Trim(Tx_RazonSocial) & "<" & Trim(Tx_EMail) & ">"
     Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
     
   End If
   
End Sub

Private Sub Bt_FConstit_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FConstit)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_FInicioAct_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FInicioAct)
   Set Frm = Nothing
   
End Sub

Private Sub bt_OK_Click()
   
   If Valida() = True Then
      Call SaveAll
      Unload Me
   End If
   
End Sub
Public Function FView(ByVal IdEmpresa As Long) As Integer
   lOper = O_VIEW
   Me.Show vbModal
   FView = lRc
End Function
Public Function FEdit(ByVal IdEmpresa As Long) As Integer
   lOper = O_EDIT
   Me.Show vbModal
   FEdit = lRc
End Function

Private Sub Bt_Web_Click()
   Dim Rc As Long
   
   If Trim(Tx_Web) <> "" Then
      Rc = ShellExecute(Me.hWnd, "open", Tx_Web, "", "", 1)
   End If
   
End Sub

Private Sub Cb_ActEcon_Click()
   'Tx_CodActEcon = Right("000000" & ItemData(Cb_ActEcon), 5)
   
   'PS se cambio códgo de Actividad
   lb_MsgCodAct = IIf(Val(lcbCodActiv.Matrix(2)) = VERSION_1, "Código de Actividad Económica descontinuado, verifique su código en ", "")
   Im_Exc(0).Visible = IIf(Val(lcbCodActiv.Matrix(2)) = VERSION_1, True, False)
   Im_Exc(1).Visible = Im_Exc(0).Visible
   La_Url.Visible = Im_Exc(0).Visible
   
   Tx_CodActEcon = Right("000000" & lcbCodActiv.ItemData, 6)
   
End Sub

Private Sub Cb_Comuna_Click()
   Call SelItem(Cb_ComPostal, ItemData(Cb_Comuna))
End Sub

Private Sub Cb_Region_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Cod As String
   
   Cod = Right("00" & ItemData(Cb_Region), 2)
   
   Q1 = "SELECT Comuna,id FROM Regiones"
   Q1 = Q1 & " WHERE Codigo='" & Cod & "'"
   Q1 = Q1 & " ORDER BY Comuna"
   Cb_Comuna.Clear
   Cb_Comuna.AddItem "<Ninguna>"
   Cb_Comuna.ItemData(Cb_Comuna.NewIndex) = 0
   Call FillComboSinRepetir(Cb_Comuna, DbMain, Q1, -1)
   
End Sub

Private Sub Ch_Franquicia_Click(Index As Integer)
   
   If lInLoad Then
      Exit Sub
   End If
   
   If Index = FRANQ_14TER Then
      If Ch_Franquicia(Index) <> 0 Then
         Fr_LibCaja.Enabled = True
      Else
         Ch_ObligaLibComprasVentas = 0
         Fr_LibCaja.Enabled = False
      End If
   End If

End Sub

Private Sub Ch_ObligaLibComprasVentas_Click()
   If lInLoad Then
      Exit Sub
   End If

   If Ch_ObligaLibComprasVentas <> 0 Then
      MsgBox1 "ATENCIÓN: Si marca esta opción NO podrá ingresar manualmente nuevos documentos al libro de caja, sólo podrá traerlos desde los libros de Compras y Ventas", vbInformation
   Else
      MsgBox1 "ATENCIÓN: Si desmarca esta opción, deberá ingresar manualmente los documentos al libro de caja y no podrá traerlos desde el Libro de Compras y Ventas", vbInformation
   End If

End Sub

Private Sub Form_Load()
   SSTab1.Tab = 0
   
   lInLoad = True
     
   Call FillCombosFrm
   
   Call LoadAll
'   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetTxRO(Tx_RUT, True)
   Call SetTxRO(tx_NombreCorto, True)

   If gAppCode.Demo Then
      Call SetTxRO(Tx_RazonSocial, True)
      Call SetTxRO(Tx_Nombre, True)
      Call SetTxRO(Tx_ApMaterno, True)
   End If
   
   lInLoad = False
     
End Sub

Private Sub Op_TipoContrib_Click(Index As Integer)

   If Index = CONTRIB_SAABIERTA Then
      Fr_TrBolsa.Enabled = True
   Else
      Fr_TrBolsa.Enabled = False
      Op_TrBolsaNo = True
   End If
   
End Sub

Private Sub Tx_CodActEcon_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_CodActEcon_LostFocus()
   Call FindCbActEcon
   
End Sub

Private Sub Tx_Fax_KeyPress(KeyAscii As Integer)
   Call KeyTel(KeyAscii)
End Sub

Private Sub Tx_RUTContador_KeyPress(KeyAscii As Integer)
    Call KeyCID(KeyAscii)
End Sub

Private Sub Tx_RUTContador_LostFocus()
    If Tx_RutContador = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RutContador) Then
      Tx_RutContador.SetFocus
      Exit Sub
      
   End If
      
   MousePointer = vbHourglass
      
   Tx_RutContador = FmtCID(vFmtCID(Tx_RutContador))
   MousePointer = vbDefault
   
End Sub

Private Sub Tx_RUTRep_KeyPress(Index As Integer, KeyAscii As Integer)
    Call KeyCID(KeyAscii)
End Sub

Private Sub Tx_RUTRep_LostFocus(Index As Integer)
    If Tx_RUTRep(Index) = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RUTRep(Index)) Then
      Tx_RUTRep(Index).SetFocus
      Exit Sub
      
   End If
      
   MousePointer = vbHourglass
      
   Tx_RUTRep(Index) = FmtCID(vFmtCID(Tx_RUTRep(Index)))
   MousePointer = vbDefault
   
End Sub
Private Sub FillCombosFrm()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Codigo As String
   Dim MrkAnt As String
    
   'ACTIVIDAD ECONOMICA
   Set lcbCodActiv = New ClsCombo
   Call lcbCodActiv.SetControl(Cb_ActEcon)
   
   Q1 = "SELECT Descrip,Codigo,Version FROM CodActiv "
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
     ' Cb_ActEcon.AddItem vFld(Rs("Codigo")) & "  " & vFld(Rs("Descrip"), True)
     ' Cb_ActEcon.ItemData(Cb_ActEcon.NewIndex) = Val(vFld(Rs("Codigo")))
      
      '*** PS
      If vFld(Rs("Version")) = 1 Then
         MrkAnt = " ! "
      Else
         MrkAnt = "   "
      End If
      
      lcbCodActiv.AddItem vFld(Rs("Codigo")) & MrkAnt & vFld(Rs("Descrip"), True)
      lcbCodActiv.ItemData(lcbCodActiv.NewIndex) = vFld(Rs("Codigo")) 'Val(vFld(Rs("Codigo")))
      lcbCodActiv.List2(lcbCodActiv.NewIndex) = vFld(Rs("Version"))
     
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
  
   'COMBO REGION
   Call FillRegion(Cb_Region)
   
   Cb_Region.ListIndex = 0
   
   'COMUNA POSTAL, SE MUESTRAN TODAS LAS COMUNAS QU EXISTEN
   Q1 = "SELECT Comuna,id FROM Regiones"
   Q1 = Q1 & " ORDER BY Comuna"
   Cb_ComPostal.AddItem "< Ninguna >"
   Cb_ComPostal.ItemData(Cb_ComPostal.NewIndex) = 0
   Call FillCombo(Cb_ComPostal, DbMain, Q1, -1)
   
End Sub
Private Sub FindCbActEcon()
   If Tx_CodActEcon <> "" Then
      'Call SelItem(Cb_ActEcon, Right("00000" & Tx_CodActEcon, 5))
    '  Call SelItem(Cb_ActEcon, Val(Tx_CodActEcon))   'franca
      Call lcbCodActiv.SelItem(Val(Tx_CodActEcon))
      If lcbCodActiv.ListIndex = -1 Then
         MsgBox1 "Código actividad económica no existe", vbExclamation
         Tx_CodActEcon = ""
         Tx_CodActEcon.SetFocus
         
      End If
   End If
End Sub
Private Sub LoadAll()
   
   Dim Q1 As String
   Dim Rs As Recordset
   
'   Q1 = "UPDATE Empresa LEFT JOIN CodActiv ON Empresa.CodActEconom = CodActiv.OldCodigo SET Empresa.CodActEconom = CodActiv.Codigo, ActEconom = 0"
'      Call ExecSQL(DbMain, Q1)
   
   Tx_RUT = FmtCID(gEmpresa.Rut)
   tx_NombreCorto = gEmpresa.NombreCorto
      
   Q1 = "SELECT RazonSocial, ApMaterno, Nombre, Calle, Numero,"
   Q1 = Q1 & "EMail, Web, Dpto, Telefonos, Fax, Region, Comuna, Ciudad, ObsDTE, "
   Q1 = Q1 & "ActEconom, DomPostal, ComunaPostal, RepConjunta, RutRepLegal1,"
   Q1 = Q1 & "RepLegal1, RutRepLegal2, RepLegal2, CodActEconom, "
   Q1 = Q1 & "Giro, RutContador, Contador, FechaConstitucion, "
   Q1 = Q1 & "FechaInicioAct, "
   Q1 = Q1 & "TipoContrib, TransaBolsa, Franq14bis, FranqLey18392, FranqDL600, FranqDL701, FranqDS341, "
   Q1 = Q1 & "Franq14ter, Franq14quater, ObligaLibComprasVentas "
   Q1 = Q1 & " FROM Empresa"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      Tx_RazonSocial = vFld(Rs("RazonSocial"), True)
      Tx_Nombre = vFld(Rs("Nombre"), True)
      Tx_ApMaterno = vFld(Rs("ApMaterno"), True)
      Tx_Calle = vFld(Rs("Calle"), True)
      Tx_Numero = vFld(Rs("Numero"))
      Tx_EMail = vFld(Rs("Email"), True)
      Tx_Web = vFld(Rs("Web"), True)
      Tx_Dpto = vFld(Rs("Dpto"))
      Tx_Telefonos = vFld(Rs("Telefonos"))
      Tx_Fax = vFld(Rs("Fax"))
      Tx_Ciudad = vFld(Rs("Ciudad"))
      Tx_DirPostal = vFld(Rs("DomPostal"))
      Tx_ObsDTE = vFld(Rs("ObsDTE"))
      Call SetTxDate(Tx_FConstit, vFld(Rs("FechaConstitucion")))
      Call SetTxDate(Tx_FInicioAct, vFld(Rs("FechaInicioAct")))
      Tx_CodActEcon = vFld(Rs("CodActEconom"))
      Ch_RepConjunta = Abs(vFld(Rs("RepConjunta")) <> 0)
      If vFld(Rs("RutRepLegal1")) <> "" Then
         If vFld(Rs("RutRepLegal1")) = "0" Then
            Tx_RUTRep(0) = ""
         Else
            Tx_RUTRep(0) = FmtCID(vFld(Rs("RutRepLegal1")))
         End If
      End If
      Tx_NombreRep(0) = vFld(Rs("RepLegal1"), True)
      If vFld(Rs("RutRepLegal2")) <> "" Then
         If vFld(Rs("RutRepLegal2")) = "0" Then
            Tx_RUTRep(1) = ""
         Else
            Tx_RUTRep(1) = FmtCID(vFld(Rs("RutRepLegal2")))
         End If
      End If
      Tx_NombreRep(1) = vFld(Rs("RepLegal2"), True)
      Tx_Contador = vFld(Rs("Contador"), True)
      If vFld(Rs("RutContador")) <> "" Then
         Tx_RutContador = FmtCID(vFld(Rs("RutContador")))
      End If
      Tx_Giro = vFld(Rs("Giro"), True)
      
      Call SelItem(Cb_Region, vFld(Rs("Region")))
      Call SelItem(Cb_Comuna, vFld(Rs("Comuna")))
      
      'PS Ocupo siempre el CodActEconom, el Otro estaba de antes y ya no es necesario
      If vFld(Rs("CodActEconom"), True) <> "" Then
         Call lcbCodActiv.SelItem(Right("000000" & vFld(Rs("CodActEconom")), 6))
      End If
      
      Call SelItem(Cb_ComPostal, vFld(Rs("ComunaPostal")))
      
      If vFld(Rs("TipoContrib")) > 0 Then
      
         Op_TipoContrib(vFld(Rs("TipoContrib"))) = True
         
         Op_TrBolsaNo = True
         
         If vFld(Rs("TipoContrib")) = CONTRIB_SAABIERTA Then
            If vFld(Rs("TransaBolsa")) <> 0 Then
               Op_TrBolsaSi = True
            End If
         Else
            Fr_TrBolsa.Enabled = False
         End If
      
      Else
         Op_TrBolsaNo = True
         Fr_TrBolsa.Enabled = False

      End If
      
      Ch_Franquicia(FRANQ_14BIS) = Abs(vFld(Rs("Franq14bis")))
      Ch_Franquicia(FRANQ_LEY18392) = Abs(vFld(Rs("FranqLey18392")))
      Ch_Franquicia(FRANQ_DL600) = Abs(vFld(Rs("FranqDL600")))
      Ch_Franquicia(FRANQ_DL701) = Abs(vFld(Rs("FranqDL701")))
      Ch_Franquicia(FRANQ_DS341) = Abs(vFld(Rs("FranqDS341")))
      Ch_Franquicia(FRANQ_14TER) = Abs(vFld(Rs("Franq14ter")))
      Ch_Franquicia(FRANQ_14QUATER) = Abs(vFld(Rs("Franq14quater")))
      
      If Ch_Franquicia(FRANQ_14TER) = 0 Then
         Fr_LibCaja.Enabled = False
         Ch_ObligaLibComprasVentas = 0
      Else
         Ch_ObligaLibComprasVentas = Abs(vFld(Rs("ObligaLibComprasVentas")))
      End If
      
   End If
   Call CloseRs(Rs)
   
End Sub
Private Function Valida() As Boolean
   Dim i As Integer
   
   Valida = False
   
   If Trim(Tx_RazonSocial) = "" Then
      MsgBox1 "No se ha ingresado la razón social de la empresa.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Calle) = "" Or Trim(Tx_Numero) = "" Then
      MsgBox1 "No se ha ingresado la Dirección de la empresa. Falta Calle y/o Número.", vbExclamation
      Exit Function
   End If
   
   If CbItemData(Cb_Comuna) = 0 Then
      MsgBox1 "No se ha seleccionado la Comuna de la empresa.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Ciudad) = "" Then
      MsgBox1 "No se ha ingresado la Ciudad de la empresa.", vbExclamation
      Exit Function
   End If
   
   If Trim(Cb_Comuna) <> "" And InStr(Cb_Comuna, "(*)") <> 0 Then
      MsgBox1 "La comuna seleccionada está descontinuada por el SII.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_EMail) <> "" And ValidEmail(Tx_EMail) = False Then
      MsgBox1 "E-Mail inválido.", vbExclamation
      Exit Function
   End If
      
    If Trim(Tx_Giro) = "" Then
      MsgBox1 "No se ha ingresado el Giro de la empresa. Vaya a Antecedentes Legales para ingresarlo.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_CodActEcon) = "" Then
      MsgBox1 "No se ha ingresado el Código de la Actividad Económica de la empresa. Vaya a Antecedentes Legales para ingresarlo.", vbExclamation
      Exit Function
   End If

   For i = 0 To 1
      If Trim(Tx_RUTRep(i)) <> "" Then
         If vFmtCID(Tx_RUTRep(i)) > 50000000 Then   'tiene que ser personas naturales
            MsgBox1 "El RUT del representante legal debe corresponder a una persona natural."
            Exit Function
         End If
      End If
   Next i
   
   If Trim(Tx_RUTRep(0)) = "0-0" Then
      Tx_RUTRep(0) = ""
   End If
   If Trim(Tx_RUTRep(1)) = "0-0" Then
      Tx_RUTRep(1) = ""
   End If
   
   If Trim(Tx_RUTRep(0)) <> "" And Trim(Tx_RUTRep(1)) <> "" And vFmtCID(Tx_RUTRep(0)) = vFmtCID(Tx_RUTRep(1)) Then   'tiene que ser personas naturales
      MsgBox1 "Los RUTs de los representantes legales son iguales.", vbExclamation
      Exit Function
   End If
   
   'PS
'   If Val(lcbCodActiv.Matrix(2)) = VERSION_1 And gEmpresa.Ano = 2005 Then
'      If MsgBox1("¡ ADVERTENCIA ! " & vbNewLine & vbNewLine & "Código de Actividad Económica descontinuado." & vbNewLine & vbNewLine & "¿ Desea continuar ?", vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
'         Exit Function
'      End If
'   End If
   
   Valida = True
   
End Function

Private Sub SaveAll()
   Dim Q1 As String
   Dim TipoContrib As Integer
   Dim i As Integer
   Dim CodActEcono As String
         
   Q1 = "UPDATE Empresa SET "
   Q1 = Q1 & "RazonSocial='" & ParaSQL(Tx_RazonSocial) & "'"
   Q1 = Q1 & ", Nombre='" & ParaSQL(Tx_Nombre) & "'"
   Q1 = Q1 & ", ApMaterno='" & ParaSQL(Tx_ApMaterno) & "'"
   Q1 = Q1 & ", Calle='" & ParaSQL(Tx_Calle) & "'"
   Q1 = Q1 & ", Numero='" & ParaSQL(Tx_Numero) & "'"
   Q1 = Q1 & ", Dpto='" & ParaSQL(Tx_Dpto) & "'"
   Q1 = Q1 & ", EMail='" & ParaSQL(Tx_EMail) & "'"
   Q1 = Q1 & ", Telefonos='" & ParaSQL(Tx_Telefonos) & "'"
   Q1 = Q1 & ", Fax='" & ParaSQL(Tx_Fax) & "'"
   Q1 = Q1 & ", Ciudad='" & ParaSQL(Tx_Ciudad) & "'"
   Q1 = Q1 & ", DomPostal='" & ParaSQL(Tx_DirPostal) & "'"
   Q1 = Q1 & ", ComunaPostal=" & ItemData(Cb_ComPostal)
   Q1 = Q1 & ", Web='" & ParaSQL(Tx_Web) & "'"
   Q1 = Q1 & ", ObsDTE='" & ParaSQL(Tx_ObsDTE) & "'"
   'PS, elige una opción en la combo y borra el codigo
   If Trim(Tx_CodActEcon) = "" And lcbCodActiv.ListIndex > 0 Then
      Q1 = Q1 & ", CodActEconom='" & ParaSQL(lcbCodActiv.ItemData) & "'"
      CodActEcono = lcbCodActiv.ItemData
   Else
      Q1 = Q1 & ", CodActEconom='" & ParaSQL(Tx_CodActEcon) & "'"
      CodActEcono = Trim(Tx_CodActEcon)
   End If
   Q1 = Q1 & ", FechaConstitucion=" & GetTxDate(Tx_FConstit)
   Q1 = Q1 & ", FechaInicioAct=" & GetTxDate(Tx_FInicioAct)
   Q1 = Q1 & ", Region=" & ItemData(Cb_Region)
   Q1 = Q1 & ", Comuna=" & ItemData(Cb_Comuna)
   Q1 = Q1 & ", ActEconom=" & lcbCodActiv.ItemData  'ItemData(Cb_ActEcon) 'Este ya no se ocupa, porque el Codigo es String
   Q1 = Q1 & ", RepConjunta='" & CInt(Ch_RepConjunta <> 0) & "'"
   Q1 = Q1 & ", RutRepLegal1='" & vFmtCID(Tx_RUTRep(0)) & "'"
   Q1 = Q1 & ", RutRepLegal2='" & vFmtCID(Tx_RUTRep(1)) & "'"
   Q1 = Q1 & ", RepLegal1='" & ParaSQL(Tx_NombreRep(0)) & "'"
   Q1 = Q1 & ", RepLegal2='" & ParaSQL(Tx_NombreRep(1)) & "'"
   Q1 = Q1 & ", Giro='" & ParaSQL(Tx_Giro) & "'"
   Q1 = Q1 & ", Contador='" & ParaSQL(Tx_Contador) & "'"
   Q1 = Q1 & ", RutContador='" & vFmtCID(Tx_RutContador) & "'"
   
   TipoContrib = 0
   'For i = 1 To MAX_CONTRIB
   For i = 1 To 7
      If Op_TipoContrib(i) = True Then
         TipoContrib = i
         Exit For
      End If
   Next i
   
   Q1 = Q1 & ", TipoContrib=" & TipoContrib
   Q1 = Q1 & ", TContribFUT=" & TipoContrib
   Q1 = Q1 & ", TransaBolsa=" & CInt(Op_TrBolsaSi = True)
   Q1 = Q1 & ", Franq14bis=" & CInt(Ch_Franquicia(FRANQ_14BIS) <> 0)
   Q1 = Q1 & ", FranqLey18392=" & CInt(Ch_Franquicia(FRANQ_LEY18392) <> 0)
   Q1 = Q1 & ", FranqDL600=" & CInt(Ch_Franquicia(FRANQ_DL600) <> 0)
   Q1 = Q1 & ", FranqDL701=" & CInt(Ch_Franquicia(FRANQ_DL701) <> 0)
   Q1 = Q1 & ", FranqDS341=" & CInt(Ch_Franquicia(FRANQ_DS341) <> 0)
   Q1 = Q1 & ", Franq14ter=" & CInt(Ch_Franquicia(FRANQ_14TER) <> 0)
   Q1 = Q1 & ", Franq14quater=" & CInt(Ch_Franquicia(FRANQ_14QUATER) <> 0)
   
   Q1 = Q1 & ", ObligaLibComprasVentas=" & CInt(Ch_ObligaLibComprasVentas <> 0)
      
   Call ExecSQL(DbMain, Q1)
   
   lRc = vbOK
      
'   Q1 = "UPDATE ControlEmpresa SET"
'   Q1 = Q1 & " RazonSocial= '" & ParaSQL(Tx_RazonSocial) & "'"
'   Q1 = Q1 & " WHERE IdEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
'   Call ExecSQL(DbMain, Q1)

   gEmpresa.RazonSocial = Trim(Tx_RazonSocial) & " " & Trim(Tx_ApMaterno) & " " & Trim(Tx_Nombre)
   gEmpresa.Direccion = Trim(Tx_Calle) & " " & Trim(Tx_Numero) & " " & Trim(Tx_Dpto)
   gEmpresa.Telefono = Trim(Tx_Telefonos)
   gEmpresa.Comuna = Cb_Comuna
   gEmpresa.Ciudad = Tx_Ciudad
   gEmpresa.Fax = Trim(Tx_Fax)
   gEmpresa.Giro = Trim(Tx_Giro)
   gEmpresa.CodActEcono = CodActEcono
   gEmpresa.RepConjunta = (Ch_RepConjunta <> 0)
   gEmpresa.RutRepLegal1 = vFmtCID(Tx_RUTRep(0))
   gEmpresa.RutRepLegal2 = vFmtCID(Tx_RUTRep(1))
   gEmpresa.RepLegal1 = Trim(Tx_NombreRep(0))
   gEmpresa.RepLegal2 = Trim(Tx_NombreRep(1))
   gEmpresa.Franq14Ter = CInt(Ch_Franquicia(FRANQ_14TER) <> 0)
   gEmpresa.ObligaLibComprasVentas = CInt(Ch_ObligaLibComprasVentas <> 0)
   gEmpresa.email = Trim(Tx_EMail)
   gEmpresa.ObsDTE = Trim(Tx_ObsDTE)
   
   Call SetPrtData
End Sub
Private Sub Tx_FConstit_GotFocus()
   Call DtGotFocus(Tx_FConstit)
End Sub

Private Sub Tx_FConstit_LostFocus()
   
   If Trim$(Tx_FConstit) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FConstit)
   
End Sub

Private Sub Tx_FConstit_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_FInicioAct_GotFocus()
   Call DtGotFocus(Tx_FInicioAct)
End Sub

Private Sub Tx_FInicioAct_LostFocus()
   
   If Trim$(Tx_FInicioAct) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FInicioAct)
   
End Sub

Private Sub Tx_FInicioAct_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub
'PS
Private Sub La_Url_Click()
   Dim Rc As Long
   Dim Url As String
   
   Url = "http://www.sii.cl/catastro/codigos_economica.htm"
   Rc = Shell(gHtmExt.OpenCmd & " " & Url, vbNormalFocus)

End Sub

Private Sub Tx_Telefonos_KeyPress(KeyAscii As Integer)
   Call KeyTel(KeyAscii)
End Sub
