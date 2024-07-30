VERSION 5.00
Begin VB.Form FrmDatosAdicDTE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Adicionales"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Transporte 
      Caption         =   "Información Transporte"
      Height          =   2115
      Left            =   1320
      TabIndex        =   5
      Top             =   540
      Width           =   5295
      Begin VB.CommandButton Bt_SelConductor 
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
         Left            =   3840
         Picture         =   "FrmDatosAdicDTE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccionar Conductor"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Bt_SelVehiculo 
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
         Left            =   3840
         Picture         =   "FrmDatosAdicDTE.frx":047E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Seleccionar vehículo"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Tx_NombreChofer 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1440
         Width           =   3075
      End
      Begin VB.TextBox Tx_RutChofer 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   1
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox Tx_Patente 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   0
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre  Chofer:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   9
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT Chofer:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   7
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Patente vehículo:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   420
      Picture         =   "FrmDatosAdicDTE.frx":093C
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   660
      Width           =   555
   End
   Begin VB.CommandButton bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   3
      Top             =   960
      Width           =   1275
   End
End
Attribute VB_Name = "FrmDatosAdicDTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lDTEGuiaDesp As DTEGuiaDesp_t
Dim lRc As Integer

Private Sub Bt_Cancelar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub bt_OK_Click()

   Call SaveData
   
   lRc = vbOK
   
   Unload Me
End Sub

Friend Function FEdit(DTEGuiaDesp As DTEGuiaDesp_t) As Integer

   lDTEGuiaDesp = DTEGuiaDesp
   
   Me.Show vbModal
   
   DTEGuiaDesp = lDTEGuiaDesp
   
   FEdit = lRc
   
End Function
Private Function LoadData()

   Tx_Patente = lDTEGuiaDesp.Patente
   Tx_RutChofer = IIf(lDTEGuiaDesp.RutChofer <> "", FmtCID(lDTEGuiaDesp.RutChofer), "")
   Tx_NombreChofer = IIf(lDTEGuiaDesp.NombreChofer <> "", lDTEGuiaDesp.NombreChofer, "")

End Function
Private Function SaveData()

   lDTEGuiaDesp.Patente = Trim(Tx_Patente)
   lDTEGuiaDesp.RutChofer = vFmtCID(Tx_RutChofer)
   lDTEGuiaDesp.NombreChofer = Trim(Tx_NombreChofer)
   
End Function

Private Sub Bt_SelConductor_Click(Index As Integer)
   Dim Frm As FrmMantConductores
   Dim IdConductor As Long
   Dim Nombre As String, Rut As String
   Dim Rc As Integer
   
   Set Frm = New FrmMantConductores
   Rc = Frm.FView(IdConductor, Nombre, Rut)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Tx_NombreChofer = Nombre
      Tx_RutChofer = Rut
   End If
   
End Sub

Private Sub Bt_SelVehiculo_Click(Index As Integer)
   Dim Frm As FrmMantVehiculos
   Dim IdVehiculo As Long
   Dim Patente As String
   Dim Rc As Integer
   
   Set Frm = New FrmMantVehiculos
   Rc = Frm.FView(IdVehiculo, Patente)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Tx_Patente = Patente
   End If
End Sub

Private Sub Form_Load()
   Call LoadData
End Sub

Private Sub Tx_Patente_KeyPress(KeyAscii As Integer)
   
   Call KeyUCod(KeyAscii)
   
End Sub

Private Sub Tx_RutChofer_KeyPress(KeyAscii As Integer)

   Call KeyRut(KeyAscii)
   
End Sub

Private Sub Tx_RutChofer_LostFocus()
   
   If Tx_RutChofer = "" Then
      Exit Sub
   End If
   
   If vFmtCID(Tx_RutChofer) = 0 Then
      Tx_RutChofer = ""
      Tx_RutChofer.SetFocus
      Exit Sub
   End If
   
'   If Not MsgValidRut(Tx_RutChofer) Then
'      Tx_RutChofer.SetFocus
'      Exit Sub
'
'   End If
'
   Tx_RutChofer = FmtCID(vFmtCID(Tx_RutChofer))
   
End Sub
Private Sub Tx_RutChofer_Validate(Cancel As Boolean)
   
   If Tx_RutChofer = "" Then
      Exit Sub
   End If
   
   If Trim(Tx_RutChofer) = "0-0" Then
      MsgBox1 "RUT Inválido.", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   
   If Not MsgValidRut(Tx_RutChofer) Then
      Tx_RutChofer.SetFocus
      Cancel = True
      Exit Sub
      
   End If
   
   
End Sub


