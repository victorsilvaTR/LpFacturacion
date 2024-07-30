VERSION 5.00
Begin VB.Form FrmIdUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Identificación del Usuario"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   Icon            =   "FrmIdUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5160
   Begin VB.CheckBox Ck_Version 
      Caption         =   "Verificar si hay nueva actualización"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CheckBox Ck_Link 
      Caption         =   "Crear &ícono en el Escritorio"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Tx_Clave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   900
      Width           =   1575
   End
   Begin VB.TextBox Tx_Nombre 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label La_Demo 
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
      ForeColor       =   &H002A01A6&
      Height          =   330
      Left            =   3300
      TabIndex        =   8
      Top             =   1380
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clave:"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   540
      Width           =   585
   End
End
Attribute VB_Name = "FrmIdUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_Cancel_Click()
   
   gUsuario.Rc = vbCancel
      
   Unload Me
   
End Sub

Private Sub Form_Activate()

   On Error Resume Next

   If Tx_Nombre <> "" Then
      Tx_Clave.SetFocus
   End If

End Sub
Private Sub Form_Load()
   Dim Rc As Long
   Dim Buf As String * 21

   gUsuario.Rc = vbCancel
   
   Tx_Nombre = GetLastUser(gIniFile, gAdmUser)
   
   FrmIdUser.Top = FrmStart.Top + (FrmStart.Height - FrmIdUser.Height) / 2 - 560
   FrmIdUser.Left = FrmStart.Left + FrmStart.Width - FrmIdUser.Width
   
   'Ck_Link.Value = Abs(Val(GetIniString(gIniFile, "Config", "Link", "1")) <> 0)
   
   Ck_Version.Value = Abs(Val(GetIniString(gIniFile, "Config", "ChkVer1", "1")) <> 0)
   
   If Ck_Version.Value = 0 Then
      Ck_Version.Value = Abs(Abs(CLng(Now) - Val(GetIniString(gIniFile, "Config", "LstChk"))) > 14) ' cada 14 días se
   End If

   La_demo.Visible = gAppCode.Demo
   
   If APP_DEMO Then
      Me.Caption = Me.Caption & " - Sólo DEMO"
   End If
     
End Sub

Private Sub bt_OK_Click()
   Dim Rc As Long
   
   If ValidaUsuario() = False Then
      Exit Sub
   End If
        
   Call SetLastUser(gIniFile, gUsuario.Nombre)
   
   Call SetIniString(gIniFile, "Config", "Link", Abs(Ck_Link))
   Call SetIniString(gIniFile, "Config", "ChkVer1", Abs(Ck_Version))
   
   If Ck_Link Then
      Call CreateLnk("$Desktop\Legal Publishing Facturación")
   End If
   
   If Ck_Version.Value Then  ' pam : 23 jun 2015
      If FwCheckVersion(Me, True, APP_NAME, APP_URL) Then
         Call SetIniString(gIniFile, "Config", "LstChk", CLng(Int(Now)))
      End If
   End If

   
   gUsuario.Rc = vbOK
   
   Unload Me
      
End Sub

Private Function ValidaUsuario() As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   ValidaUsuario = False
   gUsuario.Nombre = LCase(Trim(Tx_Nombre))
      
   Q1 = "SELECT IdUsuario, Clave, NombreLargo, Privilegios "
   Q1 = Q1 & " FROM Usuarios "
   Q1 = Q1 & " LEFT JOIN Perfiles ON Usuarios.IdPerfil = Perfiles.IdPerfil "
   Q1 = Q1 & " WHERE Usuario = '" & gUsuario.Nombre & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then
      MsgBox1 "Usuario desconocido", vbExclamation
      Call CloseRs(Rs)
      Exit Function
   End If
            
   If Trim(Rs("Clave")) <> GenClave(LCase(gUsuario.Nombre & Trim(Tx_Clave))) Then
      Debug.Print GenClave(LCase(gUsuario.Nombre & Trim(Tx_Clave)))
      If W.InDesign Then
         If MsgBox1("Modo diseño ¿Desea verificar la clave?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            MsgBox1 "Clave incorrecta", vbExclamation
            Call CloseRs(Rs)
            Exit Function
         End If
      Else
         MsgBox1 "Clave incorrecta", vbExclamation
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
      
'   If vFld(Rs("idEmpresas")) = "" And gUsuario.Nombre <> gAdmUser Then
'      MsgBox1 "Usuario no tiene empresas asignadas para seleccionar", vbExclamation
'      Call CloseRs(Rs)
'      Exit Function
'   End If
   
   gUsuario.IdUsuario = vFld(Rs("idUsuario"))
   gUsuario.ClaveACtual = Val(vFld(Rs("Clave")))
   gUsuario.NombreLargo = vFld(Rs("NombreLargo"), True)
   
   If gUsuario.Nombre = gAdmUser Then
      gUsuario.Priv = PRV_ADMIN
   Else
      gUsuario.Priv = vFld(Rs("Privilegios"))
   End If
      
   Call CloseRs(Rs)
   
   ValidaUsuario = True
   
End Function
