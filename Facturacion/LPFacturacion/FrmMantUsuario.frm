VERSION 5.00
Begin VB.Form FrmMantUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración de usuario"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8850
   Icon            =   "FrmMantUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7380
      TabIndex        =   8
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7380
      TabIndex        =   7
      Top             =   480
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   3195
      Left            =   1440
      TabIndex        =   9
      Top             =   420
      Width           =   5535
      Begin VB.ComboBox Cb_Perfil 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2460
         Width           =   3015
      End
      Begin VB.TextBox Tx_Clave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1920
         Width           =   1035
      End
      Begin VB.TextBox Tx_Clave1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   1515
      End
      Begin VB.TextBox Tx_NombreLargo 
         Height          =   315
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   1
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox Ck_PrvAdm 
         Caption         =   "Administrador del Sistema"
         Height          =   255
         Left            =   2940
         TabIndex        =   5
         Top             =   1980
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CheckBox Ch_Clave 
         Caption         =   "Modifica clave"
         Height          =   255
         Left            =   2940
         TabIndex        =   4
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Perfil:"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   14
         Top             =   2580
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Repita  clave:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   1980
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese clave:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   540
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Largo:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   10
         Top             =   1020
         Width           =   1050
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   420
      Picture         =   "FrmMantUsuario.frx":000C
      Top             =   540
      Width           =   750
   End
End
Attribute VB_Name = "FrmMantUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Oper As Integer
Dim OldNombre As String
Dim ModClave As Integer
Dim lidUsuario As Long
Dim lNombre As String
Dim lRc As Integer
Dim ClaveACtual As String

Private Sub bt_Cancel_Click()
   Unload Me
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   lRc = vbCancel
   
   Q1 = "SELECT Nombre, IdPerfil FROM Perfiles WHERE IdApp = " & 0 & " ORDER BY Nombre"
   
   Call FillCombo(Cb_Perfil, DbMain, Q1, -1)
  
   If Oper = O_EDIT Then
      Caption = "Editar usuario"
      Call FillForm
      
   Else
      Caption = "Nuevo usuario"
      
   End If
      
            
End Sub

Private Sub bt_OK_Click()

   If Valida() = False Then
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass
   Call SaveAll
   MousePointer = vbDefault
   
   lRc = vbOK
   lNombre = Tx_Nombre
   Unload Me
   
End Sub
Public Function FNew(Nombre As String, IdUsuario As Long) As Integer
  
   Oper = O_NEW
   Me.Show vbModal
   
   FNew = lRc
   Nombre = lNombre
   IdUsuario = lidUsuario
   
End Function

Public Function FEdit(Nombre As String, ByVal IdUsuario As Long) As Integer
   lidUsuario = IdUsuario
   Oper = O_EDIT

   Me.Show vbModal
   FEdit = lRc
   Nombre = lNombre
   
End Function

Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim t_Nombre As String
   Dim Clave As String
   Dim UserExist As Integer
   Dim StrClave As String, QryWhere As String
   
   t_Nombre = LCase(Trim(Tx_Nombre))
   Clave = LCase(Trim(Tx_Clave1))
   
   Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & t_Nombre & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      UserExist = True
      
   End If
   Call CloseRs(Rs)
   
   If Ch_Clave Or Oper = O_NEW Then
      Clave = GenClave(LCase(t_Nombre & Clave))
      StrClave = " Clave =" & Clave & ","
   End If
   
   Select Case Oper
      
      Case O_NEW
         If UserExist = True Then
            MsgBox1 "Este usuario ya existe.", vbExclamation
            Exit Sub
                     
         Else
            lidUsuario = TbAddNew(DbMain, "Usuarios", "idUsuario", "Usuario")
            
            Q1 = "UPDATE Usuarios SET " & StrClave & " Usuario = '" & t_Nombre & "'"
            Q1 = Q1 & ",NombreLargo='" & ParaSQL(Tx_NombreLargo) & "'"
            Q1 = Q1 & ",IdPerfil=" & CbItemData(Cb_Perfil)
            Q1 = Q1 & " WHERE IdUsuario =" & lidUsuario & QryWhere
            Call ExecSQL(DbMain, Q1)
            
         End If
      
      Case O_EDIT
         If t_Nombre <> OldNombre And UserExist = True Then    'cambió nombre por uno que ya existe
            MsgBox1 "Este usuario ya existe.", vbExclamation
            Exit Sub
         
         Else
            
            Q1 = "UPDATE Usuarios SET " & StrClave & " Usuario = '" & t_Nombre & "'"
            Q1 = Q1 & ", NombreLargo ='" & ParaSQL(Tx_NombreLargo) & "'"
            Q1 = Q1 & ", IdPerfil =" & CbItemData(Cb_Perfil)
            Q1 = Q1 & " WHERE IdUsuario =" & lidUsuario & QryWhere
            Call ExecSQL(DbMain, Q1)
            
         End If
         
         
   End Select
     
   If gUsuario.IdUsuario = lidUsuario Then    'actualizamos los datos del usuario actual si corresponde
   
      gUsuario.Nombre = t_Nombre
      gUsuario.NombreLargo = Trim(Tx_NombreLargo)
      If Ch_Clave Or Oper = O_NEW Then
         gUsuario.ClaveACtual = Clave
      End If
      Q1 = "SELECT Privilegios FROM Perfiles WHERE IdPerfil = " & CbItemData(Cb_Perfil)
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         gUsuario.Priv = vFld(Rs("Privilegios"))
      End If
      Call CloseRs(Rs)
   End If
   
End Sub

Private Function Valida() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim t_Nombre As String
   Dim i As Integer
   
   Valida = False
   t_Nombre = LCase(Trim(Tx_Nombre))
  
   If t_Nombre = "" Then
      MsgBox1 "Debe ingresar nombre.", vbExclamation
      Tx_Nombre.SetFocus
      Exit Function
   End If
   
   If Trim(Tx_NombreLargo) = "" Then
      MsgBox1 "Debe ingresar nombre largo", vbExclamation
      Tx_NombreLargo.SetFocus
      Exit Function
      
   End If
   
   Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & t_Nombre & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then    ' existe
      If Oper = O_NEW Then
         MsgBox1 "Este usuario ya existe.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
         
      ElseIf Oper = O_EDIT And t_Nombre <> OldNombre Then    'cambió nombre por uno que ya existe
         MsgBox1 "Este usuario ya existe.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
         
      End If
      
   End If
   Call CloseRs(Rs)
   
   If LCase(Trim(Tx_Clave1)) <> LCase(Trim(Tx_Clave2)) Then
      MsgBox1 "Las claves son distintas.", vbExclamation + vbOKOnly
      Tx_Clave1.SetFocus
      Exit Function
      
   End If
     
   Valida = True
   
End Function


Private Sub Tx_Clave1_Change()
   Ch_Clave.Value = 1
End Sub

Private Sub Tx_Clave1_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub


Private Sub Tx_Clave2_Change()
   Ch_Clave.Value = 1

End Sub

Private Sub Tx_Clave2_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
   
End Sub

Private Sub Tx_Nombre_KeyPress(KeyAscii As Integer)
   Call KeyUserId(KeyAscii)
   
End Sub
Private Sub FillForm()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Usuario, Clave, NombreLargo, idPerfil FROM Usuarios "
   Q1 = Q1 & " WHERE idUsuario=" & lidUsuario
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      OldNombre = vFld(Rs("Usuario"))
      Tx_Nombre = vFld(Rs("Usuario"))
      ClaveACtual = vFld(Rs("Clave"))
      Tx_NombreLargo = vFld(Rs("NombreLargo"), True)
      Call CbSelItem(Cb_Perfil, vFld(Rs("IdPerfil")))
      
   End If
   Call CloseRs(Rs)
End Sub
