VERSION 5.00
Begin VB.Form FrmConfigConect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Conexión con Portal de Facturación Electrónica"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9390
   Icon            =   "FrmConfigConect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Firma"
      Height          =   735
      Left            =   1680
      TabIndex        =   23
      Top             =   5880
      Width           =   5655
      Begin VB.TextBox Tx_RutFirma 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Rut:"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   26
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "(Ej: 12345678-9)"
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   300
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mail Emisor"
      Height          =   1335
      Left            =   1680
      TabIndex        =   18
      Top             =   2520
      Width           =   5595
      Begin VB.TextBox Tx_Mail 
         Height          =   315
         Left            =   1620
         TabIndex        =   19
         ToolTipText     =   "Puede ingresar más de un mail separado por , (coma) o ; (punto y coma)"
         Top             =   300
         Width           =   3735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nota: permite más de un mail separado por , (coma) o ; (punto y coma)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   420
         TabIndex        =   22
         Top             =   840
         Width           =   4950
      End
      Begin VB.Label Label2 
         Caption         =   "Mail:"
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   20
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Fr_Certificado 
      Caption         =   "Clave Certificado"
      Height          =   1515
      Left            =   1680
      TabIndex        =   15
      Top             =   4080
      Width           =   5595
      Begin VB.TextBox Tx_ClaveCert 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Tx_Clave2Cert 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Clave:"
         Height          =   315
         Left            =   420
         TabIndex        =   17
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Repite Clave:"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   900
         Width           =   960
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7740
      TabIndex        =   7
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7740
      TabIndex        =   6
      Top             =   420
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de conexión"
      Height          =   1935
      Left            =   1740
      TabIndex        =   11
      Top             =   360
      Width           =   5595
      Begin VB.TextBox Tx_Clave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1260
         Width           =   2175
      End
      Begin VB.TextBox Tx_Clave 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Tx_Usuario 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Lb_UsrAcepta 
         Caption         =   "RUT del usuario"
         Height          =   255
         Left            =   3900
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Repite Clave:"
         Height          =   195
         Left            =   420
         TabIndex        =   14
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   315
         Left            =   420
         TabIndex        =   13
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario:"
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   12
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Proveedor de Facturación Electrónica"
      Height          =   975
      Left            =   1680
      TabIndex        =   9
      Top             =   7080
      Width           =   5595
      Begin VB.ComboBox Cb_Proveedor 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor:"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   240
      Picture         =   "FrmConfigConect.frx":000C
      ScaleHeight     =   675
      ScaleWidth      =   1125
      TabIndex        =   8
      Top             =   420
      Width           =   1185
   End
End
Attribute VB_Name = "FrmConfigConect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   If Not Valida Then
      Exit Sub
   End If
   
   Call SaveAll
   Unload Me
End Sub

Private Sub Cb_Proveedor_Click()
   Tx_Usuario = ""

   If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      Fr_Certificado.Visible = True
'      Me.Height = 7300
      Lb_UsrAcepta.Visible = True
   Else
      Fr_Certificado.Visible = False
      Me.Height = 4300
      Lb_UsrAcepta.Visible = False
   End If
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset

'   Call CbAddItem(Cb_Proveedor, "", 0, True)
   For i = 2 To UBound(gProvFactElect)       'no se agrega proveedor Thomson, sólo Acepta
      Call CbAddItem(Cb_Proveedor, gProvFactElect(i), i)
   Next i
   
'   Call CbAddItem(Cb_Proveedor, gProvFactElect(1), 1, True)
   gConectData.Proveedor = PROV_ACEPTA
   If gConectData.Proveedor > 0 Then
      Call CbSelItem(Cb_Proveedor, gConectData.Proveedor)
   End If
   
   If gConectData.Usuario <> "" Then
      If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      
'         If IsNumeric(gConectData.Usuario) = False Then
'            MsgBox1 "El Usuario de Conexión con Acepta debe ser un RUT (" & gConectData.Usuario & ")", vbExclamation
'            Exit Sub
'         End If
      
         Tx_Usuario = FmtRut(Val(gConectData.Usuario))
      Else
         Tx_Usuario = gConectData.Usuario
      End If
   
   End If
   
      If gConectData.RutFirma <> "" Then
      If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      
         If IsNumeric(gConectData.RutFirma) = False Then
            MsgBox1 "RUT firmante NO es valido : (" & gConectData.Usuario & ")", vbExclamation
            Exit Sub
         End If
      
         Tx_RutFirma = FmtRut(Val(gConectData.RutFirma))
      Else
         Tx_RutFirma = FmtRut(Val(gConectData.RutFirma))
      End If
   
   End If
   
   If gConectData.MailEmisor <> "" Then
      Tx_Mail = gConectData.MailEmisor
   End If
   
   Q1 = "SELECT Valor From ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_CLAVE & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      Tx_Clave.Text = vFld(Rs("Valor"))
      Tx_Clave2.Text = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor From ParamEmpDTE WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_CLAVECERT & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      Tx_ClaveCert.Text = vFld(Rs("Valor"))
      Tx_Clave2Cert.Text = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
End Sub

Private Function Valida() As Boolean
   Valida = False
   
   If CbItemData(Cb_Proveedor) = 0 Then
      MsgBox1 "Debe seleccionar un proveedor de Facturación Electrónica.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Usuario) = "" Then
      MsgBox1 "Debe ingresar un usuario para conectarse con el proveedor de Facturación Electrónica seleccionado.", vbExclamation
      Exit Function
   End If
   
   If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      If Not ValidRut(Tx_Usuario) Then
         MsgBox1 "Rut inválido.", vbExclamation
         Exit Function
      End If
   End If
   
   
   If Trim(Tx_Clave) = "" Then
      MsgBox1 "Debe ingresar una clave para conectarse con el proveedor de Facturación Electrónica seleccionado.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Clave) <> Trim(Tx_Clave2) Then
      MsgBox1 "Las claves ingresadas son distintas. Debe repetir la clave", vbExclamation
      Exit Function
   End If
  
   If Fr_Certificado.Visible Then
   
      If Trim(Tx_Mail) = "" Then
         MsgBox1 "Debe ingresar el mail del emisor (empresa) de los documentos electrónicos.", vbExclamation
         Exit Function
      End If
   
     If Trim(Tx_ClaveCert) = "" Then
         MsgBox1 "Debe ingresar la clave del usuario del certificado, correspondiente al proveedor de Facturación Electrónica seleccionado.", vbExclamation
         Exit Function
      End If
      
      If Trim(Tx_ClaveCert) <> Trim(Tx_Clave2Cert) Then
         MsgBox1 "Las claves del certificado ingresadas son distintas. Debe repetir la clave del certificado.", vbExclamation
         Exit Function
      End If
   End If
   
   If Trim(Tx_RutFirma) = "" Then
      MsgBox1 "Debe ingresar un RUT Firmante.", vbExclamation
      Exit Function
   ElseIf ValidRut(Tx_RutFirma) = False Then
      MsgBox1 "Rut Firma No es Valido", vbExclamation
      Exit Function
   End If
   
   Valida = True
  
End Function

Private Sub SaveAll()
   Dim Q1 As String
   
   Q1 = "UPDATE ParamEmpDTE SET Valor = '" & CbItemData(Cb_Proveedor) & "' WHERE Tipo ='" & CONECT_PROV & "'" & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   gConectData.Proveedor = CbItemData(Cb_Proveedor)
   
   If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      gConectData.Usuario = vFmtRut(Tx_Usuario)
   Else
      gConectData.Usuario = Trim(Tx_Usuario)
   End If
   Q1 = "UPDATE ParamEmpDTE SET Valor = '" & ParaSQL(gConectData.Usuario) & "' WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_USUARIO & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE ParamEmpDTE SET Valor = '" & ParaSQL(Tx_Clave) & "' WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_CLAVE & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   gConectData.Clave = Trim(Tx_Clave)
  
   Q1 = "UPDATE ParamEmpDTE SET Valor = '" & ParaSQL(Tx_Mail) & "' WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_MAILEMISOR & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   gConectData.MailEmisor = Trim(Tx_Mail)
   
   Q1 = "UPDATE ParamEmpDTE SET Valor = '" & ParaSQL(Tx_ClaveCert) & "' WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_CLAVECERT & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   gConectData.ClaveCert = Trim(Tx_ClaveCert)
   
   Q1 = "UPDATE ParamEmpDTE SET Valor = '" & ParaSQL(vFmtRut(Tx_RutFirma)) & "' WHERE Tipo ='" & CONECT_DATA & "' AND Codigo =" & CONECT_RUTFIRMA & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   gConectData.RutFirma = vFmtRut(Tx_RutFirma)
   
  
End Sub


Private Sub Tx_Mail_KeyPress(KeyAscii As Integer)
   Call KeyMail(KeyAscii)
End Sub

Private Sub Tx_RutFirma_KeyPress(KeyAscii As Integer)
Call KeyRut(KeyAscii)
End Sub

Private Sub Tx_RutFirma_LostFocus()
    If ValidRut(Tx_RutFirma) Then
       Tx_RutFirma = FmtRut(vFmtRut(Tx_RutFirma))
    Else
       MsgBox1 "Rut inválido.", vbExclamation
    End If
End Sub

Private Sub Tx_Usuario_KeyPress(KeyAscii As Integer)
   If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      Call KeyRut(KeyAscii)
   Else
      Call KeyName(KeyAscii)
   End If
End Sub
Private Sub Tx_Clave_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub
Private Sub Tx_Clave2_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub
Private Sub Tx_ClaveCert_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub
Private Sub Tx_Clave2Cert_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub

Private Sub Tx_Usuario_LostFocus()
   If CbItemData(Cb_Proveedor) = PROV_ACEPTA Then
      If ValidRut(Tx_Usuario) Then
         Tx_Usuario = FmtRut(vFmtRut(Tx_Usuario))
      Else
         MsgBox1 "Rut inválido.", vbExclamation
      End If
      
   Else
      Tx_Usuario = Trim(Tx_Usuario)
   End If
End Sub

