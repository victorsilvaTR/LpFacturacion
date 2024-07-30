VERSION 5.00
Begin VB.Form FrmSelEmpresas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Empresa"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "FrmSelEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   1620
      TabIndex        =   16
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   4020
         Picture         =   "FrmSelEmpresas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   900
         TabIndex        =   5
         ToolTipText     =   "Ingrese cualquier parte dell nombre o razón social de la Entidad"
         Top             =   720
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   780
         Width           =   600
      End
   End
   Begin VB.ListBox Ls_Ano 
      Height          =   1035
      Left            =   7260
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Fr_Sort 
      Caption         =   "Ordenar por"
      Height          =   975
      Left            =   7260
      TabIndex        =   14
      Top             =   3240
      Width           =   1275
      Begin VB.OptionButton Op_SortNombre 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Op_SortRUT 
         Caption         =   "RUT"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.ListBox Ls_Empresas 
      Height          =   4155
      Left            =   1620
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Default         =   -1  'True
      Height          =   315
      Left            =   7320
      TabIndex        =   7
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Año:"
      Height          =   195
      Index           =   2
      Left            =   7260
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   0
      Left            =   480
      Picture         =   "FrmSelEmpresas.frx":055C
      Top             =   360
      Width           =   690
   End
   Begin VB.Label La_nEmp 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   195
      Left            =   7200
      TabIndex        =   13
      Top             =   4800
      Width           =   270
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
      ForeColor       =   &H002A01A6&
      Height          =   330
      Left            =   7440
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   1
      Left            =   3060
      TabIndex        =   11
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUT"
      Height          =   315
      Index           =   6
      Left            =   1620
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   0
      Left            =   2820
      TabIndex        =   9
      Top             =   1560
      Width           =   2115
   End
End
Attribute VB_Name = "FrmSelEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_NOMLARGO = 2
Private Const C_IDEMPRESA = 3
Private Const C_IDPERFIL = 4
Private Const C_PRIV = 5

Private Const C_FCIERRE = 2
Private Const C_FAPERTURA = 3

Dim lRc As Integer
Dim lsEmpresa As ClsCombo
Dim LsAno As ClsCombo
Friend Function FSelect() As Integer
   Me.Show vbModal
   
   FSelect = lRc
End Function
Private Sub bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub
Private Sub Bt_Sel_Click()
   Dim IdEmpresa As Long
   Dim Ano As Integer
   Dim Rut As String
   Dim Nombre As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 1", 1)
   
   If lsEmpresa.ListIndex < 0 Then
      Exit Sub
   End If
   
'   If LsAno.ListIndex < 0 Then
'      Exit Sub
'   End If
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 2 - " & lsEmpresa.ListIndex & " - " & LsAno.ListIndex, 1)
   
   MousePointer = vbHourglass
   DoEvents
   
   IdEmpresa = Val(lsEmpresa.Matrix(C_IDEMPRESA))
'   Ano = Val(LsAno.ItemData)
   Rut = lsEmpresa.ItemData
   Nombre = lsEmpresa.List2(lsEmpresa.ListIndex)
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 3", 1)

   'Creo o chequeo base de datos de la empresa
'   If CrearNuevoAnoFact(IdEmpresa, Ano, Rut, Nombre) = False Then
   If CrearNuevaEmprFact(IdEmpresa, Rut, Nombre) = False Then
      MousePointer = vbDefault
      Exit Sub
   End If
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 4", 1)
   
   'ASIGNO DATOS A LA ESTRUCTURA
   gEmpresa.Rut = Rut
   gEmpresa.NombreCorto = Nombre
   gEmpresa.id = IdEmpresa
'   gEmpresa.Ano = Ano
   'gEmpresa.FCierre = vFmt(LsAno.ItemData(LsAno.ListIndex))
'   gEmpresa.FCierre = vFmt(LsAno.Matrix(C_FCIERRE))
'   gEmpresa.FApertura = vFmt(LsAno.Matrix(C_FAPERTURA))
      
'   gUsuario.idPerfil = lsEmpresa.Matrix(C_IDPERFIL)
'   gUsuario.Priv = lsEmpresa.Matrix(C_PRIV)
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: FIN", 1)
   MousePointer = vbDefault
   
   lRc = vbOK
   Unload Me
   
  
End Sub

Private Sub Form_Load()

   lRc = vbCancel
   
   Call AddDebug("FrmSelEmpresas: Load", 1)
   
   Set LsAno = New ClsCombo
   Call LsAno.SetControl(Ls_Ano)
      
   If gVarIniFile.SelEmprPorRUT Then
      Op_SortRUT = True    'llama a FillList
   Else
      Op_SortNombre = True    'llama a FillList
   End If
   
   Call AddDebug("FrmSelEmpresas: Después de FillList", 1)
   
   La_demo.Visible = gAppCode.Demo
   'Fr_Sort.Visible = Not gAppCode.Demo
   
End Sub

Private Sub FillList()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
      
   Set lsEmpresa = New ClsCombo
   Call lsEmpresa.SetControl(Ls_Empresas)
      
   Q1 = "SELECT Empresas.idEmpresa, Rut, NombreCorto"
   Q1 = Q1 & " FROM Empresas"
   Q1 = Q1 & " WHERE Estado = 0 "   'Empresas activas
   
   If gAppCode.Demo Then
      Q1 = Q1 & " AND RUT IN ('1','2','3')"
   
   Else
      If Tx_Rut <> "" Then
         Q1 = Q1 & " AND Rut = '" & vFmtCID(Tx_Rut) & "'"
      End If
   
      If Tx_Nombre <> "" Then
         Q1 = Q1 & " AND " & GenLike(DbMain, Tx_Nombre, "NombreCorto")
      End If

   End If
   
   If Op_SortNombre Then
      Q1 = Q1 & " ORDER BY NombreCorto, RUT"
   Else
      Q1 = Q1 & " ORDER BY right('0' & RUT,8)"
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      lsEmpresa.AddItem FmtStRut(vFld(Rs("Rut"))) & vbTab & vFld(Rs("NombreCorto"))
      lsEmpresa.ItemData(lsEmpresa.NewIndex) = vFld(Rs("Rut"))
      lsEmpresa.Matrix(C_NOMLARGO, lsEmpresa.NewIndex) = vFld(Rs("NombreCorto"))
      lsEmpresa.Matrix(C_IDEMPRESA, lsEmpresa.NewIndex) = vFld(Rs("idEmpresa"))
'      lsEmpresa.Matrix(C_IDPERFIL, lsEmpresa.NewIndex) = vFld(Rs("idPerfil"))
'
'      If gUsuario.Nombre = gAdmUser Then
'         lsEmpresa.Matrix(C_PRIV, lsEmpresa.NewIndex) = PRV_ADMIN
'      Else
'         lsEmpresa.Matrix(C_PRIV, lsEmpresa.NewIndex) = vFld(Rs("Privilegios"))
'      End If
      
      Rs.MoveNext
      
      If gAppCode.Demo And lsEmpresa.ListCount >= 3 Then
         Exit Do
      End If
      
   Loop
   Call CloseRs(Rs)
   
   La_nEmp = Format(lsEmpresa.ListCount, NUMFMT)
   
   If gAppCode.NivProd = VER_5EMP And lsEmpresa.ListCount > 5 Then
      MsgBox1 "Esta versión sólo permite trabajar con a lo más 5 empresas." & vbCrLf & vbCrLf & "Utilice el administrador para eliminar algunas empresas y así poder utilizar el sistema.", vbExclamation
      Bt_Sel.Enabled = False
   End If
   
End Sub



Private Sub Ls_Ano_DblClick()
   Call PostClick(Bt_Sel)
   
   Call AddDebug("FrmSelEmpresas.Ls_Ano_DblClick: FIN", 1)

End Sub
'
'Private Sub Ls_Empresas_Click()
'   Dim Q1 As String
'   Dim Ano As Integer, MaxAno As Integer
'   Dim Rs As Recordset
'   Dim UltAño As Integer
'   Dim i As Integer
'   Dim AnoTope As Integer
'
'   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: 1", 1)
'
'   LsAno.Clear
'   Ano = Year(Int(Now))
'
'   Q1 = "SELECT Max(Ano) as MaxAno FROM EmpresasAno"
'   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
'
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF = False Then
'      MaxAno = vFld(Rs("MaxAno"))
'   End If
'   Call CloseRs(Rs)
'
'   If MaxAno <= 0 Then
'      MaxAno = Year(Now)
'   End If
'
'   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: 2 - " & MaxAno, 1)
'
'   Q1 = "SELECT Ano, FCierre, FApertura FROM EmpresasAno"
'   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
'   Q1 = Q1 & " ORDER BY Ano DESC"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: 3", 1)
'
'   AnoTope = Year(Now) + 2
'
'   For i = AnoTope To 2000 Step -1
'      Call LsAno.AddItem(i, i, 0, 0)
'      If i = MaxAno Then
'         LsAno.ListIndex = LsAno.NewIndex
'      End If
'
'      Do Until Rs.EOF
'         Ano = vFld(Rs("Ano"))
'
'         If Ano > AnoTope Then ' años muy del futuro o tiene mal la fecha del computador
'            Exit Do
'         End If
'
'         If i = Ano Then
'            LsAno.List(LsAno.NewIndex) = Ano & " *"
'            LsAno.Matrix(C_FCIERRE, LsAno.NewIndex) = vFld(Rs("FCierre"))
'            LsAno.Matrix(C_FAPERTURA, LsAno.NewIndex) = vFld(Rs("FApertura"))
'            Rs.MoveNext
'            Exit Do
'         ElseIf i > Ano Then
'            Exit Do
'         End If
'      Loop
'
'   Next i
'
'   Call CloseRs(Rs)
'
'   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: FIN", 1)
'
'End Sub
Private Sub Ls_Empresas_DblClick()

   Call PostClick(Bt_Sel)
   
End Sub

Private Sub Op_SortNombre_Click()

   Me.MousePointer = vbHourglass
   Call FillList
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", "0")
   gVarIniFile.SelEmprPorRUT = 0
   Me.MousePointer = vbDefault

End Sub

Private Sub Op_SortRUT_Click()

   Me.MousePointer = vbHourglass
   Call FillList
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", "1")
   gVarIniFile.SelEmprPorRUT = 1
   Me.MousePointer = vbDefault

End Sub
Private Sub Tx_Rut_LostFocus()
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If vFmtCID(Tx_Rut) = "" Then
      Tx_Rut = ""
      Tx_Rut.SetFocus
      Exit Sub
   End If
   
'   If Not MsgValidRut(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
'
   Tx_Rut = FmtCID(vFmtCID(Tx_Rut))
   
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Trim(Tx_Rut) = "0-0" Then
      MsgBox1 "RUT Inválido.", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_Rut) Then
      Tx_Rut.SetFocus
      Cancel = True
      Exit Sub
      
   End If
   
   
End Sub
Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)
      
   Call KeyCID(KeyAscii)
   
End Sub


