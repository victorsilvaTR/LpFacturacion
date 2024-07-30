VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Codigos de Licencia de Red - Fairware - Legal Publishing"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8220
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Copy2Lic 
      Caption         =   "Lic"
      Height          =   315
      Left            =   7080
      TabIndex        =   21
      ToolTipText     =   "Copiar para Licencias"
      Top             =   180
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   240
      TabIndex        =   12
      Top             =   540
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Cliente"
      TabPicture(0)   =   "FrmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Tx_Nivel"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Ck_Clip"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Tx_RUT"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Tx_Clave"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Tx_nPC"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Tx_NetCode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Bt_Parse"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Tx_Verif"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Tx_Info"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Br_Clear"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Tx_nLic"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Bt_Paste"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Equipos a Inscribir"
      TabPicture(1)   =   "FrmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tx_nPC_Dif"
      Tab(1).Control(1)=   "Cb_Equ"
      Tab(1).Control(2)=   "Ck_Celda"
      Tab(1).Control(3)=   "Grid"
      Tab(1).ControlCount=   4
      Begin VB.TextBox tx_nPC_Dif 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   -68580
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "Cantidad de equipos diferentes"
         Top             =   6300
         Width           =   675
      End
      Begin VB.ComboBox Cb_Equ 
         Height          =   315
         Left            =   -69240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2700
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Ck_Celda 
         Caption         =   "Copiar sólo la celda seleccionada"
         Height          =   315
         Left            =   -74160
         TabIndex        =   22
         Top             =   6300
         Width           =   3255
      End
      Begin VB.CommandButton Bt_Paste 
         Caption         =   "Pegar"
         Height          =   315
         Left            =   4740
         TabIndex        =   0
         Top             =   420
         Width           =   1155
      End
      Begin VB.TextBox Tx_nLic 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   5940
         Width           =   675
      End
      Begin VB.CommandButton Br_Clear 
         Caption         =   "Limpiar"
         Height          =   315
         Left            =   6300
         TabIndex        =   10
         Top             =   420
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5715
         Left            =   -74220
         TabIndex        =   19
         Top             =   540
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   10081
         _Version        =   393216
         Rows            =   30
         FixedCols       =   0
      End
      Begin VB.TextBox Tx_Info 
         Height          =   4635
         Left            =   180
         MaxLength       =   20000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   780
         Width           =   7335
      End
      Begin VB.TextBox Tx_Verif 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         Top             =   5520
         Width           =   1395
      End
      Begin VB.CommandButton Bt_Parse 
         Caption         =   "Generar Código de Red"
         Default         =   -1  'True
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   6360
         Width           =   2175
      End
      Begin VB.TextBox Tx_NetCode 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6360
         Width           =   1755
      End
      Begin VB.TextBox Tx_nPC 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5940
         Width           =   675
      End
      Begin VB.TextBox Tx_Clave 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   6360
         Width           =   1395
      End
      Begin VB.TextBox Tx_RUT 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5940
         Width           =   1395
      End
      Begin VB.CheckBox Ck_Clip 
         Height          =   315
         Left            =   4260
         TabIndex        =   9
         ToolTipText     =   "Copiar código al portapapeles."
         Top             =   6360
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox Tx_Nivel 
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   5520
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Licencias:"
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   20
         Top             =   6000
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pegue aquí la información enviada por el cliente:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   5580
         Width           =   405
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "&Verificador:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   5280
         TabIndex        =   16
         Top             =   5580
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Equipos:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   6000
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT cliente:"
         Height          =   195
         Index           =   4
         Left            =   5160
         TabIndex        =   14
         Top             =   6000
         Width           =   900
      End
   End
   Begin VB.Label La_Prod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "LP Contabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   300
      TabIndex        =   11
      Top             =   180
      Width           =   7650
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' **** NO CAMBIAR *********
' Está en HyperComun.bas
Private Const PC_SEED = 765432
Private Const PC_SEED_C4 = 391719      ' LP Contab 4
Private Const PC_SEED_CS = 319175      ' LP Contab SQL - 17 jul 2019
Private Const PC_SEED_F1 = 637717      ' Fair Facturacion

'***************************

Private Const Cfg = "FwLg.cfg"
Private Const LPT = "LPT1"
Private lCfgCode As String

Private Const C_IX = 0
Private Const C_PC = 1
Private Const C_COD = 2
Private Const C_MAC = 3

Private Const LASTCOL = C_MAC

Private bLogged As Boolean
Private lClave(2) As String
Private lNivel As Integer
Private lPrd As String

Private Sub Br_Clear_Click()

   Tx_Info = ""
   Call Clean
   
End Sub

Private Sub Bt_Copy2Lic_Click()
   Call Copy2Lic

End Sub

Private Sub Bt_Parse_Click()

   Call Clean
   
   Call Parse(1)
   
   Call FGrVRows(Grid)
   
End Sub

Private Sub Bt_Paste_Click()
   Dim Buf As String

   Call Clean

   Buf = Clipboard.GetText()
   Buf = ReplaceStr(Buf, vbCrLf & vbCrLf, vbCrLf)

   Tx_Info = ""
   Tx_Info = Buf
   
'   Call Parse(0)

End Sub

Private Sub Form_Load()
   Dim Buf As String, i As Integer, PrimerUso As Long, Hasta As Long
         
   Set gFrmMain = Me
   bLogged = False
   
   Call PamInit
      
   Call Rnd(CDbl(Now))
      
   'Call ChkFairware
      
'   Call InitLexComun
   
   Call SetupGrid
   
   lCfgCode = "a" & "x" & Int(657432 + Rnd() * 10000)
   
   ' Rango en que se puede usar la primera vez para que se incriba
   PrimerUso = DateSerial(2021, 6, 3)
   
   Hasta = DateSerial(2022, 12, 15)
   
   If Hasta - PrimerUso < 30 Then
      Debug.Print " *** OJO ****"
   End If
   
   If Now >= PrimerUso Then
      lCfgCode = "a" & "x" & GenClave2(W.PcName & "#" & Hasta, 54321 + Hasta * 3)
      If Now >= PrimerUso And Now < PrimerUso + 5 Then ' Hasta esta fecha se puede usar por primera vez
         Call SetIniString(Cfg, "Port", LPT, lCfgCode)
      ElseIf Now > Hasta Then ' hasta aquí funciona, luego hay que generar otro ejecutable
         Buf = "a" & "x" & Int(874821 + Rnd() * 10000)
         Call SetIniString(Cfg, "Port", LPT, Buf)
      End If
   Else
      Buf = "a" & "x" & Int(834821 + Rnd() * 10000)
      Call SetIniString(Cfg, "Port", LPT, Buf)
   End If
   
   lClave(0) = "L" & "3" & "z"
   lClave(0) = lClave(0) & "-" & 7 & "5" & 1

   lClave(1) = "a" & "K" & "e"
   lClave(1) = lClave(1) & "l" & "a" & "." & Day(Now) * 2
   lClave(1) = UCase(lClave(1))

   lClave(2) = "r" & "I" & "f"
   lClave(2) = lClave(2) & "f" & "o" & "."
   lClave(2) = UCase(lClave(2))

   On Error Resume Next
   MkDir W.AppPath & "\Log"

   If GetIniString(Cfg, "Port", LPT, "") = lCfgCode Then
      Me.Caption = Me.Caption & "."
   End If

   i = 0
   gAppCode.Nivel(i).id = VER_ILIM
   gAppCode.Nivel(i).Desc = "Sin límite de empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_5EMP
   gAppCode.Nivel(i).Desc = "Máximo cinco empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_50EMP
   gAppCode.Nivel(i).Desc = "Máximo cincuenta empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_100EMP
   gAppCode.Nivel(i).Desc = "Máximo cien empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_200EMP
   gAppCode.Nivel(i).Desc = "Máximo doscientas empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_400EMP
   gAppCode.Nivel(i).Desc = "Máximo cuatrocientas empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_800EMP
   gAppCode.Nivel(i).Desc = "Máximo ochocientas empresas"

   SSTab1.Tab = 0

End Sub

Private Sub Parse(ByVal bGen As Boolean)
   Dim Buf As String, i As Long, j As Long, k As Long
   Dim nPC As Integer, PCs As String, Verif As Long, Verif1 As Long, nLic As Integer
   Dim Info As String, Fd As Long, FName As String, Rut As String
   Dim bMsg As Boolean, bFact As Boolean
   
'   bFact = 1    ' Incluye Facturación desde 24 abr 2017
   bFact = gInFairware   ' Incluye Facturación desde 24 abr 2017 - 2 ene 2019: solo en Fairware

   Buf = Trim(Tx_Info)
   If Buf = "" Then
      Exit Sub
   End If

   i = InStr(Buf, "*** NO ")
   If i <= 0 Then
      MsgBox1 "La información está incompleta, no se encuentra '*** NO ' al comienzo de la información.", vbExclamation
      Exit Sub
   End If
   
   Buf = Mid(Buf, i)

   i = InStr(Buf, "*****")
   If i <= 0 Then
      MsgBox1 "La información está incompleta, no se encuentra '*****' al final de la información.", vbExclamation
      Exit Sub
   End If
   
   Buf = Left(Buf, i + 10)

   Buf = ReplaceStr(Buf, " ", "")
   Buf = ReplaceStr(Buf, vbCr, "")
   Buf = ReplaceStr(Buf, vbLf, "")
      
   i = InStr(Buf, "::")
   If i <= 0 Then
      MsgBox1 "La información está incompleta, faltan separadores ::.", vbExclamation
      Exit Sub
   End If
   
   j = InStr(i, Buf, "[", vbBinaryCompare)
   If j <= 0 Then
      MsgBox1 "La información está incompleta, faltan separadores [.", vbExclamation
      Exit Sub
   End If
   
   k = InStr(j, Buf, "]", vbBinaryCompare)
   If k <= 0 Then
      MsgBox1 "La información está incompleta, faltan separadores ].", vbExclamation
      Exit Sub
   End If
   
   ' [ Nivel : nPc : nLic ]
   ' j       i            k
   
   Info = Mid(Buf, i, k - i + 1)
   
   i = InStr(j + 1, Buf, ":", vbBinaryCompare)
   If i < j + 1 Then
      MsgBox1 "La información está incompleta, faltan separadores :.", vbExclamation
      Exit Sub
   End If
      
   lNivel = Val(Mid(Buf, j + 1, i - j - 1))
   nPC = Val(Mid(Buf, i + 1, k - i - 1))
   
   i = InStr(i + 1, Buf, ":", vbBinaryCompare)
   If i > 0 And i < k Then  ' por si viene nLic
      nLic = Val(Mid(Buf, i + 1, k - i - 1))
   Else
      nLic = nPC
   End If
   
   i = InStr(k + 1, Buf, ":", vbBinaryCompare)
   j = InStr(i + 1, Buf, ";", vbBinaryCompare)
   If i > 0 And j > 0 Then
      Rut = Mid(Buf, i + 1, j - i - 1)
   Else
      bMsg = True
   End If
   k = j
   
   lPrd = ""
   i = InStr(k + 1, Buf, ":", vbBinaryCompare)
   j = InStr(i + 1, Buf, ";", vbBinaryCompare)
   If i > 0 And j > 0 Then
      If Mid(Buf, k + 1, i - k - 1) = "PRD" Then
         lPrd = Mid(Buf, i + 1, j - i - 1)
         k = j
      End If
   Else
      bMsg = True
   End If
   
   If lPrd <> "" And StrComp(Left(lPrd, 5), "LpCon", vbTextCompare) And StrComp(Left(lPrd, 5), "LPFac", vbTextCompare) Then
      La_Prod = "???"
      MsgBox1 "Producto no soportado.", vbExclamation
      Exit Sub
   End If
      
   i = InStr(k + 1, Buf, ":", vbBinaryCompare)
   j = InStr(i + 1, Buf, ";", vbBinaryCompare)
   If i > 0 And j > 0 Then
      Verif = Val(Mid(Buf, i + 1, j - i - 1))
   Else
      bMsg = True
   End If

   If bMsg Then
      MsgBox1 "La información está incompleta.", vbExclamation
      Exit Sub
   End If

   Buf = Info & "RUT:" & Rut & ";"

   If lPrd <> "" Then
      Buf = Buf & "PRD:" & lPrd & ";"
   End If

   Verif1 = GenClave(Buf, 54321) Mod 65432
   If Verif <> Verif1 Then
      MsgBox1 "La información fue modificada, el verificador no corresponde.", vbExclamation
      Exit Sub
   End If

   For i = 0 To UBound(gAppCode.Nivel)
      If gAppCode.Nivel(i).id = lNivel Then
         Tx_Nivel = gAppCode.Nivel(i).Desc
         Exit For
      End If
   Next i

   Tx_RUT = Rut
   Tx_nPC = nPC
   Tx_nLic = nLic
   Tx_Verif = Verif

   Call FillEquipos(Info, bGen)

   If nLic > nPC Then
      Tx_nLic.BackColor = VBCOLOR_LIGHTYELLOW2
      MsgBox1 "Atención: la cantidad de licencias es mayor que la cantidad de equipos.", vbInformation
   Else
      Tx_nLic.BackColor = vbWindowBackground
   End If

   If bLogged = False Then
      Tx_Clave = UCase(Tx_Clave)
      For i = 0 To 2
         If Tx_Clave = UCase(lClave(i)) Then
            bFact = (i > 0 And gInFairware)
            Tx_Clave = Space(Len(Tx_Clave))
            bLogged = True ' después de asignar
            Exit For
         End If
      Next i
   End If
   
'   If GetIniString(Cfg, "Port", LPT, "") <> lCfgCode Then
'      bLogged = False
'   End If

   If StrComp(lPrd, "LpContab4", vbTextCompare) = 0 Then
      La_Prod = "LP Contabilidad"
   ElseIf StrComp(lPrd, "LpContabSql", vbTextCompare) = 0 Then
      If gInFairware Then
         La_Prod = "LP Contabilidad SQL"
      Else
         bGen = False
      End If
   ElseIf StrComp(lPrd, "LpFactura", vbTextCompare) = 0 Then
      If gInFairware Then
         La_Prod = "LP Facturación"
      Else
         bGen = False
      End If
   Else
      La_Prod = "LP Contab"
   End If

   If bLogged Then 'And bGen Then
   
      If StrComp(lPrd, "LpContab4", vbTextCompare) = 0 Then
         Tx_NetCode = GenCode(UCase(Buf), PC_SEED_C4)
      ElseIf StrComp(lPrd, "LpContabsql", vbTextCompare) = 0 Then
         Tx_NetCode = GenCode(UCase(Buf), PC_SEED_CS)
      ElseIf bFact And StrComp(lPrd, "LpFactura", vbTextCompare) = 0 Then
         Tx_NetCode = GenCode(UCase(Buf), PC_SEED_F1)
      Else
         Tx_NetCode = GenCode(UCase(Info), PC_SEED)
      End If
      On Error Resume Next
      
      If Ck_Clip.Value Then
         Clipboard.Clear
         Clipboard.SetText Tx_NetCode
      End If
      
      FName = W.AppPath & "\Log\NetCode" & Format(Now, "yyyymm") & ".log"
      Fd = FreeFile
      Open FName For Append Access Write As #Fd
      If Err = 0 Then
         Print #Fd, Format(Now, "d mmm yyyy h:nn") & vbTab & GetComputerName() & vbTab & GetUserName()
         Print #Fd, "RUT: " & Tx_RUT
         Print #Fd, "Versión: " & Tx_Nivel
         Print #Fd, "NetCode: " & Tx_NetCode
         Print #Fd, Tx_Info
         Print #Fd, String(40, "-")
         Close #Fd
      End If
      
   End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
   
      If Ck_Celda Then
         Call SetClipText(Grid.TextMatrix(Grid.Row, Grid.Col))
      Else
         Call FGr2Clip(Grid, "Producto:" & vbTab & La_Prod & vbCrLf & "RUT:" & vbTab & Tx_RUT)
      End If
   End If

End Sub

Private Sub Tx_Clave_Change()
   bLogged = False
End Sub

Private Sub Tx_Info_Change()
   Call Clean
   Call Parse(0)
End Sub


Private Sub FillEquipos(ByVal Info As String, ByVal bGen As Boolean)
   Dim g As Integer, i As Integer, j As Integer, k As Integer, l As Integer, e As Integer, fe As Boolean, ne As Integer
   Dim r As Integer, PC As String, Mac As String, Cod As String, Buf As String, Equipos() As String
      
   ReDim Equipos(10)
   
   ' 21 jun 2016: se ordena la lista de equipos
   
   Cb_Equ.Clear
   
   g = 1
   r = 1
   Do
      i = InStr(g, Info, "::", vbBinaryCompare) ' desde g
      If i <= 0 Then
         Exit Do
      End If
            
      j = InStr(i + 2, Info, ":", vbBinaryCompare)
      k = InStr(j + 1, Info, ":", vbBinaryCompare)
      l = InStr(k + 1, Info, ":", vbBinaryCompare)
      
      PC = Trim(Mid(Info, i + 2, j - i - 2))
      Mac = Trim(Mid(Info, j + 1, k - j - 1))
      Cod = Trim(Mid(Info, k + 1, l - k - 1))

      Call CbAddItem(Cb_Equ, PC & "#" & Cod & "#" & Mac, r)

      fe = False
      For e = 0 To UBound(Equipos)
         If Len(Equipos(e)) = 0 Then
            Exit For
         ElseIf StrComp(PC, Equipos(e), vbTextCompare) = 0 Then
            fe = True
            Exit For
         End If
      Next e
      
      If fe = False Then
         If ne > UBound(Equipos) Then
            ReDim Preserve Equipos(e + 10)
         End If
      
         Equipos(ne) = PC
         ne = ne + 1
      End If

      r = r + 1
      g = l + 1   ' siguiente posicion en el string
   Loop
   
   tx_nPC_Dif = ne
   
   r = Grid.FixedRows
   Grid.rows = r
   For g = 0 To Cb_Equ.ListCount - 1
      
      Grid.rows = r + 1
      Buf = Cb_Equ.List(g)
      
      j = InStr(Buf, "#")
      k = InStr(j + 1, Buf, "#", vbBinaryCompare)
      
      Grid.TextMatrix(r, C_IX) = Cb_Equ.ItemData(g)
      Grid.TextMatrix(r, C_PC) = Trim(Left(Buf, j - 1))
      Grid.TextMatrix(r, C_COD) = Trim(Mid(Buf, j + 1, k - j - 1))
      Grid.TextMatrix(r, C_MAC) = Trim(Mid(Buf, k + 1))

      For j = Grid.FixedRows To r - 1
         If StrComp(Grid.TextMatrix(j, C_PC), Grid.TextMatrix(r, C_PC), vbTextCompare) = 0 Then
         
            If bGen = False Then
               SSTab1.Tab = 1
            End If
            
            Call FGrForeColor(Grid, r, C_PC, vbBlue)
            Call FGrForeColor(Grid, j, C_PC, vbBlue)
            
            If StrComp(Grid.TextMatrix(j, C_MAC), Grid.TextMatrix(r, C_MAC), vbTextCompare) <> 0 Then
               MsgBox1 " El equipo " & Grid.TextMatrix(j, C_PC) & " se repite en las filas " & j & " y " & r & "." & vbCrLf & "No deben haber nombres repetidos.", vbExclamation
            End If
         End If
      Next j

      r = r + 1
   Next g

   Call FGrVRows(Grid)

End Sub

Private Sub SetupGrid()

   Grid.Cols = LASTCOL + 1

   Call FGrSetup(Grid)

   Grid.TextMatrix(0, C_PC) = "PC"
   Grid.TextMatrix(0, C_MAC) = "MAC"
   Grid.TextMatrix(0, C_COD) = "Cód. PC"

   Grid.ColWidth(C_IX) = 500
   Grid.ColWidth(C_PC) = 2300
   Grid.ColWidth(C_COD) = 1500
   Grid.ColWidth(C_MAC) = 1600

   Grid.ColAlignment(C_IX) = flexAlignRightCenter
   Grid.ColAlignment(C_PC) = flexAlignLeftCenter
   Grid.ColAlignment(C_MAC) = flexAlignLeftCenter
   Grid.ColAlignment(C_COD) = flexAlignLeftCenter

End Sub

Private Sub Clean()

   Tx_Nivel = ""
   Tx_nPC = ""
   Tx_nLic = ""
   Tx_Verif = ""
   Tx_NetCode = ""
   Tx_RUT = ""
   Grid.rows = Grid.FixedRows

End Sub

Private Sub Tx_Info_KeyPress(KeyAscii As Integer)

   If KeyAscii <> 22 And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub

Private Sub Copy2Lic()
   Dim Buf As String, r As Integer
   
   If Trim(Tx_NetCode) = "" Then
      MsgBox1 "Falta generar el código de red.", vbExclamation
      Exit Sub
   End If
   
   If lPrd = "" Then
'      Buf = "LegContab2"
      Buf = "uLegContab2"
   Else
      Buf = lPrd
   End If
   
   Buf = "#N:%" & vbTab & Tx_RUT & vbTab & Buf & vbTab & Tx_nPC & vbCrLf
   For r = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(r, C_PC) = "" Then
         Exit For
      End If
      Buf = Buf & Grid.TextMatrix(r, C_PC) & vbTab & Grid.TextMatrix(r, C_COD)
      Buf = Buf & vbTab & Grid.TextMatrix(r, C_MAC) & vbTab & lNivel & vbTab & Tx_Verif & vbTab & Tx_NetCode & vbCrLf
   Next r

   Clipboard.Clear
   Clipboard.SetText Buf

End Sub


