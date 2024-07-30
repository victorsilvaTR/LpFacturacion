VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LP Contabilidad"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8535
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   8325
      Left            =   0
      Picture         =   "FrmAbout.frx":000C
      ScaleHeight     =   8265
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thomson Reuters Chile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   5700
         Width           =   8415
      End
      Begin VB.Label Lb_ubicacion 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   7920
         Width           =   8355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel/Fax: (56 2) 2510 5000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   9
         Top             =   7560
         Width           =   2055
      End
      Begin VB.Label la_Link 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "email  soporte.chile@thomsonreuters.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   5280
         MouseIcon       =   "FrmAbout.frx":1C02A
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Tag             =   "mailto:soporte@legalpublishing.cl?Subject=Legal Publishing%20Contabilidad"
         Top             =   7560
         Width           =   3075
      End
      Begin VB.Label la_Link 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.thomsonreuters.cl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   2640
         MouseIcon       =   "FrmAbout.frx":1C17C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Tag             =   "http://www.legalpublishing.cl"
         Top             =   7560
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Fairware Ltda."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   60
         TabIndex        =   6
         Top             =   6120
         Width           =   8415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Desarrollado por:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Top             =   5400
         Width           =   8415
      End
      Begin VB.Label La_Ver 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 00.00.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label La_Fecha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00 mmm 0000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label La_Nivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   4680
         Width           =   5535
      End
      Begin VB.Label Lb_Demo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEMO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6060
         TabIndex        =   1
         Top             =   4920
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
   Dim Buf As String, Q1 As String
   Dim Dt As Long
   Dim i As Integer
   Dim Rs As Recordset
   
   Me.Icon = FrmMain.Icon
   'Im_Icon.Picture = Me.Icon
   Lb_Demo.Visible = gAppCode.Demo
   
  ' Image1.Picture = Me.Icon
  ' la_Link(0).MouseIcon = FrmMain.Fr_Invivisible.MouseIcon
  ' la_Link(1).MouseIcon = FrmMain.Fr_Invivisible.MouseIcon
   
   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      Buf = "/" & Val(vFld(Rs("Valor")))
   Else
      Buf = ""
   End If
   Call CloseRs(Rs)
   
   'Tx_Ubicacion = "Ubicación: " & W.AppPath
   Lb_ubicacion = "Ubicación: " & W.AppPath
   
   Me.Caption = "Acerca de " & gLexContab
   'La_Title = gLexContab
   
   La_Ver = "Versión " & W.Version & Buf
   
   La_Nivel = gAppCode.NivProd
   For i = 0 To UBound(gAppCode.Nivel)
      If gAppCode.Nivel(i).id = gAppCode.NivProd Then
         La_Nivel = gAppCode.Nivel(i).Desc
      End If
   Next i
   
'   Select Case gAppCode.NivProd
'      Case VMANT_2005
'         La_Nivel = "c/Mant. 2005"
'
'      Case 1:
'         La_Nivel = "Básico"
'
'      Case Else:
'         La_Nivel = "¿" & gAppCode.NivProd & "?"
'
'   End Select
      
   La_Fecha = Format(W.FVersion, "mmm d, yyyy")

End Sub

Private Sub OK_Click()
   Unload Me
End Sub
Private Sub la_Link_Click(Index As Integer)
   Dim Rc As Long
   Dim Buf As String
   
   If la_Link(Index).Tag <> "" Then
      Buf = la_Link(Index).Tag
   Else
      Buf = la_Link(Index)
   End If
   
   Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", SW_SHOWNORMAL)
   
End Sub

