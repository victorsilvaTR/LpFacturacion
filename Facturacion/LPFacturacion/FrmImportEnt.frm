VERSION 5.00
Begin VB.Form FrmImportEnt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Entidades desde Contabilidad"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Operación"
      Height          =   855
      Left            =   1200
      TabIndex        =   6
      Top             =   1740
      Width           =   5895
      Begin VB.CheckBox Ch_Update 
         Caption         =   "Actualizar entidades ya existentes"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox Ch_ImpNew 
         Caption         =   "Importar entidades nuevas"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo de la Empresa-Año en LPContabilidad"
      Height          =   1095
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   9315
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   7800
         Picture         =   "FrmImportEnt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   7455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   240
      Picture         =   "FrmImportEnt.frx":056A
      ScaleHeight     =   690
      ScaleWidth      =   690
      TabIndex        =   2
      Top             =   540
      Width           =   690
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   1980
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   1980
      Width           =   1395
   End
End
Attribute VB_Name = "FrmImportEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFName As String
Dim lFTitle As String

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Importar_Click()
   If Ch_ImpNew = 0 And Ch_Update = 0 Then
      MsgBox1 "Seleccione la operación que desea realizar.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   Call ImportEnt
End Sub

Private Function ImportEnt() As Boolean
   Dim LpDbName As String
   Dim ImpName As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim n1 As Long, n2 As Integer
   Dim Msg As String
   Dim i As Integer
   Dim Tbl As TableDef
   Dim fld As Field
   Dim CreateEnable As Boolean
   Dim RutFName As String
   Dim Idx As Integer
    
   ImportEnt = False
   
   If lFName = "" Or lFTitle = "" Then
      Exit Function
   End If
      
   Idx = InStr(lFTitle, "-")
   If Idx > 0 Then
      RutFName = Left(lFTitle, Idx - 1)
      If gEmpresa.Rut <> RutFName Then
         If MsgBox1("El RUT del archivo a importar no coincide con el RUT de esta empresa." & vbCrLf & vbCrLf & "¿Desea continuar de todas maneras?", vbYesNo + vbQuestion) = vbNo Then
            Exit Function
         End If
      End If
   End If
   
   LpDbName = lFName

   CreateEnable = LockAction(DbMain, LK_IMPENTIDADES, 0)
   
   If CreateEnable = False Then    'alguien más está importando entidades
      MsgBox1 "Esta operación ya se está realizando en el equipo '" & IsLockedAction(DbMain, LK_IMPENTIDADES, 0) & "'. No se realizará la importación.", vbInformation
      Exit Function
   End If
   

   'linkeamos a la DB las tablas que necesitamos
   Call LinkMdbTable(DbMain, LpDbName, "Entidades", "LP_Entidades", , , gEmpresa.ConnStr)
   
   'vemos cuantas entidades hay ahora
   Q1 = "SELECT Count(*) FROM Entidades"
   Set Rs = OpenRs(DbMain, Q1)
   n1 = Rs(0)
   Call CloseRs(Rs)
   
   If Ch_Update <> 0 Then
      'Primero actualizamos la entidades ya existentes
      Q1 = "UPDATE Entidades INNER JOIN LP_Entidades"
      Q1 = Q1 & " ON Entidades.RUT = LP_Entidades.RUT "
      Q1 = Q1 & " SET Entidades.Codigo = LP_Entidades.Codigo "
      Q1 = Q1 & ", Entidades.Nombre = LP_Entidades.Nombre "
      Q1 = Q1 & ", Entidades.Direccion = LP_Entidades.Direccion "
      Q1 = Q1 & ", Entidades.Region = LP_Entidades.Region "
      Q1 = Q1 & ", Entidades.Comuna = LP_Entidades.Comuna "
      Q1 = Q1 & ", Entidades.Ciudad = LP_Entidades.Ciudad "
      Q1 = Q1 & ", Entidades.Telefonos = LP_Entidades.Telefonos "
      Q1 = Q1 & ", Entidades.Fax = LP_Entidades.Fax "
      Q1 = Q1 & ", Entidades.ActEcon = LP_Entidades.ActEcon "
      Q1 = Q1 & ", Entidades.CodActEcon = LP_Entidades.CodActEcon "
      Q1 = Q1 & ", Entidades.DomPostal = LP_Entidades.DomPostal "
      Q1 = Q1 & ", Entidades.ComPostal = LP_Entidades.ComPostal "
      Q1 = Q1 & ", Entidades.Email = LP_Entidades.Email "
      Q1 = Q1 & ", Entidades.Web = LP_Entidades.Web "
      Q1 = Q1 & ", Entidades.Estado = LP_Entidades.Estado "
      Q1 = Q1 & ", Entidades.Obs = LP_Entidades.Obs "
      Q1 = Q1 & ", Entidades.Clasif0 = LP_Entidades.Clasif0 "
      Q1 = Q1 & ", Entidades.Clasif1 = LP_Entidades.Clasif1 "
      Q1 = Q1 & ", Entidades.Clasif2 = LP_Entidades.Clasif2 "
      Q1 = Q1 & ", Entidades.Clasif3 = LP_Entidades.Clasif3 "
      Q1 = Q1 & ", Entidades.Clasif4 = LP_Entidades.Clasif4 "
      Q1 = Q1 & ", Entidades.Clasif5 = LP_Entidades.Clasif5 "
      Q1 = Q1 & ", Entidades.Giro = LP_Entidades.Giro "
      Q1 = Q1 & ", Entidades.NotValidRut = LP_Entidades.NotValidRut "
      Q1 = Q1 & ", Entidades.EsSupermercado = LP_Entidades.EsSupermercado "
      Q1 = Q1 & ", Entidades.EntRelacionada = LP_Entidades.EntRelacionada "
      Call ExecSQL(DbMain, Q1)
    
      MsgBox1 "Se actualizaron las entidades que pudieran haber cambiado.", vbInformation + vbOKOnly
   End If
     
   If Ch_ImpNew <> 0 Then
      'Segundo Insertamos Todas las entidades que no existen acá
'      Q1 = "INSERT INTO Entidades SELECT LP_Entidades.*, " & gEmpresa.id & " as IdEmpresa "
      Q1 = "INSERT INTO Entidades SELECT " & gEmpresa.Id & "  as IdEmpresa, RUT, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal, ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut, EsSupermercado, EntRelacionada "
      Q1 = Q1 & " FROM LP_Entidades "
      Call ExecSQL(DbMain, Q1)
      
      
      'vemos cuantas entidades hay después de la importación
      Q1 = "SELECT Count(*) FROM Entidades"
      Set Rs = OpenRs(DbMain, Q1)
      n2 = Rs(0)
      Call CloseRs(Rs)
   
      Select Case n2 - n1
         Case 0
            Msg = "No se encontraron entidades nuevas para importar."
         Case 1
            Msg = "Se importó una entidad nueva." & vbNewLine & vbNewLine
         Case Else
            Msg = "Se importaron " & n2 - n1 & " entidades nuevas." & vbNewLine & vbNewLine
      End Select
      
      MsgBox1 Msg, vbInformation + vbOKOnly
   End If
   
   Call UnLinkTable(DbMain, "LP_Entidades")
      
   Call UnLockAction(DbMain, LK_IMPENTIDADES, 0)
   
   
   ImportEnt = True
   
End Function

Private Sub Bt_Browse_Click()
   Dim RutFName As String
   Dim Idx As Integer
   Dim FTit As String

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos Access (*.mdb)|*.mdb"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If Err = cdlCancel Then
      Exit Sub
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      lFName = ""
      lFTitle = ""
      Exit Sub
   End If
   Err.Clear
   
   lFName = FrmMain.Cm_ComDlg.Filename
   
   lFTitle = FrmMain.Cm_ComDlg.FileTitle
   
   Tx_FName = lFName
   
   DoEvents
      
End Sub

Private Sub Form_Load()

   Ch_ImpNew = 1
   Ch_Update = 1
   
End Sub
