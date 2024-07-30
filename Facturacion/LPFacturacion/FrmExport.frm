VERSION 5.00
Begin VB.Form FrmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a LPContabilidad"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   1140
      TabIndex        =   9
      Top             =   420
      Width           =   2595
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Compras"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2235
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   2235
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   240
      Picture         =   "FrmExport.frx":0000
      ScaleHeight     =   660
      ScaleWidth      =   660
      TabIndex        =   8
      Top             =   540
      Width           =   660
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   1140
      TabIndex        =   5
      Top             =   1980
      Width           =   4935
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   7
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.CommandButton Bt_Exportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   540
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   1020
      Width           =   1575
   End
End
Attribute VB_Name = "FrmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Exportar_Click()
   Dim Mes As Long
   
   Mes = DateSerial(CbItemData(Cb_Ano), CbItemData(Cb_Mes), 1)
   
   
   If MsgBox1("ATENCIÓN: Sólo se exportarán los DTE que se encuentren en estado EMITIDO." & vbCrLf & vbCrLf & "Asegúrese de haber revisado y actualizado el estado de cada DTE, ingresando a DTE Emitidos >> Ver Detalle Estado" & vbCrLf & "antes de realizar esta operación. " & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

   If Op_Libros(LIB_VENTAS) Then
      Call ExportDTE(LIB_VENTAS, Mes)
   Else
      Call ExportDTE(LIB_COMPRAS, Mes)
   End If
   
End Sub

Private Sub Form_Load()

   Op_Libros(LIB_VENTAS) = 1
   Call CbFillMes(Cb_Mes, Month(Now))

   Call CbFillAno(Cb_Ano, Year(Now), Year(Now) - 5, Year(Now))
      
End Sub

Private Function ExportDTE(ByVal TipoLib, ByVal Mes As Long) As Boolean
   Dim DbName As String
   Dim Db As Database
   Dim LibExpName As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim StrMes As String
   Dim CreateEnable As Boolean
   Dim n As Long
   Dim ExpDbPath As String
   Dim Msg As String
   Dim StrAño As String
   Dim i As Integer
   Dim ChkSumCuentas As Long
   Dim Tbl As TableDef
   Dim fld As Field
    
   If TipoLib = 0 Or Mes = 0 Then
      Exit Function
   End If
   
   ExportDTE = False
   
   StrMes = Year(Mes) & Right("0" & Month(Mes), 2)
   StrAño = Year(Mes)
   
   'Creamos el nombre de la DB de exportación: "Libro-AñoMes.mdb"
   LibExpName = ReplaceStr(gTipoLib(TipoLib), "Libro de ", "")
   LibExpName = "-" & UCase(Left(LibExpName, 3))
   LibExpName = LibExpName & "-" & StrMes
    
   ExpDbPath = gExportPath & "\" & StrAño
   
   On Error Resume Next
   MkDir (ExpDbPath)
   On Error GoTo 0
       
   If Err Then
      MsgBox1 "Error " & Err & ": " & Error & " al momento de crear la carpeta de exportación.", vbExclamation
      Exit Function
   End If
   
   DbName = ExpDbPath & "\" & gEmpresa.Rut & LibExpName & ".mdb"

   CreateEnable = LockAction(DbMain, LK_EXPLIBROS, Mes)
   
   If CreateEnable = False Then    'alguien más está exportando este mes
      MsgBox1 "Esta operación ya se está realizando en el equipo '" & IsLockedAction(DbMain, LK_EXPLIBROS, Mes) & "'. No se realizará la exportación.", vbInformation
      Exit Function
   End If
   
   On Error Resume Next
   
   Kill (DbName)
   Err.Clear
   
   'creamos la DB
   Set Db = CreateDatabase(DbName, dbLangGeneral)
      
   If (Err Or Db Is Nothing) And Err <> 3204 Then
      MsgBox "Error " & Err & ", " & Error & NL & DbName, vbExclamation
      Db.Close
      Set Db = Nothing
      Exit Function
   End If
   
   On Error GoTo 0

   'linkeamos a la DB las tablas que necesitamos
   On Error Resume Next
   
   If TipoLib = LIB_VENTAS Then
      Call LinkMdbTable(Db, DbMain.Name, "DTE", , , False, gEmpresa.ConnStr)
      Call LinkMdbTable(Db, DbMain.Name, "DetDTE", , , False, gEmpresa.ConnStr)
   Else
      Call LinkMdbTable(Db, DbMain.Name, "DTERecibidos", , , False, gEmpresa.ConnStr)
   End If
   
   Call LinkMdbTable(Db, DbMain.Name, "Entidades", , , False, gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "Empresa", , , False, gEmpresa.ConnStr)
   'Call LinkMdbTable(Db, DbMain.Name, "TipoDocs", "TipoDocsDTE", , False, gEmpresa.ConnStr)
   Call LinkMdbTable(Db, gDbPath & "\TRFactura.mdb", "TipoDocs", "TipoDocsDTE", , False, gComunConnStr)
'   Call LinkMdbTable(Db, DbMain.Name, "AreaNegocio", , , False, gEmpresa.ConnStr)
'   Call LinkMdbTable(Db, DbMain.Name, "CentroCosto", , , False, gEmpresa.ConnStr)
   
   If Err > 0 Then
      MsgBox1 "Error al generar el archivo de exportación. Asegúrese que éste no esté abierto por otro usuario.", vbExclamation
      Exit Function
   End If
      
   On Error GoTo 0
   
   'Insertamos la empresa seleccionada
   Q1 = "SELECT Empresa.* INTO EmpresaDTESel "
   Q1 = Q1 & " FROM Empresa WHERE Id = " & gEmpresa.Id
   Call ExecSQL(Db, Q1)
   
   Q1 = "SELECT TipoDocsDTE.* INTO TipoDocs "
   Q1 = Q1 & " FROM TipoDocsDTE "
   Call ExecSQL(Db, Q1)
   
   
   If TipoLib = LIB_VENTAS Then
      'generamos los registros del libro-mes de ventas
      Q1 = "SELECT DTE.* "
      Q1 = Q1 & " INTO DTE_" & StrMes
      Q1 = Q1 & " FROM DTE LEFT JOIN Entidades ON DTE.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & " WHERE Year(Fecha) = " & Year(Mes) & " AND Month(Fecha) = " & Month(Mes)
      Q1 = Q1 & " AND DTE.TipoLib = " & TipoLib & " AND Folio > 0 AND idEstado = " & EDTE_EMITIDO & " AND DTE.IdEmpresa = " & gEmpresa.Id
         
      Call ExecSQL(Db, Q1)
         
      'Insertamos los detalles de los documentos de venta seleccionados
      Q1 = "SELECT DetDTE.*"
      Q1 = Q1 & " INTO DetDTE_" & StrMes
      Q1 = Q1 & " FROM DetDTE INNER JOIN DTE_" & StrMes & " ON DetDTE.IdDTE = DTE_" & StrMes & ".IdDTE"
      Call ExecSQL(Db, Q1)
      
      'Insertamos las entidades asociadas a los docs seleccionados
      Q1 = "SELECT DISTINCT Entidades.* INTO Entidades_" & StrMes
      Q1 = Q1 & " FROM Entidades INNER JOIN DTE_" & StrMes & " ON Entidades.IdEntidad = DTE_" & StrMes & ".IdEntidad "
      Call ExecSQL(Db, Q1)
      
      Q1 = "SELECT Count(*) FROM DTE_" & StrMes
      
      
   
   Else
   
      'insertamos los documentos recibidos (compras)
      Q1 = "SELECT DTERecibidos.* "
      Q1 = Q1 & " INTO DTERecibidos_" & StrMes
      Q1 = Q1 & " FROM DTERecibidos LEFT JOIN Entidades ON DTERecibidos.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & " WHERE Year(FEmision) = " & Year(Mes) & " AND Month(FEmision) = " & Month(Mes)
      Q1 = Q1 & " AND DTERecibidos.TipoLib = " & TipoLib & " AND Folio > 0 AND DTERecibidos.IdEmpresa = " & gEmpresa.Id
'      Q1 = Q1 & " AND idEstado = " & EDTE_EMITIDO
         
      Call ExecSQL(Db, Q1)
      
      'Insertamos las entidades asociadas a los docs seleccionados
      Q1 = "SELECT DISTINCT Entidades.* INTO Entidades_" & StrMes
      Q1 = Q1 & " FROM Entidades INNER JOIN DTERecibidos_" & StrMes & " ON Entidades.IdEntidad = DTERecibidos_" & StrMes & ".IdEntidad "
      Call ExecSQL(Db, Q1)
   
      Q1 = "SELECT Count(*) FROM DTERecibidos_" & StrMes
   
   End If
   
   
   'vemos cuántos docs se exportaron
   Set Rs = OpenRs(Db, Q1)
   n = Rs(0)
   Call CloseRs(Rs)
   
   Select Case n
      Case 0
         Msg = "No se encontraron documentos para exportar."
      Case 1
         Msg = "Se exportó un Documento Electrónico." & vbNewLine & vbNewLine
         Msg = Msg & "Archivo generado:" & vbNewLine & vbNewLine
         Msg = Msg & "      " & DbName
      Case Else
         Msg = "Se exportaron " & n & " Documentos Electrónicos." & vbNewLine & vbNewLine
         Msg = Msg & "Archivo generado:" & vbNewLine & vbNewLine
         Msg = Msg & "      " & DbName
   End Select
   
   
   Call UnLinkTable(Db, "DTE")
   Call UnLinkTable(Db, "DetDTE")
   Call UnLinkTable(Db, "Entidades")
   
   Call CloseDb(Db)
   
   Call UnLockAction(DbMain, LK_EXPLIBROS, Mes)
   
   MsgBox1 Msg, vbInformation + vbOKOnly
   
   ExportDTE = True
   
End Function

