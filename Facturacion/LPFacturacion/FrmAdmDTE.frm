VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAdmDTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Documentos Electrónicos Emitidos"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_VerDetEstado 
      Caption         =   "Ver y Actualizar Estado DTE..."
      Height          =   1035
      Left            =   11760
      Picture         =   "FrmAdmDTE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Revisar y Actualizar Estado DTE Seleccionado"
      Top             =   780
      Width           =   1515
   End
   Begin VB.TextBox Tx_CurCel 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   9240
      Width           =   7995
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   17
      Top             =   660
      Width           =   11535
      Begin VB.CommandButton Bt_SelFechaDesde 
         Height          =   315
         Left            =   4680
         Picture         =   "FrmAdmDTE.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaDesde 
         Height          =   315
         Left            =   3420
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Bt_SelFechaHasta 
         Height          =   315
         Left            =   6840
         Picture         =   "FrmAdmDTE.frx":06B7
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaHasta 
         Height          =   315
         Left            =   5580
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   735
         Left            =   10140
         Picture         =   "FrmAdmDTE.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   300
         Width           =   1155
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   8280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1635
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox Tx_Folio 
         Height          =   315
         Left            =   660
         TabIndex        =   4
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox Tx_RUT 
         Height          =   315
         Left            =   660
         MaxLength       =   12
         TabIndex        =   1
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox Tx_RazonSocial 
         Height          =   315
         Left            =   3420
         TabIndex        =   2
         ToolTipText     =   "Ingrese cualquier parte de la Razón Social"
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde:"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   29
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "hasta:"
         Height          =   195
         Index           =   2
         Left            =   5100
         TabIndex        =   28
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Index           =   6
         Left            =   7440
         TabIndex        =   22
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   5
         Left            =   7440
         TabIndex        =   21
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   20
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social: "
         Height          =   195
         Index           =   3
         Left            =   2340
         TabIndex        =   18
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   13395
      Begin VB.CommandButton Bt_CopiarURL 
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
         Left            =   1020
         Picture         =   "FrmAdmDTE.frx":0C7C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Copiar URL del DTE seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_DetEstadoDTE 
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
         Left            =   540
         Picture         =   "FrmAdmDTE.frx":1072
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Revisar y Actualizar Estado DTE Seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Del 
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
         Left            =   60
         Picture         =   "FrmAdmDTE.frx":14D7
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar DTE Erróneo Seleccionado "
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Print 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         Picture         =   "FrmAdmDTE.frx":18D3
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Picture         =   "FrmAdmDTE.frx":1D8D
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Select 
         Caption         =   "Seleccionar"
         Height          =   315
         Left            =   10620
         TabIndex        =   14
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   12000
         TabIndex        =   15
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Calendar 
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
         Left            =   2100
         Picture         =   "FrmAdmDTE.frx":21D2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calc 
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
         Left            =   1680
         Picture         =   "FrmAdmDTE.frx":25FB
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Calculadora"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7275
      Left            =   60
      TabIndex        =   0
      Top             =   1920
      Width           =   13460
      _ExtentX        =   23733
      _ExtentY        =   12832
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "FrmAdmDTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDTE = 0
Const C_RUT = 1
Const C_RSOCIAL = 2
Const C_CODDOCSII = 3
Const C_TIPODOC = 4
Const C_DIMINUTIVO = 5
Const C_FOLIO = 6
Const C_FECHA = 7
Const C_LNGFECHA = 8
Const C_TOTAL = 9   'Con IVA
Const C_ESTADODTE = 10
Const C_IDESTADODTE = 11
Const C_IDESTADOSII = 12
Const C_ERRORSII = 13
Const C_RESPUESTASII = 14
Const C_GLOSASII = 15
Const C_TRACKID = 16
Const C_VERPDF = 17
Const C_VERPDFCED = 18
Const C_USUARIO = 19
Const C_URLDTE = 20
Const C_NULL = 21             'para evitar que al ampliar última columna aparezca la URL

Const NCOLS = C_NULL

Dim lOrdenGr(NCOLS) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual


Dim lTipoLib As Integer

'Para seleccionar documento que se va a copiar
Dim lCodDocSII As String

Dim lOper As Integer
Dim lIdDTE As Long
Dim lRc As Integer
Dim lDiminutivoDoc As String        'diminutivo Doc que se va a emitir
Dim lDiminutivoDocRef As String     'diminutivo del doc de referencia selecionado
Dim lOrientacion As Integer
Dim lCodRef As Integer
Dim lNotNotasCredDeb As Boolean
Dim lEsNotaCredDebFactCompra As Boolean

Dim FirstActivate As Boolean

Public Function FView()
   lOper = O_VIEW
   Me.Show vbModal
End Function

Public Function FSelect(ByVal DiminutivoDoc As String, ByVal CodRef As Integer, IdDTE As Long, DiminutivoDocRef As String, ByVal EsNotaCredDebFactCompra As Boolean) As Integer
   lOper = O_SELECT
   lDiminutivoDoc = DiminutivoDoc
   lEsNotaCredDebFactCompra = EsNotaCredDebFactCompra
   lCodRef = CodRef
   
   Me.Show vbModal
   IdDTE = lIdDTE
   DiminutivoDocRef = lDiminutivoDocRef
   FSelect = lRc
End Function
Public Function FSelectCopy(ByVal CodDocSII As String, IdDTE As Long) As Integer
   lOper = O_SELCOPY
   lCodDocSII = CodDocSII
   
   lNotNotasCredDeb = True  'no debe permitir seleccionar notas de crédito o débito de cualquier tipo
   
   Me.Show vbModal
   IdDTE = lIdDTE
   FSelectCopy = lRc
End Function

Private Sub Bt_Buscar_Click()
   Me.MousePointer = vbHourglass
   LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_Calc_Click()
   Call Calculadora

End Sub

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing

End Sub

Private Sub Bt_Cerrar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_CopiarURL_Click()
   Dim Clip As String
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Grid.Row, C_IDDTE)) = 0 Then
      Exit Sub
   End If
   
   Clip = Trim(Grid.TextMatrix(Grid.Row, C_URLDTE))
   
   If Clip = "" Then
      MsgBox1 "No se encontró URL para el DTE seleccionado.", vbExclamation
      Exit Sub
   End If
   
   Call SetClipText(Clip)
   

End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Filtros As String
   Dim wVerPdf As Integer
   Dim wUrlDTe As Integer
   

   If Trim(Cb_Estado) <> "" Then
      Filtros = vbTab & "Estado: " & Cb_Estado
   End If

   If Trim(Cb_TipoDoc) <> "" Then
      Filtros = Filtros & vbTab & "Tipo Doc.: " & Cb_TipoDoc
   End If

   Grid.Redraw = False

   wVerPdf = Grid.ColWidth(C_VERPDF)
   wUrlDTe = Grid.ColWidth(C_URLDTE)
   
   Grid.ColWidth(C_VERPDF) = 0
   Grid.ColWidth(C_URLDTE) = wVerPdf
   Grid.TextMatrix(0, C_URLDTE) = "URL DTE"
   
   Call FGr2Clip(Grid, "Documentos Electrónicos Emitidos" & Filtros)
   
   Grid.ColWidth(C_VERPDF) = wVerPdf
   Grid.ColWidth(C_URLDTE) = 0
   Grid.TextMatrix(0, C_URLDTE) = ""
   
   Grid.Redraw = True
      
End Sub


Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdDTE As Long
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Row <> Grid.RowSel Then
      MsgBox1 "Debe eliminar un documento a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.RowHeight(Row) = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   IdDTE = Val(Grid.TextMatrix(Row, C_IDDTE))
   If IdDTE = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If Val(Grid.TextMatrix(Row, C_IDESTADODTE)) > 0 And Val(Grid.TextMatrix(Row, C_IDESTADODTE)) <> EDTE_ERROR And Val(Grid.TextMatrix(Row, C_FOLIO)) > 0 Then
      If MsgBox1("Este DTE ya que ha sido enviado al SII." & vbCrLf & vbCrLf & "¿Está seguro que lo desea eliminar, bajo su responsabilidad?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
   ElseIf MsgBox1("¿Está seguro que desea borrar este DTE?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
      
   End If
      
   Q1 = "DELETE * FROM DTE WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   Q1 = "DELETE * FROM DetDTE WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   Q1 = "DELETE * FROM Referencias WHERE IdDTE = " & IdDTE & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
   
   Grid.RowHeight(Row) = 0
   
   ' 2 oct 2020: se deja registro
   Call AddLog("DelDTE: Se elimina el DTE: " & Grid.TextMatrix(Row, C_DIMINUTIVO) & " " & Grid.TextMatrix(Row, C_FOLIO) & " RUT " & Grid.TextMatrix(Row, C_RUT) & " " & Grid.TextMatrix(Row, C_RSOCIAL))
    
End Sub

Private Sub Bt_DetEstadoDTE_Click()
   Call Bt_VerDetEstado_Click
End Sub

Private Sub Bt_Print_Click()

   If SelPrinter() Then
      Exit Sub
   End If

               
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = lOrientacion
   
   Call ResetPrtBas(gPrtReportes)
   MousePointer = vbDefault

End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(0) As String
   Dim Encabezados(3) As String
   
   lOrientacion = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   gPrtReportes.Titulos = Titulos
    
   i = 0
   If Tx_RUT <> "" Then
      Encabezados(i) = "RUT entidad:" & vbTab & Tx_RUT
      i = i + 1
   End If
   If Tx_FechaDesde <> "" Then
      Encabezados(i) = "Rango Fechas:" & vbTab & Tx_FechaDesde & " - " & Tx_FechaHasta
      i = i + 1
   End If
   If CbItemData(Cb_TipoDoc) > 0 Then
      Encabezados(i) = "Tipo Doc:" & vbTab & Cb_TipoDoc
      i = i + 1
   End If
   If CbItemData(Cb_Estado) > 0 Then
      Encabezados(i) = "Estado:" & vbTab & Cb_Estado
      i = i + 1
   End If
   
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   ColWi(C_ESTADODTE) = 1200
   ColWi(C_VERPDF) = 0
   ColWi(C_VERPDFCED) = 0
               
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDDTE
   gPrtReportes.NTotLines = 0
   

End Sub

Private Sub Bt_Select_Click()
   Dim Row As Integer
   Dim IdDTE As Long
   
   Row = Grid.Row

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   lDiminutivoDocRef = Grid.TextMatrix(Row, C_DIMINUTIVO)
   
   If lDiminutivoDoc = "NCV" Then
   
      If lEsNotaCredDebFactCompra Then
         If lDiminutivoDocRef <> "FCV" Then
            MsgBox1 "Documento inválido para una Nota de Crédito de una Factura de Compra.", vbExclamation
            Exit Sub
         End If
         
      ElseIf lDiminutivoDocRef <> "FAV" And lDiminutivoDocRef <> "FVE" And lDiminutivoDocRef <> "NDV" Then
         MsgBox1 "Documento inválido para una Nota de Crédito.", vbExclamation
         Exit Sub
      End If
         
   ElseIf lDiminutivoDoc = "NCE" Then
   
      If lDiminutivoDocRef <> "EXP" Or lDiminutivoDocRef <> "NDE" Then
         MsgBox1 "Documento inválido para una Nota de Crédito de Exportación.", vbExclamation
         Exit Sub
      End If
      
   ElseIf lDiminutivoDoc = "NDV" Then
      If lCodRef = REF_CORRIGEMONTOS Then
         If lEsNotaCredDebFactCompra Then
            If lDiminutivoDocRef <> "FCV" Then
               MsgBox1 "Documento inválido para una Nota de Débito de Factura de Compra.", vbExclamation
               Exit Sub
            End If
            
         ElseIf lDiminutivoDocRef <> "FAV" And lDiminutivoDocRef <> "FVE" And lDiminutivoDocRef <> "NCV" Then
            MsgBox1 "Documento inválido para una Nota de Débito.", vbExclamation
            Exit Sub
         End If
         
      ElseIf lCodRef = REF_ANULA Then
         If lDiminutivoDocRef <> "NCV" Then
            MsgBox1 "Documento inválido para una Nota de Débito de anulación de Nota de Crédito.", vbExclamation
            Exit Sub
         End If
      End If
   
   ElseIf lDiminutivoDoc = "NDE" Then
      If lCodRef = REF_CORRIGEMONTOS Then
         If lDiminutivoDocRef <> "EXP" Or lDiminutivoDocRef <> "NCE" Then
            MsgBox1 "Documento inválido para una Nota de Crédito de Exportación.", vbExclamation
            Exit Sub
         End If
      ElseIf lCodRef = REF_ANULA Then
         If lDiminutivoDocRef <> "NDE" Then
            MsgBox1 "Documento inválido para una Nota de Débito de Exportación de anulación de Nota de Crédito de Exportación.", vbExclamation
            Exit Sub
         End If
      End If
   
   End If
   
   If lOper = O_SELECT Then
      If Grid.TextMatrix(Grid.Row, C_IDESTADODTE) <> EDTE_EMITIDO Then
         MsgBox1 "Este documento no está en estado Emitido.", vbExclamation
         Exit Sub
      End If
   ElseIf lOper = O_SELCOPY Then
      If Grid.TextMatrix(Grid.Row, C_IDESTADODTE) <> EDTE_ENVIADO And Grid.TextMatrix(Grid.Row, C_IDESTADODTE) <> EDTE_EMITIDO Then
         MsgBox1 "Este documento no está en estado Enviado o Emitido.", vbExclamation
         Exit Sub
      End If
   End If
   
   
   IdDTE = Val(Grid.TextMatrix(Row, C_IDDTE))
   
   If IdDTE = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   lIdDTE = IdDTE
   lRc = vbOK
   
   Unload Me
End Sub

Private Sub Bt_VerDetEstado_Click()
'   Dim Frm As FrmDetEstadoDTE
   Dim TrackID As String
   Dim Row As Integer
   Dim idEstado As Integer, TxtEstado As String
   Dim trazaEvento As AcpTrazaEvento_t

   Row = Grid.Row
   
   If Row < Grid.FixedRows Or Val(Grid.TextMatrix(Row, C_IDDTE)) = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Bt_VerDetEstado.Enabled = False
   MousePointer = vbHourglass
   DoEvents
   
   If gConectData.Proveedor = PROV_ACEPTA Then
   
      If Grid.TextMatrix(Row, C_URLDTE) = "" Then
         MsgBox1 "No se encuentra disponible la información del Estado del DTE.", vbExclamation
      Else
         trazaEvento.Folio = Grid.TextMatrix(Row, C_FOLIO)
         trazaEvento.TipoDTE = Grid.TextMatrix(Row, C_CODDOCSII)
         trazaEvento.RUTRecep = Replace(Grid.TextMatrix(Row, C_RUT), ".", "")
         trazaEvento.RutEmisor = gEmpresa.Rut & "-" & DV_Rut(gEmpresa.Rut)
         trazaEvento.canal = "EMITIDO"
         
         If W.InDesign Then
           trazaEvento.RutEmisor = "77049060-K"
        End If
         Call AcpShowEstadoDTE(Val(Grid.TextMatrix(Row, C_IDDTE)), Grid.TextMatrix(Row, C_URLDTE), idEstado, TxtEstado, trazaEvento)
         If idEstado <> 0 Then
            Grid.TextMatrix(Row, C_IDESTADODTE) = idEstado
            Grid.TextMatrix(Row, C_ESTADODTE) = IIf(TxtEstado = "", gEstadoDTE(idEstado), TxtEstado)
         End If
      End If
      
   End If
   
   Bt_VerDetEstado.Enabled = True
   MousePointer = vbDefault
   
End Sub


Private Sub Form_Activate()

   If Not FirstActivate Then
      MsgBox1 "Recuerde revisar y actualizar el estado de los DTE ingresando a Ver Detalle Estado." & vbCrLf & vbCrLf & "De esta manera evitará problemas en la emisión de documentos." & vbCrLf & vbCrLf & "El SII puede demorar varios minutos en Acpetar o Rechazar un DTE, por lo cual el Estado Final del documento puede demorar en actualizarse.", vbInformation
      FirstActivate = True
   End If
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim FDesde As Long
   
   lTipoLib = LIB_VENTAS

   Call SetUpGrid
   
   If lOper = O_VIEW Then
      Bt_Select.Visible = False
   Else
      Bt_Cerrar.Caption = "Cancelar"
   End If
      
   If lOper <> O_SELCOPY Then
      Call CbAddItem(Cb_TipoDoc, "", 0, True)
   End If
   
   For i = 0 To UBound(gTipoDocDTE)
      If gTipoDocDTE(i).IdxTipoDoc = 0 Then
         Exit For
      End If
      
      If lOper = O_SELECT Then
         If lDiminutivoDoc = "NCV" Or lDiminutivoDoc = "NDV" Then
            If gTipoDocDTE(i).Diminutivo = "FAV" Or gTipoDocDTE(i).Diminutivo = "FVE" Or gTipoDocDTE(i).Diminutivo = "NDV" Or gTipoDocDTE(i).Diminutivo = "FCV" Then
               Call CbAddItem(Cb_TipoDoc, gTipoDocDTE(i).Nombre, Val(gTipoDocDTE(i).CodDocDTESII))
            End If
         
         ElseIf lDiminutivoDoc = "NCE" Or lDiminutivoDoc = "NDE" Then
            If gTipoDocDTE(i).Diminutivo = "EXP" Then
               Call CbAddItem(Cb_TipoDoc, gTipoDocDTE(i).Nombre, Val(gTipoDocDTE(i).CodDocDTESII))
            End If
         End If
         
      ElseIf lOper = O_SELCOPY Then
         If gTipoDocDTE(i).CodDocDTESII = lCodDocSII Then
            Call CbAddItem(Cb_TipoDoc, gTipoDocDTE(i).Nombre, Val(gTipoDocDTE(i).CodDocDTESII))
            Cb_TipoDoc.ListIndex = 0
         End If
      Else     'lOper <> O_SELECT AND lOper <> O_SELCOPY
         Call CbAddItem(Cb_TipoDoc, gTipoDocDTE(i).Nombre, Val(gTipoDocDTE(i).CodDocDTESII))
      
      End If
   Next i
      
   Call CbAddItem(Cb_Estado, "", 0, True)
   
   For i = 1 To UBound(gEstadoDTE)
      If gEstadoDTE(i) = "" Then
         Exit For
      End If
      
      Call CbAddItem(Cb_Estado, gEstadoDTE(i), i)
   Next i
     
   FDesde = DateAdd("m", -3, Now)
   Call SetTxDate(Tx_FechaDesde, DateSerial(Year(FDesde), Month(FDesde), 1))
   Call SetTxDate(Tx_FechaHasta, Now)
     
   
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_FECHA) = "Fecha Desc, IdDTE Desc"
   
   lOrdenGr(C_RUT) = "Entidades.RUT, " & lOrdenGr(C_FECHA)
   lOrdenGr(C_RSOCIAL) = "Entidades.Nombre, " & lOrdenGr(C_FECHA)
   lOrdenGr(C_TIPODOC) = "TipoDoc, " & lOrdenGr(C_FECHA)
   lOrdenGr(C_FOLIO) = "Folio, " & lOrdenGr(C_FECHA)
   lOrdenGr(C_TOTAL) = "Total, " & lOrdenGr(C_FECHA)
   lOrdenGr(C_ESTADODTE) = "IdEstado, " & lOrdenGr(C_FECHA)
   lOrdenGr(C_USUARIO) = "Usuario, " & lOrdenGr(C_FECHA)
   
   lOrdenSel = C_FECHA
   
   Call LoadAll
   
   Me.Caption = Me.Caption & " - " & FmtRut(gEmpresa.Rut)
   
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid, True)
   Grid.FixedCols = 0
 
   Grid.ColWidth(C_IDDTE) = 0
   Grid.ColWidth(C_RUT) = 1200
   Grid.ColWidth(C_RSOCIAL) = 3300
   Grid.ColWidth(C_TIPODOC) = 2000
   Grid.ColWidth(C_DIMINUTIVO) = 0
   Grid.ColWidth(C_CODDOCSII) = 0
   Grid.ColWidth(C_FOLIO) = 1100
   Grid.ColWidth(C_FECHA) = 1100
   Grid.ColWidth(C_LNGFECHA) = 0
   Grid.ColWidth(C_TOTAL) = 1200
   Grid.ColWidth(C_ESTADODTE) = 1700
   Grid.ColWidth(C_IDESTADODTE) = 0
   Grid.ColWidth(C_IDESTADOSII) = 0
   Grid.ColWidth(C_ERRORSII) = 0
   Grid.ColWidth(C_RESPUESTASII) = 0
   Grid.ColWidth(C_GLOSASII) = 0
   Grid.ColWidth(C_TRACKID) = 0
'   If gConectData.Proveedor = PROV_LP Then
      Grid.ColWidth(C_VERPDF) = 500
      Grid.ColWidth(C_VERPDFCED) = 0
'   Else
'      Grid.ColWidth(C_VERPDF) = 1000
'      Grid.ColWidth(C_VERPDFCED) = 0
'   End If
   Grid.ColWidth(C_USUARIO) = 1000
   Grid.ColWidth(C_URLDTE) = 0
   Grid.ColWidth(C_NULL) = 0
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_FOLIO) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHA) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT Receptor"
   Grid.TextMatrix(0, C_RSOCIAL) = "Razón Social"
   Grid.TextMatrix(0, C_TIPODOC) = "Tipo Documento"
   Grid.TextMatrix(0, C_FOLIO) = "Folio"
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_TOTAL) = "Monto"
   Grid.TextMatrix(0, C_ESTADODTE) = "Estado"
   Grid.TextMatrix(0, C_VERPDF) = "PDF"
'   Grid.TextMatrix(0, C_VERPDFCED) = "CED"
   Grid.TextMatrix(0, C_USUARIO) = "Usuario"
   
   Call FGrVRows(Grid, 1)
End Sub
Private Sub Form_Resize()

   Grid.Width = Me.Width - 200
   Grid.Height = Me.Height - Grid.Top - Tx_CurCel.Height - 600
   Tx_CurCel.Top = Grid.Top + Grid.Height + 60
   
   Call FGrVRows(Grid, 1)

End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
   Dim FDesde As Long, FHasta As Long
   Dim i As Integer
   Dim Diminutivo As String
   
   FDesde = GetTxDate(Tx_FechaDesde)
   FHasta = GetTxDate(Tx_FechaHasta)
      
   If FDesde > FHasta Then
      MsgBox1 "Rango de fecha inválido.", vbExclamation
      Exit Sub
   End If
   
   Wh = " WHERE DTE.IdEmpresa = " & gEmpresa.Id & " AND (DTE.TipoLib = " & lTipoLib & " OR DTE.TipoLib = " & LIB_OTROS & ")"  ' LIB_OTROS por la Guía de Despacho
   If Tx_RUT <> "" Then
      Wh = Wh & " AND Entidades.Rut = '" & vFmtCID(Tx_RUT) & "'"
   End If
   
   If Tx_RazonSocial <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Trim(Tx_RazonSocial), "Entidades.Nombre")
   End If
   
   If Tx_Folio <> "" Then
      Wh = Wh & " AND DTE.Folio = " & Trim(Tx_Folio)
   End If
    
   If CbItemData(Cb_TipoDoc) > 0 Then
      Wh = Wh & " AND DTE.CodDocSII= '" & CbItemData(Cb_TipoDoc) & "'"
   End If
    
   If CbItemData(Cb_Estado) > 0 Then
      Wh = Wh & " AND DTE.IdEstado = " & CbItemData(Cb_Estado)
      
   ElseIf O_SELCOPY Then
      Wh = Wh & " AND DTE.IdEstado IN (" & EDTE_ENVIADO & "," & EDTE_PROCESADO & "," & EDTE_EMITIDO & ")"
      
   End If
    
   If FDesde > 0 And FHasta > 0 Then
      Wh = Wh & " AND (DTE.Fecha BETWEEN " & FDesde & " AND " & FHasta & ")"
   ElseIf FDesde > 0 Then
      Wh = Wh & " AND DTE.Fecha >= " & FDesde
   ElseIf FHasta > 0 Then
      Wh = Wh & " AND DTE.Fecha <= " & FHasta
   End If
   
   If lNotNotasCredDeb Then
      Wh = Wh & " AND TipoDocs.Diminutivo NOT IN ( 'NCV', 'NDV', 'NCE', 'NDE')"
   End If
     
   Q1 = "SELECT IdDTE, Entidades.RUT, Entidades.Nombre, DTE.TipoLib, DTE.TipoDoc, DTE.CodDocSII, "
   Q1 = Q1 & " Folio, Fecha, Total, IdEstado, IdEstadoSII, ErrorSII, RespuestaSII, GlosaSII, TrackID, UrlDTE, Usuario "
   Q1 = Q1 & " FROM ((DTE INNER JOIN Entidades ON DTE.IdEntidad = Entidades.IdEntidad) "
   Q1 = Q1 & " INNER JOIN Usuarios ON DTE.IdUsuario = Usuarios.IdUsuario) "
   
   If lNotNotasCredDeb Then
      Q1 = Q1 & " INNER JOIN TipoDocs ON DTE.TipoLib = TipoDocs.TipoLib AND DTE.TipoDoc = TipoDocs.TipoDoc "
   End If
   
   Q1 = Q1 & Wh
   Q1 = Q1 & " AND DTE.TIPODOC <> 0 "
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Grid.Redraw = False
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDTE) = vFld(Rs("IdDTE"))
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RUT")))
      Grid.TextMatrix(i, C_RSOCIAL) = vFld(Rs("Nombre"))
      
      If vFld(Rs("TipoLib")) = LIB_OTROS And vFld(Rs("TipoDoc")) = TIPODOC_GUIADESPACHO Then
         Grid.TextMatrix(i, C_TIPODOC) = gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).Nombre
         Grid.TextMatrix(i, C_DIMINUTIVO) = gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).Diminutivo
      Else
         Grid.TextMatrix(i, C_TIPODOC) = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Nombre
         Grid.TextMatrix(i, C_DIMINUTIVO) = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Diminutivo
      End If
      
      Grid.TextMatrix(i, C_CODDOCSII) = vFld(Rs("CodDocSII"))
      Grid.TextMatrix(i, C_FOLIO) = vFld(Rs("Folio"))
      Grid.TextMatrix(i, C_FECHA) = Format(vFld(Rs("Fecha")), EDATEFMT)
      Grid.TextMatrix(i, C_LNGFECHA) = vFld(Rs("Fecha"))
      Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NUMFMT)
      Grid.TextMatrix(i, C_ESTADODTE) = gEstadoDTE(vFld(Rs("IdEstado")))
      If vFld(Rs("IdEstadoSII")) > 0 Then
         Grid.TextMatrix(i, C_ESTADODTE) = Grid.TextMatrix(i, C_ESTADODTE) & "/" & gDesEstadoDTESII(vFld(Rs("IdEstadoSII"))) 'gEstadoDTESII(vFld(Rs("IdEstadoSII")))
      End If
      Grid.TextMatrix(i, C_IDESTADODTE) = vFld(Rs("IdEstado"))
      Grid.TextMatrix(i, C_IDESTADOSII) = vFld(Rs("IdEstadoSII"))
      Grid.TextMatrix(i, C_ERRORSII) = vFld(Rs("ErrorSII"))
      Grid.TextMatrix(i, C_RESPUESTASII) = vFld(Rs("RespuestaSII"))
      Grid.TextMatrix(i, C_GLOSASII) = vFld(Rs("GlosaSII"))
      Grid.TextMatrix(i, C_TRACKID) = vFld(Rs("TrackID"))
      Grid.TextMatrix(i, C_URLDTE) = vFld(Rs("UrlDTE"))
      Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"))
      Grid.Row = i
      Grid.Col = C_VERPDF
      Grid.CellPictureAlignment = flexAlignCenterCenter
      Set Grid.CellPicture = FrmMain.Pc_Doc
      
      
'      If Grid.TextMatrix(i, C_DIMINUTIVO) = "FAV" Or Grid.TextMatrix(i, C_DIMINUTIVO) = "FVE" Then
'         Grid.Col = C_VERPDFCED
'         Grid.CellPictureAlignment = flexAlignCenterCenter
'         Set Grid.CellPicture = FrmMain.Pc_DocGreen
'      End If
      i = i + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
      
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Grid.Redraw = True
End Sub

Private Sub Grid_Click()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row >= Grid.FixedRows Then
      Exit Sub
   End If

   Call OrdenaPorCol(Col)
   
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   
   If Col = C_VERPDF Then
      If Grid.TextMatrix(Row, C_IDESTADODTE) <> EDTE_EMITIDO Then
         MsgBox1 "No es posible ver el PDF de este documento.", vbExclamation
         Exit Sub
      End If
      Me.MousePointer = vbHourglass
      
      If gConectData.Proveedor = PROV_ACEPTA Then
         If Grid.TextMatrix(Row, C_URLDTE) = "" Then
            MsgBox1 "No se encuentra disponible el DTE para ser impreso." & vbCrLf & vbCrLf & "Verifique el estado del DTE.", vbExclamation
         Else
            Call AcpShowDTE(Me, Grid.TextMatrix(Row, C_URLDTE))
         End If
      End If
      
      Me.MousePointer = vbDefault
      
   ElseIf Col = C_VERPDFCED Then
   
      If gConectData.Proveedor = PROV_ACEPTA Then
         If Grid.TextMatrix(Row, C_URLDTE) = "" Then
            MsgBox1 "No se encuentra disponible el DTE para ser impreso." & vbCrLf & vbCrLf & "Verifique el estado del DTE.", vbExclamation
         Else
            Call AcpShowDTE(Me, Grid.TextMatrix(Row, C_URLDTE))
         End If
      End If
      
   ElseIf Col = C_ESTADODTE Then
      Call Bt_VerDetEstado_Click
      
   
   ElseIf Bt_Select.Visible Then
'      If Grid.TextMatrix(Row, C_IDESTADODTE) <> EDTE_EMITIDO Then
'         MsgBox1 "Este documento no está en estado emitido.", vbExclamation
'         Exit Sub
''         If MsgBox1("Este documento no está en estado emitido." & vbCrLf & vbCrLf & "¿Desea seleccionarlo de todas maneras, BAJO SU RESPONSABILIDAD, estando seguro que aparece emitido en el SII?", vbExclamation + vbYesNoCancel) <> vbYes Then
''            Exit Sub
''         End If
''         MsgBox1 "ATENCIÓN: Verifique el folio del documento anulado y corrijalo si es necesario.", vbInformation
'      End If
      Call Bt_Select_Click
      
   End If
   
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Grid.Col = C_VERPDF Then
      Grid.ToolTipText = "Ver PDF Documento"
   ElseIf Grid.Col = C_VERPDFCED Then
      Grid.ToolTipText = "Ver PDF Copia Cedible"
   Else
      Grid.ToolTipText = ""
   End If
     
End Sub

Private Sub Grid_SelChange()
   Tx_CurCel = Grid.TextMatrix(Grid.Row, Grid.Col)
   
End Sub

Private Sub Tx_FechaDesde_GotFocus()
   Call DtGotFocus(Tx_FechaDesde)
End Sub

Private Sub Tx_FechaDesde_LostFocus()
   
   If Trim$(Tx_FechaDesde) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FechaDesde)
   
End Sub

Private Sub Tx_FechaDesde_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_SelFechaDesde_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaDesde)
   Set Frm = Nothing
   
   
End Sub
Private Sub Tx_FechaHasta_GotFocus()
   Call DtGotFocus(Tx_FechaHasta)
End Sub

Private Sub Tx_FechaHasta_LostFocus()
   
   If Trim$(Tx_FechaHasta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FechaHasta)
   
End Sub

Private Sub Tx_FechaHasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_SelFechaHasta_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaHasta)
   Set Frm = Nothing
   
   
End Sub
'
'Private Function GetPdfDTE(ByVal Folio As Long, ByVal TipoDTE As Integer, ByVal FechaDTE As Long, Optional ByVal CopiaCedible As Boolean = False) As Integer
'   Dim Rc As Integer
'   Dim Fn As String, FnCed As String
'
'   GetPdfDTE = False
'
'   Fn = GetFNamePDF(FechaDTE, TipoDTE, Folio, CopiaCedible)
'
'   If ExistFile(Fn) Then
'      Call AbrirPDF(Fn)
'      GetPdfDTE = True
'      Exit Function
'   End If
'
'   Rc = LpObtenerLink(LP_TC_NORMAL, Folio, TipoDTE, FechaDTE, Fn, False)
'
'   If Rc = 0 Then
'      Call LpObtenerLink(LP_TC_CEDIBLE, Folio, TipoDTE, FechaDTE, FnCed, False)
'
'      If CopiaCedible Then
'         Call AbrirPDF(FnCed)
'      Else
'         Call AbrirPDF(Fn)
'      End If
'      GetPdfDTE = True
'
'   ElseIf Rc = LP_ERR_PDFNOTAVAILABLE Then
'      MsgBox1 "Aún no está disponible el PDF del documento", vbInformation
'   End If
'
'End Function

Private Sub Tx_RUT_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Tx_RUT_LostFocus
      KeyAscii = 0
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   End If
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RUT) Then
      Cancel = True
      Exit Sub
   End If
   
End Sub

Private Sub Tx_RUT_LostFocus()
   Dim AuxRut As String
   
   AuxRut = FmtCID(vFmtCID(Tx_RUT))
   If AuxRut <> "0-0" Then
      Tx_RUT = AuxRut
   End If
   
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   If Col > C_ESTADODTE And Col <> C_USUARIO Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadAll
      
   Me.MousePointer = vbDefault
      
End Sub

