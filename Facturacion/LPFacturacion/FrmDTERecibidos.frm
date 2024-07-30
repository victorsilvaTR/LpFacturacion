VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDTERecibidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos Electrónicos Recibidos"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13590
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_Titulo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      TabIndex        =   30
      Top             =   1920
      Width           =   13455
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   60
      TabIndex        =   27
      Top             =   600
      Width           =   1995
      Begin VB.OptionButton Op_GuiaDesp 
         Caption         =   "Guías Despacho"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1515
      End
      Begin VB.OptionButton Op_DTECompra 
         Caption         =   "Docs. Compra"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.TextBox Tx_CurCel 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   9240
      Width           =   7995
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2160
      TabIndex        =   13
      Top             =   600
      Width           =   11295
      Begin VB.CommandButton Bt_Buscar 
         Height          =   735
         Left            =   9900
         Picture         =   "FrmDTERecibidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton Bt_SelFechaDesde 
         Height          =   315
         Left            =   4620
         Picture         =   "FrmDTERecibidos.frx":0550
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaDesde 
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Bt_SelFechaHasta 
         Height          =   315
         Left            =   6780
         Picture         =   "FrmDTERecibidos.frx":05C5
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaHasta 
         Height          =   315
         Left            =   5520
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   8100
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1635
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   8100
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox Tx_Folio 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox Tx_RUT 
         Height          =   315
         Left            =   720
         MaxLength       =   12
         TabIndex        =   1
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox Tx_RazonSocial 
         Height          =   315
         Left            =   3360
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
         Left            =   2280
         TabIndex        =   25
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "hasta:"
         Height          =   195
         Index           =   2
         Left            =   5040
         TabIndex        =   24
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Index           =   6
         Left            =   7260
         TabIndex        =   18
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   5
         Left            =   7260
         TabIndex        =   17
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social: "
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   12
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
         Left            =   600
         Picture         =   "FrmDTERecibidos.frx":063A
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   120
         Picture         =   "FrmDTERecibidos.frx":0A30
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Revisar y Actualizar Estado DTE Seleccionado"
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
         Left            =   2280
         Picture         =   "FrmDTERecibidos.frx":0E95
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   2700
         Picture         =   "FrmDTERecibidos.frx":134F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   11940
         TabIndex        =   11
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
         Left            =   1680
         Picture         =   "FrmDTERecibidos.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1260
         Picture         =   "FrmDTERecibidos.frx":1BBD
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Calculadora"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6975
      Left            =   60
      TabIndex        =   0
      Top             =   2220
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "FrmDTERecibidos"
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
Const C_FEMISION = 7
Const C_LNGFEMISION = 8
Const C_TOTAL = 9   'Con IVA
Const C_FPUBLICACION = 10
Const C_LNGFPUBLICACION = 11
Const C_VERPDF = 12
Const C_URLDTE = 13
Const C_NULL = 14       'para evitar que al ampliar última columna aparezca la URL (truco)

Const NCOLS = C_NULL

Dim lOrdenGr(C_URLDTE) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual


Dim lTipoLib As Integer

Dim lOper As Integer
Dim lIdDTE As Long
Dim lRc As Integer
Dim lOrientacion As Integer


Public Function FView()
   lOper = O_VIEW
   Me.Show vbModal
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

Private Sub bt_Cerrar_Click()
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
   
   If Op_DTECompra = True Then
      Call FGr2Clip(Grid, "Documentos de Compra Electrónicos Recibidos" & Filtros)
   Else
      Call FGr2Clip(Grid, "Guías de Despacho Electrónicas Recibidas" & Filtros)
   End If
   
   Grid.ColWidth(C_VERPDF) = wVerPdf
   Grid.ColWidth(C_URLDTE) = 0
   Grid.TextMatrix(0, C_URLDTE) = ""
   
   Grid.Redraw = True
   
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
   
   If Op_DTECompra = True Then
      Titulos(0) = "Documentos de Compra Electrónicos Recibidos"
   Else
      Titulos(0) = "Guías de Despacho Electrónicas Recibidas"
   End If

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
   ColWi(C_VERPDF) = 0
   ColWi(C_RSOCIAL) = 3500
   ColWi(C_TIPODOC) = 1700

   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDDTE
   gPrtReportes.NTotLines = 0
   

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim FDesde As Long
   
   lTipoLib = LIB_COMPRAS
   Tx_Titulo = "Documentos de Compra"


   Call SetUpGrid
         
   Call CbAddItem(Cb_TipoDoc, "", 0, True)
   
   For i = 0 To UBound(gTipoDoc)
      If gTipoDoc(i).TipoLib = LIB_COMPRAS And Val(gTipoDoc(i).CodDocDTESII) > 0 Then
         Call CbAddItem(Cb_TipoDoc, gTipoDoc(i).Nombre, Val(gTipoDoc(i).CodDocDTESII))
      End If
   Next i
        
   FDesde = DateAdd("m", -1, Now)
   Call SetTxDate(Tx_FechaDesde, DateSerial(Year(FDesde), Month(FDesde), 1))
   Call SetTxDate(Tx_FechaHasta, Now)
     
   Call CbAddItem(Cb_Estado, "Aceptado por SII", 0, True)
   
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_FEMISION) = "FEmision Desc, IdDTE Desc"
   lOrdenGr(C_FPUBLICACION) = "FPublicacion Desc, IdDTE Desc"
   
   lOrdenGr(C_RUT) = "Entidades.RUT, " & lOrdenGr(C_FEMISION)
   lOrdenGr(C_RSOCIAL) = "Entidades.Nombre, " & lOrdenGr(C_FEMISION)
   lOrdenGr(C_TIPODOC) = "TipoDoc, " & lOrdenGr(C_FEMISION)
   lOrdenGr(C_FOLIO) = "Folio, " & lOrdenGr(C_FEMISION)
   lOrdenGr(C_TOTAL) = "Total, " & lOrdenGr(C_FEMISION)
   
   lOrdenSel = C_FEMISION
   
   Call LoadAll
   
   Me.Caption = Me.Caption & " - " & FmtRut(gEmpresa.Rut)

End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid, True)
   Grid.FixedCols = 0
   
   Grid.ColWidth(C_IDDTE) = 0
   Grid.ColWidth(C_RUT) = 1200
   Grid.ColWidth(C_RSOCIAL) = 4500
   Grid.ColWidth(C_TIPODOC) = 2000
   Grid.ColWidth(C_DIMINUTIVO) = 0
   Grid.ColWidth(C_CODDOCSII) = 0
   Grid.ColWidth(C_FOLIO) = 1300
   Grid.ColWidth(C_FEMISION) = 1200
   Grid.ColWidth(C_LNGFEMISION) = 0
   Grid.ColWidth(C_TOTAL) = 1200
   Grid.ColWidth(C_FPUBLICACION) = 1200
   Grid.ColWidth(C_LNGFPUBLICACION) = 0
   Grid.ColWidth(C_VERPDF) = 500
   Grid.ColWidth(C_URLDTE) = 0
   Grid.ColWidth(C_NULL) = 0
   
   Grid.ColAlignment(C_RSOCIAL) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_FOLIO) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_FEMISION) = flexAlignRightCenter
   Grid.ColAlignment(C_FPUBLICACION) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT Emisor"
   Grid.TextMatrix(0, C_RSOCIAL) = "Razón Social"
   Grid.TextMatrix(0, C_TIPODOC) = "Tipo Documento"
   Grid.TextMatrix(0, C_FOLIO) = "Folio"
   Grid.TextMatrix(0, C_FEMISION) = "Fecha Emisión"
   Grid.TextMatrix(0, C_TOTAL) = "Monto"
   Grid.TextMatrix(0, C_FPUBLICACION) = "Fecha Public."
   Grid.TextMatrix(0, C_VERPDF) = "PDF"
   
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
   
   Wh = " WHERE DTERecibidos.IdEmpresa = " & gEmpresa.id
   If Op_DTECompra Then
      Wh = Wh & " AND DTERecibidos.TipoLib = " & lTipoLib
   Else
      Wh = Wh & " AND (DTERecibidos.TipoLib = " & LIB_OTROS & " AND TipoDoc = " & TIPODOC_GUIADESPACHO & ")"   ' LIB_OTROS por la Guía de Despacho
   End If
   
   If Tx_RUT <> "" Then
      Wh = Wh & " AND Entidades.Rut = '" & vFmtCID(Tx_RUT) & "'"
   End If
   
   If Tx_RazonSocial <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Trim(Tx_RazonSocial), "Entidades.Nombre")
   End If
   
   If Tx_Folio <> "" Then
      Wh = Wh & " AND DTERecibidos.Folio = " & Trim(Tx_Folio)
   End If
    
   If CbItemData(Cb_TipoDoc) > 0 Then
      Wh = Wh & " AND DTERecibidos.CodDocSII= '" & CbItemData(Cb_TipoDoc) & "'"
   End If
    
'   If CbItemData(Cb_Estado) > 0 Then
'      Wh = Wh & " AND DTERecibidos.IdEstado = " & CbItemData(Cb_Estado)
'   End If
    
   If FDesde > 0 And FHasta > 0 Then
      Wh = Wh & " AND (DTERecibidos.FEmision BETWEEN " & FDesde & " AND " & FHasta & ")"
   ElseIf FDesde > 0 Then
      Wh = Wh & " AND DTERecibidos.FEmision >= " & FDesde
   ElseIf FHasta > 0 Then
      Wh = Wh & " AND DTERecibidos.FEmision <= " & FHasta
   End If
   
   Q1 = "SELECT IdDTE, Entidades.RUT, Entidades.Nombre, DTERecibidos.TipoLib, DTERecibidos.TipoDoc, DTERecibidos.CodDocSII, "
   Q1 = Q1 & " Folio, FEmision, FPublicacion, Total, UrlDTE "
   Q1 = Q1 & " FROM DTERecibidos INNER JOIN Entidades ON DTERecibidos.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & Wh
   
   Q1 = Q1 & " ORDER BY " & IIf(lOrdenGr(lOrdenSel) = "", lOrdenGr(C_FOLIO), lOrdenGr(lOrdenSel)) ' 29 ene 2021 - pam: por si no está vacío
   
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
      Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("FEmision")), EDATEFMT)
      Grid.TextMatrix(i, C_LNGFEMISION) = vFld(Rs("FEmision"))
      Grid.TextMatrix(i, C_FPUBLICACION) = Format(vFld(Rs("FPublicacion")), EDATEFMT)
      Grid.TextMatrix(i, C_LNGFPUBLICACION) = vFld(Rs("FPublicacion"))
      Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NUMFMT)
      Grid.TextMatrix(i, C_URLDTE) = vFld(Rs("UrlDTE"))
      Grid.Row = i
      Grid.Col = C_VERPDF
      Grid.CellPictureAlignment = flexAlignCenterCenter
      Set Grid.CellPicture = FrmMain.Pc_Doc
      
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
      Me.MousePointer = vbHourglass
      
      If Grid.TextMatrix(Row, C_URLDTE) = "" Then
         MsgBox1 "No se encuentra disponible el DTE para ser impreso.", vbExclamation
      Else
         Call AcpShowDTE(Me, Grid.TextMatrix(Row, C_URLDTE))
      End If
          
      Me.MousePointer = vbDefault
      
   End If
   
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Grid.Col = C_VERPDF Then
      Grid.ToolTipText = "Ver PDF Documento"
   Else
      Grid.ToolTipText = ""
   End If
     
End Sub

Private Sub Grid_SelChange()
   Tx_CurCel = Grid.TextMatrix(Grid.Row, Grid.Col)
   
End Sub

Private Sub Op_DTECompra_Click()
   Tx_Titulo = "Documentos de Compra"
   Call LoadAll
End Sub

Private Sub Op_GuiaDesp_Click()
   Tx_Titulo = "Guías de Despacho"
   Call LoadAll
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

Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Tx_Rut_LostFocus
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

Private Sub Tx_Rut_LostFocus()
   Dim AuxRut As String
   
   AuxRut = FmtCID(vFmtCID(Tx_RUT))
   If AuxRut <> "0-0" Then
      Tx_RUT = AuxRut
   End If
   
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   If Col >= C_VERPDF Then
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

