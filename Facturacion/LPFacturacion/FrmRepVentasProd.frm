VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRepVentasProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ventas de Productos"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12465
   Icon            =   "FrmRepVentasProd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   12465
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5595
      Left            =   1020
      TabIndex        =   19
      Top             =   2220
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   9869
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   1020
      TabIndex        =   5
      Top             =   240
      Width           =   9795
      Begin VB.TextBox Tx_RazonSocial 
         Height          =   315
         Left            =   3720
         TabIndex        =   22
         ToolTipText     =   "Ingrese cualquier parte de la Razón Social"
         Top             =   1260
         Width           =   5745
      End
      Begin VB.TextBox Tx_RUT 
         Height          =   315
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   21
         Top             =   1260
         Width           =   1455
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   375
         Left            =   8340
         Picture         =   "FrmRepVentasProd.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox Tx_TipoCod 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         ToolTipText     =   "Ingrese cualquier parte del tipo de código"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Ingrese cualquier parte del código"
         Top             =   780
         Width           =   2055
      End
      Begin VB.TextBox Tx_FechaHasta 
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton Bt_SelFechaHasta 
         Height          =   315
         Left            =   4920
         Picture         =   "FrmRepVentasProd.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaDesde 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton Bt_SelFechaDesde 
         Height          =   315
         Left            =   2460
         Picture         =   "FrmRepVentasProd.frx":05D1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox Tx_Producto 
         Height          =   315
         Left            =   6660
         TabIndex        =   6
         ToolTipText     =   "Ingrese cualquier parte del nombre del Producto"
         Top             =   780
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "R. Social: "
         Height          =   195
         Index           =   6
         Left            =   2940
         TabIndex        =   24
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   23
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cod.:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   2
         Left            =   2940
         TabIndex        =   17
         ToolTipText     =   "Ingrese cualquier parte del código"
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "hasta:"
         Height          =   195
         Index           =   3
         Left            =   2940
         TabIndex        =   16
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Index           =   0
         Left            =   5880
         TabIndex        =   14
         ToolTipText     =   "Ingrese cualquier parte del código"
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.Frame Fr_Buttons 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5955
      Left            =   11040
      TabIndex        =   2
      Top             =   1860
      Width           =   1155
      Begin VB.CommandButton Bt_CopyExcel 
         Caption         =   "&Copiar Excel"
         Height          =   855
         Left            =   0
         Picture         =   "FrmRepVentasProd.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Copiar datos a Excel"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   0
         Picture         =   "FrmRepVentasProd.frx":0BFB
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir Entidad"
         Top             =   4140
         Width           =   1095
      End
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   11040
      TabIndex        =   1
      Top             =   300
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   240
      Picture         =   "FrmRepVentasProd.frx":122A
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   1020
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7800
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   11
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label Lb_Nota 
      Caption         =   "Nota: este reporte sólo considera DTE en estado EMITIDO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1020
      TabIndex        =   25
      Top             =   8220
      Width           =   9735
   End
End
Attribute VB_Name = "FrmRepVentasProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDPROD = 0
Const C_TIPOCOD = 1
Const C_CODPROD = 2
Const C_PRODUCTO = 3
Const C_UMEDIDA = 4
Const C_CANTIDAD = 5
Const C_TOTAL = 6

Private Const NCOLS = C_TOTAL

Dim lTipoLib As Integer

Dim lRc As Integer
Dim lIdProd As Long
Dim lOrientacion As Integer

Dim lOrdenGr(C_TOTAL) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

Private Sub Bt_Buscar_Click()
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub bt_Cerrar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim FDesde As Long
   
   lTipoLib = LIB_VENTAS
      
      
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_PRODUCTO) = "Producto, TipoCod, CodProd"
   
   lOrdenGr(C_TIPOCOD) = "TipoCod, CodProd"
   lOrdenGr(C_CODPROD) = "CodProd, TipoCod"
   lOrdenGr(C_UMEDIDA) = "UMedida, " & lOrdenGr(C_PRODUCTO)
   lOrdenGr(C_CANTIDAD) = "Sum(Cantidad), " & lOrdenGr(C_PRODUCTO)
   lOrdenGr(C_TOTAL) = "Sum(DetDTE.SubTotal), " & lOrdenGr(C_PRODUCTO)
   
   lOrdenSel = C_PRODUCTO
   
   Call SetUpGrid
   
   FDesde = DateAdd("m", -1, Now)
   Call SetTxDate(Tx_FechaDesde, DateSerial(Year(FDesde), Month(FDesde), 1))
   Call SetTxDate(Tx_FechaHasta, Now)

   
   Call LoadAll
   
   Call FGrVRows(Grid, 1)
   
   Call SetupPriv
      
End Sub
Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_IDPROD) = 0
   Grid.ColWidth(C_TIPOCOD) = 800
   Grid.ColWidth(C_CODPROD) = 1500
   Grid.ColWidth(C_PRODUCTO) = 2700 + 1200
   Grid.ColWidth(C_UMEDIDA) = 800
   Grid.ColWidth(C_CANTIDAD) = 1200
   Grid.ColWidth(C_TOTAL) = 1200
   
   Grid.ColAlignment(C_TIPOCOD) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODPROD) = flexAlignLeftCenter
   Grid.ColAlignment(C_PRODUCTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_UMEDIDA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CANTIDAD) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_TIPOCOD) = "Tipo Cód."
   Grid.TextMatrix(0, C_CODPROD) = "Código"
   Grid.TextMatrix(0, C_PRODUCTO) = "Producto"
   Grid.TextMatrix(0, C_UMEDIDA) = "U.Medida"
   Grid.TextMatrix(0, C_CANTIDAD) = "Cantidad"
   Grid.TextMatrix(0, C_TOTAL) = "Total Neto"
   
   Call FGrTotales(Grid, GridTot)
          
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim fld As Field
   Dim Wh As String
   Dim FDesde As Long, FHasta As Long
   Dim Total As Double
   
   FDesde = GetTxDate(Tx_FechaDesde)
   FHasta = GetTxDate(Tx_FechaHasta)
   
   If FDesde > FHasta Then
      MsgBox1 "Rango de fecha inválido.", vbExclamation
      Exit Sub
   End If
   
   Wh = " WHERE DTE.IdEmpresa = " & gEmpresa.Id & " AND DTE.TipoLib = " & lTipoLib & " AND IdEstado = " & EDTE_EMITIDO
   
   If FDesde > 0 And FHasta > 0 Then
      Wh = Wh & " AND (Fecha BETWEEN " & FDesde & " AND " & FHasta & ")"
   ElseIf FDesde > 0 Then
      Wh = Wh & " AND Fecha >= " & FDesde
   ElseIf FHasta > 0 Then
      Wh = Wh & " AND Fecha <= " & FHasta
   End If
   
   If Tx_TipoCod <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Tx_TipoCod, "TipoCod")
   End If
   
   If Tx_Codigo <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Tx_Codigo, "CodProd")
   End If
   
   If Tx_Producto <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Tx_Producto, "Producto")
   End If
   
   If Tx_RUT <> "" Then
      Wh = Wh & " AND Entidades.Rut = '" & vFmtCID(Tx_RUT) & "'"
   End If
   
   If Tx_RazonSocial <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Trim(Tx_RazonSocial), "Entidades.Nombre")
   End If
   
   
   Q1 = "SELECT TipoCod, CodProd, Producto, UMedida, "
   Q1 = Q1 & " Sum(iif(not TipoDocs.EsRebaja, DetDTE.Cantidad, 0)) As TotCantSuma, "
   Q1 = Q1 & " Sum(iif(TipoDocs.EsRebaja, DetDTE.Cantidad, 0)) As TotCantResta, "
   Q1 = Q1 & " Sum(iif(not TipoDocs.EsRebaja, DetDTE.SubTotal, 0)) As TotalSuma, "
   Q1 = Q1 & " Sum(iif(TipoDocs.EsRebaja, DetDTE.SubTotal, 0)) As TotalResta "
   Q1 = Q1 & " FROM ((DTE INNER JOIN DetDTE ON DTE.IdDTE = DetDTE.IdDTE) "
   Q1 = Q1 & " INNER JOIN Entidades ON DTE.IdEntidad = Entidades.IdEntidad)"
   Q1 = Q1 & " INNER JOIN TipoDocs ON ( DTE.TipoLib = TipoDocs.TipoLib AND DTE.TipoDoc = TipoDocs.TipoDoc )"
   Q1 = Q1 & Wh
   Q1 = Q1 & " GROUP BY TipoCod, CodProd, Producto, UMedida "
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   Total = 0
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
'      Grid.TextMatrix(i, C_IDPROD) = vFld(Rs("IdProducto"))
      Grid.TextMatrix(i, C_TIPOCOD) = vFld(Rs("TipoCod"))
      Grid.TextMatrix(i, C_CODPROD) = vFld(Rs("CodProd"))
      Grid.TextMatrix(i, C_PRODUCTO) = vFld(Rs("Producto"))
      Grid.TextMatrix(i, C_UMEDIDA) = vFld(Rs("UMedida"))
      Grid.TextMatrix(i, C_CANTIDAD) = Format(vFld(Rs("TotCantSuma")) - vFld(Rs("TotCantResta")), DBLFMT2)
      Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("TotalSuma")) - vFld(Rs("TotalResta")), NUMFMT)
      Total = Total + vFmt(Grid.TextMatrix(i, C_TOTAL))
      
      i = i + 1
      Rs.MoveNext
            
   Loop
   
   Call CloseRs(Rs)
   
   GridTot.TextMatrix(0, C_PRODUCTO) = "TOTAL"
   GridTot.TextMatrix(0, C_TOTAL) = Format(Total, NUMFMT)
   
   Call FGrVRows(Grid, 1)
   
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Grid.Row = Grid.FixedRows
   Grid.TopRow = Grid.Row
   
   Grid.Redraw = True
   

End Sub


Private Sub Form_Resize()

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - Lb_Nota.Height - 800
'   Grid.Width = Me.Width - Grid.Left - Bt_CopyExcel.Width - 600
   GridTot.Top = Grid.Top + Grid.Height + 30
   Lb_Nota.Top = GridTot.Top + GridTot.Height + 30
'   Fr_Buttons.Left = Grid.Left + Grid.Width + 200
   Call FGrVRows(Grid, 1)
   
End Sub

Private Sub SetupPriv()
      
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
   Dim Titulos(1) As String
   Dim Encabezados(1) As String
   Dim Totales(NCOLS) As String
   
   lOrientacion = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   Titulos(1) = "Periodo: " & Tx_FechaDesde & " - " & Tx_FechaHasta
   gPrtReportes.Titulos = Titulos
    
   i = 0
   If Tx_TipoCod <> "" Then
      Encabezados(i) = "Tipo Código:" & vbTab & Tx_TipoCod
      i = i + 1
   End If
   If Tx_Codigo <> "" Then
      Encabezados(i) = "Código:" & vbTab & Tx_Codigo
   End If
   
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Totales(i) = GridTot.TextMatrix(0, i)
   Next i
               
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_PRODUCTO
   gPrtReportes.Total = Totales
   gPrtReportes.NTotLines = 1
   

End Sub

Private Sub Bt_Sel_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_PRODUCTO) = "" Then
      Exit Sub
   End If
   
   lIdProd = Grid.TextMatrix(Grid.Row, C_IDPROD)
   
   lRc = vbOK
   Unload Me
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

Private Sub Tx_Codigo_LostFocus()
   Tx_Codigo = Trim(Tx_Codigo)
End Sub

Private Sub Tx_TipoCod_LostFocus()
   Tx_TipoCod = Trim(Tx_TipoCod)
End Sub

Private Sub Bt_CopyExcel_Click()
   
   Call FGr2Clip(Grid, Me.Caption & vbTab & "Periodo: " & Tx_FechaDesde & " - " & Tx_FechaHasta)
   
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   If Col > C_TOTAL Then
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

