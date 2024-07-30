VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12465
   Icon            =   "FrmProductos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12465
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Buttons 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5955
      Left            =   11040
      TabIndex        =   6
      Top             =   1320
      Width           =   1155
      Begin VB.CommandButton Bt_CopyExcel 
         Caption         =   "&Copiar Excel"
         Height          =   855
         Left            =   0
         Picture         =   "FrmProductos.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Copiar datos a Excel"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmProductos.frx":05C1
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Seleccionar"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmProductos.frx":0C03
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar Entidad"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   0
         Picture         =   "FrmProductos.frx":1265
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir Entidad"
         Top             =   4140
         Width           =   1095
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   11040
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   11040
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   240
      Picture         =   "FrmProductos.frx":1894
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   420
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   1020
      TabIndex        =   1
      Top             =   240
      Width           =   9795
      Begin VB.TextBox Tx_Producto 
         Height          =   315
         Left            =   5820
         TabIndex        =   15
         ToolTipText     =   "Ingrese cualquier parte dell nombre del Producto"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   2940
         TabIndex        =   13
         ToolTipText     =   "Ingrese cualquier parte del código"
         Top             =   360
         Width           =   1875
      End
      Begin VB.TextBox Tx_TipoCod 
         Height          =   315
         Left            =   1020
         TabIndex        =   11
         ToolTipText     =   "Ingrese cualquier parte del tipo de código"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   8400
         Picture         =   "FrmProductos.frx":1E75
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   16
         ToolTipText     =   "Ingrese cualquier parte del código"
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Ingrese cualquier parte del código"
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cod.:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   420
         Width           =   735
      End
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   5895
      Left            =   1020
      TabIndex        =   0
      Top             =   1320
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10398
      Cols            =   2
      Rows            =   2
      FixedCols       =   1
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
End
Attribute VB_Name = "FrmProductos"
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
Const C_PRECIO = 5
Const C_ESPROD = 6
Const C_OBS = 7
Const C_UPDATE = 8

Private Const NCOLS = C_UPDATE

Private lFldLen(NCOLS) As Byte

Dim lFldCols(NCOLS) As String
Dim lRc As Integer
Dim lOper As Integer
Dim lIdProd As Long
Dim lOrientacion As Integer

Dim lOrdenGr(C_OBS) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

Dim lModData As Boolean

Public Function FEdit()

   lOper = O_EDIT
   Me.Show vbModal
   
End Function
Public Function FSelect() As Long

   lOper = O_SELECT
   Me.Show vbModal
   FSelect = 0
   If lRc = vbOK Then
      FSelect = lIdProd
   End If
   
End Function

Private Sub Bt_Buscar_Click()

   Call ModSave

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub bt_OK_Click()

   If Not Valida() Then
      Exit Sub
   End If
      
   Call SaveAll
   
   Unload Me
End Sub

Private Sub bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdProd As Long
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdProd = Val(Grid.TextMatrix(Row, C_IDPROD))
   
   If IdProd > 0 Then
      Q1 = "SELECT IdDTE FROM DetDTE WHERE IdProducto = " & IdProd & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         MsgBox1 "No es posible eliminar este producto. Hay Documentos cuyo detalle incluye este producto.", vbExclamation
         Call CloseRs(Rs)
         Exit Sub
      End If
   
      Call CloseRs(Rs)
   End If
   
   If MsgBox1("¿Está seguro que desea elimnar este producto?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Call FGrModRow(Grid, Row, FGR_D, C_IDPROD, C_UPDATE)
   Grid.rows = Grid.rows + 1
   lModData = True



End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   lModData = False

   
   lFldCols(C_TIPOCOD) = "TipoCod"
   lFldCols(C_CODPROD) = "CodProd"
   lFldCols(C_PRODUCTO) = "Producto"
   lFldCols(C_UMEDIDA) = "UMedida"
   lFldCols(C_PRECIO) = "Precio"
   lFldCols(C_ESPROD) = "EsProducto"
   lFldCols(C_OBS) = "Obs"
      
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_PRODUCTO) = "Producto, TipoCod, CodProd"
   
   lOrdenGr(C_TIPOCOD) = "TipoCod, CodProd"
   lOrdenGr(C_CODPROD) = "CodProd, TipoCod"
   lOrdenGr(C_UMEDIDA) = "UMedida, " & lOrdenGr(C_PRODUCTO)
   lOrdenGr(C_PRECIO) = "Precio, " & lOrdenGr(C_PRODUCTO)
   lOrdenGr(C_ESPROD) = "EsProducto, " & lOrdenGr(C_PRODUCTO)
   lOrdenGr(C_OBS) = "Obs, " & lOrdenGr(C_PRODUCTO)
   
   lOrdenSel = C_PRODUCTO
   
   Call SetUpGrid
   
   If lOper = O_EDIT Then
      Bt_Sel.Visible = False
   Else
      Bt_Del.Visible = False
      Grid.Locked = True
      Bt_OK.Visible = False
      Bt_Cancel.Top = Bt_OK.Top
      Bt_Cancel.Caption = "Cerrar"
   End If
   
   Call LoadAll
   
   Call FGrVRows(Grid, 1)
   
   Call SetupPriv
      
End Sub
Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_TIPOCOD) = 800
   Grid.ColWidth(C_CODPROD) = 1500
   Grid.ColWidth(C_PRODUCTO) = 2500
   Grid.ColWidth(C_UMEDIDA) = 800
   Grid.ColWidth(C_PRECIO) = 1000
   Grid.ColWidth(C_ESPROD) = 800
   Grid.ColWidth(C_OBS) = 3600
   Grid.ColWidth(C_IDPROD) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_TIPOCOD) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODPROD) = flexAlignLeftCenter
   Grid.ColAlignment(C_PRECIO) = flexAlignRightCenter
   Grid.ColAlignment(C_ESPROD) = flexAlignCenterCenter
   Grid.ColAlignment(C_PRODUCTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_UMEDIDA) = flexAlignLeftCenter
   Grid.ColAlignment(C_OBS) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_TIPOCOD) = "Tipo Cód."
   Grid.TextMatrix(0, C_CODPROD) = "Código"
   Grid.TextMatrix(0, C_PRODUCTO) = "Producto o Servicio"
   Grid.TextMatrix(0, C_UMEDIDA) = "U.Medida"
   Grid.TextMatrix(0, C_PRECIO) = "Precio"
   Grid.TextMatrix(0, C_ESPROD) = "Es Prod."
   Grid.TextMatrix(0, C_OBS) = "Observaciones"
      
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim fld As Field
   Dim Wh As String
   
   Tx_TipoCod = Trim(Tx_TipoCod)
   Tx_Codigo = Trim(Tx_Codigo)
   Tx_Producto = Trim(Tx_Producto)
   
   If Tx_TipoCod <> "" Then
      Wh = GenLike(DbMain, Tx_TipoCod, "TipoCod")
   End If
   
   If Tx_Codigo <> "" Then
      If Wh <> "" Then
         Wh = Wh & " AND "
      End If
      Wh = Wh & GenLike(DbMain, Tx_Codigo, "CodProd")
   End If
   
   If Tx_Producto <> "" Then
      If Wh <> "" Then
         Wh = Wh & " AND "
      End If
      Wh = Wh & GenLike(DbMain, Tx_Producto, "Producto")
   End If
   
   If Wh <> "" Then
      Wh = " AND " & Wh
   End If
      
   
   Q1 = "SELECT IdProducto, TipoCod, CodProd, Producto, UMedida, Precio, EsProducto, Obs FROM Productos "
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.Id & Wh
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   Set Rs = OpenRs(DbMain, Q1)

   If lFldLen(C_TIPOCOD) = 0 Then
      lFldLen(C_TIPOCOD) = Rs("TipoCod").Size
      lFldLen(C_CODPROD) = Rs("CodProd").Size
      lFldLen(C_PRODUCTO) = Rs("Producto").Size
      lFldLen(C_UMEDIDA) = Rs("UMedida").Size
      lFldLen(C_PRECIO) = MAX_DIGITOSVALOR
      lFldLen(C_ESPROD) = 2
      lFldLen(C_OBS) = Rs("Obs").Size
   End If
   
   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDPROD) = vFld(Rs("IdProducto"))
      Grid.TextMatrix(i, C_TIPOCOD) = vFld(Rs("TipoCod"))
      Grid.TextMatrix(i, C_TIPOCOD) = vFld(Rs("TipoCod"))
      Grid.TextMatrix(i, C_CODPROD) = vFld(Rs("CodProd"))
      Grid.TextMatrix(i, C_PRODUCTO) = vFld(Rs("Producto"))
      Grid.TextMatrix(i, C_UMEDIDA) = vFld(Rs("UMedida"))
      Grid.TextMatrix(i, C_PRECIO) = Format(vFld(Rs("Precio")), NUMFMT)
      Grid.TextMatrix(i, C_ESPROD) = FmtSiNo(Abs(vFld(Rs("EsProducto"))))
      Grid.TextMatrix(i, C_OBS) = vFld(Rs("Obs"))
            
      i = i + 1
      Rs.MoveNext
            
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   'Marco la columna Ordenada
      
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Grid.Redraw = True

End Sub


Private Sub Form_Resize()

   Grid.Height = Me.Height - Grid.Top - 700
   Grid.Width = Me.Width - Grid.Left - Bt_Sel.Width - 600
   If Me.WindowState = vbMaximized Then
      Grid.ColWidth(C_OBS) = 6000
   Else
      Grid.ColWidth(C_OBS) = 3600
   End If
   Fr_Buttons.Left = Grid.Left + Grid.Width + 200
   Call FGrVRows(Grid, 1)
   
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
     Dim r As Integer
   
   Me.MousePointer = vbHourglass
   
   Value = Trim(Value)
   
   If lOper = O_SELECT Then
     Value = ""
   End If
   
   If Col = C_TIPOCOD Or Col = C_CODPROD Or Col = C_PRODUCTO Then
      
      Grid.TextMatrix(Row, Col) = Value
      
      For r = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(r, C_PRODUCTO) = "" Then
            Exit For
         End If
         
         If Grid.RowHeight(r) > 0 And (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_IDPROD) <> "") Then
            If Grid.TextMatrix(r, C_PRODUCTO) = "" Then
               Exit For
            End If
            
            If r <> Row Then
               If (StrComp(Grid.TextMatrix(Row, C_TIPOCOD), Grid.TextMatrix(r, C_TIPOCOD), vbTextCompare) = 0 And StrComp(Grid.TextMatrix(Row, C_CODPROD), Grid.TextMatrix(r, C_CODPROD), vbTextCompare) = 0) Or StrComp(Grid.TextMatrix(Row, C_PRODUCTO), Grid.TextMatrix(r, C_PRODUCTO), vbTextCompare) = 0 Then
                  MsgBox1 "Este Producto ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Me.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
         End If
      Next r
   
   ElseIf Col = C_PRECIO Then
      Value = Format(vFmt(Value), NUMFMT)
   End If
   
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDPROD, C_UPDATE)
   lModData = True
   
   
   
   Me.MousePointer = vbDefault
     
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Row > Grid.FixedRows And (Grid.TextMatrix(Row - 1, C_CODPROD) = "" Or Grid.TextMatrix(Row - 1, C_PRODUCTO) = "" Or Grid.TextMatrix(Row - 1, C_PRECIO) = "") Then
      MsgBox1 "Debe completar el registro anterior.", vbExclamation
      Exit Sub
   End If
      
   EdType = FEG_Edit
   Grid.TxBox.MaxLength = lFldLen(Col)
   
    
   If Row = Grid.rows - 1 Then
      Grid.rows = Grid.rows + 1
   End If
   
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

Private Sub SetupPriv()

   Call EnableForm0(Me, ChkPriv(PRV_ADM_EMPRESA))
      
End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim i As Integer
   
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_PRODUCTO) <> "" Then
         
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
            
            Q1 = "INSERT INTO Productos (idEmpresa, TipoCod, CodProd, Producto, UMedida, Precio, EsProducto, Obs) "
            Q1 = Q1 & " VALUES (" & gEmpresa.Id
            Q1 = Q1 & ",'" & Grid.TextMatrix(i, C_TIPOCOD) & "'"
            Q1 = Q1 & ",'" & Grid.TextMatrix(i, C_CODPROD) & "'"
            Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_PRODUCTO)) & "'"
            Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_UMEDIDA)) & "'"
            Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_PRECIO))
            Q1 = Q1 & "," & ValSiNo(Grid.TextMatrix(i, C_ESPROD))
            Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_OBS)) & "')"
            Call ExecSQL(DbMain, Q1)
         
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
            
            Q1 = "UPDATE Productos SET "
            Q1 = Q1 & "  TipoCod='" & Grid.TextMatrix(i, C_TIPOCOD) & "'"
            Q1 = Q1 & ", CodProd='" & Grid.TextMatrix(i, C_CODPROD) & "'"
            Q1 = Q1 & ", Producto='" & ParaSQL(Grid.TextMatrix(i, C_PRODUCTO)) & "'"
            Q1 = Q1 & ", UMedida='" & ParaSQL(Grid.TextMatrix(i, C_UMEDIDA)) & "'"
            Q1 = Q1 & ", Precio=" & vFmt(Grid.TextMatrix(i, C_PRECIO))
            Q1 = Q1 & ", EsProducto=" & ValSiNo(Grid.TextMatrix(i, C_ESPROD))
            Q1 = Q1 & ", Obs ='" & ParaSQL(Grid.TextMatrix(i, C_OBS)) & "'"
            Q1 = Q1 & " WHERE IdProducto=" & Grid.TextMatrix(i, C_IDPROD) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
         
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
         
            Q1 = "DELETE * FROM Productos "
            Q1 = Q1 & " WHERE IdProducto=" & Grid.TextMatrix(i, C_IDPROD) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
         
         End If
         
      End If
   Next i

   lModData = False

End Sub
Private Function Valida() As Boolean
   Dim r As Integer, i As Integer
   
   Valida = False
   
   Me.MousePointer = vbHourglass

   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_PRODUCTO) = "" Then
         Exit For
      End If
     For r = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(r, C_PRODUCTO) = "" Then
            Exit For
         End If
         If Grid.RowHeight(r) > 0 And (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_IDPROD) <> "") Then
            If r <> i And ((StrComp(Grid.TextMatrix(i, C_TIPOCOD), Grid.TextMatrix(r, C_TIPOCOD), vbTextCompare) = 0 And StrComp(Grid.TextMatrix(i, C_CODPROD), Grid.TextMatrix(r, C_CODPROD), vbTextCompare) = 0) Or StrComp(Grid.TextMatrix(i, C_PRODUCTO), Grid.TextMatrix(r, C_PRODUCTO), vbTextCompare) = 0) Then
               MsgBox1 "Este Producto ya existe.", vbExclamation
               Call FGrSelRow(Grid, i)
               Me.MousePointer = vbDefault
               Exit Function
            End If
         End If
      Next r
   Next i

   Valida = True
   Me.MousePointer = vbDefault

End Function


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
   Dim Encabezados(1) As String
   
   lOrientacion = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
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
   
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   ColWi(C_OBS) = 3700
               
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_PRODUCTO
   gPrtReportes.NTotLines = 0
   

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

Private Sub Grid_DblClick()
   Dim Col As Integer, Row As Integer
   
   Row = Grid.Row
   Col = Grid.Col

   If Col = C_ESPROD Then
      If UCase(Grid.TextMatrix(Row, Col)) = "SI" Then
         Grid.TextMatrix(Row, Col) = "No"
      Else
         Grid.TextMatrix(Row, Col) = "Si"
      End If
      
      Grid.TextMatrix(Row, C_UPDATE) = FGR_U
      
      Exit Sub
   End If

   If Bt_Sel.Visible Then
      Call Bt_Sel_Click
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Dim Col As Integer
   
   
    Col = Grid.Col
    Grid.TxBox.MaxLength = lFldLen(Col)
    
    If Col = C_TIPOCOD Or Col = C_CODPROD Then
       Call KeyCod(KeyAscii)
    End If
   
      
   
End Sub

Private Sub Tx_Codigo_LostFocus()
   Tx_Codigo = Trim(Tx_Codigo)
End Sub

Private Sub Tx_TipoCod_LostFocus()
   Tx_TipoCod = Trim(Tx_TipoCod)
End Sub

Private Sub Bt_CopyExcel_Click()
   
   Call FGr2Clip(Grid, "Listado de Productos")
   
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   If Col > C_OBS Then
      Exit Sub
   End If
   
   Call ModSave
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadAll
      
   Me.MousePointer = vbDefault
      
End Sub

Private Sub ModSave()

   If Not lModData Then
      Exit Sub
   End If

   If MsgBox1("Antes de listar los productos se grabarán los cambios realizados." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   If Not Valida() Then
      Exit Sub
   End If
      
   Me.MousePointer = vbHourglass
   
   Call SaveAll
   
   Me.MousePointer = vbDefault

End Sub
