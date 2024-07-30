VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmEntidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entidades"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "FrmEntidades.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   6075
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   10716
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
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8760
      TabIndex        =   12
      Top             =   540
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   480
      TabIndex        =   21
      Top             =   420
      Width           =   7815
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   4380
         TabIndex        =   3
         ToolTipText     =   "Ingrese cualquier parte dell nombre o razón social de la Entidad"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   6360
         Picture         =   "FrmEntidades.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   1155
      End
      Begin VB.ComboBox Cb_OrdenarPor 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1875
      End
      Begin VB.ComboBox Cb_Clasif 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   3
         Left            =   3420
         TabIndex        =   26
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   900
         Width           =   390
      End
      Begin VB.Label Label1 
         Caption         =   "Ordenar por:"
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   24
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Clasificación:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Fr_SelEdit 
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   8760
      TabIndex        =   23
      Top             =   1980
      Width           =   1095
      Begin VB.CommandButton Bt_CopyExcel 
         Caption         =   "&Copiar Excel"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Copiar datos a Excel"
         Top             =   4500
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":0B11
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Seleccionar"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":1153
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Nueva Entidad"
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":16E5
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Eliminar Entidad"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":1D47
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Modificar Entidad"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":231A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir Entidad"
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Edit 
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   8760
      TabIndex        =   19
      Top             =   1920
      Width           =   1095
      Begin VB.CommandButton Bt_CopyExcel 
         Caption         =   "&Copiar Excel"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":2949
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Copiar datos a Excel"
         Top             =   3645
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":2EFE
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprimir Entidad"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":352D
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modificar Entidad"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":3B00
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Eliminar Entidad"
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":4162
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Nueva Entidad"
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Sel 
      BorderStyle     =   0  'None
      Height          =   5265
      Left            =   8715
      TabIndex        =   20
      Top             =   1920
      Width           =   1155
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":46F4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Seleccionar"
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Label la_nReg 
      AutoSize        =   -1  'True
      Caption         =   "..."
      Height          =   195
      Left            =   8760
      TabIndex        =   27
      Top             =   7800
      Width           =   135
   End
End
Attribute VB_Name = "FrmEntidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_RUT = 0
Const C_CODIGO = 1
Const C_NOMBRE = 2
Const C_ESTADO = 3
Const C_DIRECCION = 4
Const C_TELEFONO = 5
Const C_FAX = 6
Const C_EMAIL = 7
Const C_WEB = 8
Const C_ID = 9
Const C_IDESTADO = 10
Const C_NOTVALIDRUT = 11

Const NCOLS = C_NOTVALIDRUT

Dim lEntidad As Entidad_t
Dim lTipoEntidad As Integer
Dim lRc As Integer

Dim InLoad As Boolean

Dim lOper As Integer
Dim lNotValidRut As Boolean
Dim lAllRut As Boolean ' 15 feb 2020

Private Sub Bt_Buscar_Click()

   If Not InLoad Then
      Me.MousePointer = vbHourglass
      DoEvents
      Call LoadAll(Cb_Clasif)
      Me.MousePointer = vbDefault
   End If

End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click(Index As Integer)
   MousePointer = vbHourglass
   Call FGr2Clip(Grid, "Listado de " & Cb_Clasif)
   MousePointer = vbDefault
End Sub

Private Sub Bt_Del_Click(Index As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Exit Sub
   End If
   
   Q1 = "SELECT Count(*) as n FROM DTE WHERE idEntidad=" & vFmt(Grid.TextMatrix(Row, C_ID)) & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If vFld(Rs("n")) <> 0 Then
      MsgBox1 "No puede borrar la entidad " & Grid.TextMatrix(Row, C_NOMBRE) & ", existe un documento asociado.", vbExclamation
      Call CloseRs(Rs)
      Exit Sub
   End If
   Call CloseRs(Rs)
   
   If MsgBox1("¿Está seguro de eliminar la entidad " & Grid.TextMatrix(Row, C_NOMBRE) & "?", vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   
   Grid.RowHeight(Row) = 0
   Grid.Rows = Grid.Rows + 1
   Q1 = "DELETE FROM Entidades WHERE idEntidad=" & vFmt(Grid.TextMatrix(Row, C_ID)) & " AND IdEmpresa = " & gEmpresa.id
   Call ExecSQL(DbMain, Q1)
   
   Call AddLog("Se elimina Entidad '" & Grid.TextMatrix(Row, C_NOMBRE) & "', RUT: " & Grid.TextMatrix(Row, C_RUT))
   
End Sub

Private Sub Bt_Edit_Click(Index As Integer)
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Rc As Integer
      
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_RUT) = "" Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   Call FillStruct(Row, Cb_Clasif)
   Set Frm = New FrmEntidad
   Rc = Frm.FEdit(lEntidad)
   If Rc = vbOK Or Rc = vbRetry Then
      Call UpDateGrid(Row)
      
   End If
   Set Frm = Nothing
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_New_Click(Index As Integer)
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Rc As Integer
 
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   lEntidad.Clasif = CbItemData(Cb_Clasif)
   Rc = Frm.FNew(lEntidad)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      If lEntidad.Clasif = CbItemData(Cb_Clasif) Then
'         Row = FGrAddRow(Grid)
'         Call UpDateGrid(Row)
         Me.MousePointer = vbHourglass
         Call LoadAll(Cb_Clasif)
         Me.MousePointer = vbDefault
      End If
      
   ElseIf Rc = vbRetry Then ' ya existe
      If lEntidad.Clasif = CbItemData(Cb_Clasif) Then
         ' si ya existe lo buscamos para actualizarlo
         For Row = Grid.FixedRows To Grid.Rows - 1
            If Val(Grid.TextMatrix(Row, C_ID)) = lEntidad.id Then
               Call UpDateGrid(Row)
               Exit For
            End If
         Next Row
      End If
      
   End If
   
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_Print_Click(Index As Integer)
   Dim ColWi(C_NOTVALIDRUT) As Integer
   Dim Total(C_NOTVALIDRUT) As String
   Dim i As Integer
   Dim OldOrient As Integer
      
   If Grid.TextMatrix(1, C_RUT) = "" Then
      Exit Sub
   End If
   
   If SelPrinter() Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   Total(0) = ""
   
   OldOrient = Printer.Orientation
   Printer.Orientation = ORIENT_HOR
   Call PrtFlexGrid(Grid, "", "LISTADO DE ENTIDADES", "", Cb_Clasif, ColWi, Total, False, , , , , , , , , , , True)
   Printer.Orientation = OldOrient
   
   MousePointer = vbDefault
   
End Sub


Private Sub Bt_Sel_Click(Index As Integer)
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_RUT) = "" Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> EE_ACTIVO Then
      MsgBox1 "Esta entidad está en estado " & UCase(gEstadoEntidad(Val(Grid.TextMatrix(Row, C_IDESTADO)))) & vbCrLf & vbCrLf & "No puede ser receptora de un DTE.", vbExclamation
      lEntidad.Rut = ""
      lEntidad.Codigo = ""
      lEntidad.Nombre = ""

      lRc = vbCancel
      Exit Sub
   End If
   
'   Call FillStruct(Row, cb_ClasifSel)
   Call FillStruct(Row, Cb_Clasif)
   
   lRc = vbOK
   Unload Me
End Sub

Private Sub cb_Clasif_Click()

'   Me.MousePointer = vbHourglass
'   Call LoadAll(Cb_Clasif)
'   Me.MousePointer = vbDefault
   
End Sub

Private Sub Cb_OrdenarPor_Click()
   
'   If Not InLoad Then
'      Me.MousePointer = vbHourglass
'      Call LoadAll(Cb_Clasif)
'      Me.MousePointer = vbDefault
'   End If

End Sub

Private Sub Form_Load()
   
   lRc = vbCancel
   
   InLoad = True
   
   Call SetUpGrid
   
   Call CbAddItem(Cb_OrdenarPor, "Nombre", 1)
   Call CbAddItem(Cb_OrdenarPor, "RUT", 2)
   Cb_OrdenarPor.ListIndex = 0   'nombre
   
   Call CbAddItem(Cb_Clasif, " ", -1)
   
   Call FillCbClasifEnt(Cb_Clasif, lTipoEntidad)
   Cb_Clasif.ListIndex = 1 'clientes
   
   Fr_Edit.Visible = lOper = O_EDIT
   Fr_Sel.Visible = lOper = O_VIEW
   Fr_SelEdit.Visible = lOper = O_SELEDIT
   
'   Call FrmEnab(gEmpresa.FCierre = 0)
   Call LoadAll(Cb_Clasif)
   InLoad = False
      
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_RUT) = 1100
   Grid.ColWidth(C_CODIGO) = 1200
   Grid.ColWidth(C_NOMBRE) = 2800
   Grid.ColWidth(C_ESTADO) = 0
   
   Grid.ColWidth(C_DIRECCION) = 2400
   Grid.ColWidth(C_TELEFONO) = 1100
   Grid.ColWidth(C_FAX) = 1000
   Grid.ColWidth(C_EMAIL) = 2000
   Grid.ColWidth(C_WEB) = 2000
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_NOTVALIDRUT) = 0
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
      Grid.ColAlignment(i) = flexAlignLeftCenter
   Next i
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter

   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_CODIGO) = "Nombre Corto"
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_ESTADO) = ""   '"Estado"
   Grid.TextMatrix(0, C_DIRECCION) = "Dirección"
   Grid.TextMatrix(0, C_TELEFONO) = "Teléfonos"
   Grid.TextMatrix(0, C_FAX) = "Fax"
   Grid.TextMatrix(0, C_EMAIL) = "email"
   Grid.TextMatrix(0, C_WEB) = "WEB"
   
   
End Sub
Public Function FEdit(Optional ByVal AllRut As Boolean = 0) As Integer
   lOper = O_EDIT
   lAllRut = AllRut
   
   Me.Show vbModal
   
   FEdit = lRc
   
End Function
Friend Function FSelect(Entidad As Entidad_t, Optional ByVal TipoEntidad As Integer = ENT_CLIENTE) As Integer
   lOper = O_VIEW
   lTipoEntidad = TipoEntidad
   
   Me.Show vbModal
   
   FSelect = lRc
   Entidad = lEntidad
   
End Function
Friend Function FSelEdit(Entidad As Entidad_t, Optional ByVal TipoEntidad As Integer = ENT_CLIENTE, Optional ByVal NotValidRut As Boolean = 0) As Integer
   lOper = O_SELEDIT
   lTipoEntidad = TipoEntidad
   lNotValidRut = NotValidRut
   
   Me.Show vbModal
   
   FSelEdit = lRc
   Entidad = lEntidad
   
End Function

Private Sub UpDateGrid(Row As Integer)
   
   Grid.TextMatrix(Row, C_RUT) = lEntidad.Rut
   'Grid.TextMatrix(Row, C_NOMBRE) = lEntidad.Nombre
   Grid.TextMatrix(Row, C_CODIGO) = lEntidad.Codigo
   Grid.TextMatrix(Row, C_NOMBRE) = lEntidad.Nombre
   Grid.TextMatrix(Row, C_ESTADO) = gEstadoEntidad(lEntidad.Estado)
   Grid.TextMatrix(Row, C_IDESTADO) = lEntidad.Estado
   Grid.TextMatrix(Row, C_ID) = lEntidad.id
   
End Sub
Private Sub LoadAll(Cb As ComboBox)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Clasif As Integer
   
   Clasif = CbItemData(Cb)

   ' OJO *******
   Q1 = "UPDATE Entidades SET IdEmpresa = " & gEmpresa.id & " WHERE IdEmpresa=0" ' 15 feb 2020
   Call ExecSQL(DbMain, Q1)

   Q1 = "SELECT idEntidad, Rut, Codigo, Nombre, Estado, NotValidRut, Direccion,Ciudad,"
   Q1 = Q1 & "Telefonos,Fax,email,Web"
   Q1 = Q1 & " FROM Entidades"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   If Clasif >= 0 Then
      Q1 = Q1 & " AND Clasif" & Clasif & "=" & CON_CLASIF
   End If
   
   If lAllRut = False Then
      If lNotValidRut <> 0 Then
         Q1 = Q1 & " AND (NotValidRut <> 0 OR Rut = '" & RUT_DEFEXPORT & "')"
      Else
         Q1 = Q1 & " AND NotValidRut = 0 "
      End If
   End If
   
   If Tx_RUT <> "" Then
      Q1 = Q1 & " AND Rut = '" & vFmtCID(Tx_RUT, lNotValidRut = 0) & "'"
   End If
   
   If Tx_Nombre <> "" Then
      Q1 = Q1 & " AND " & GenLike(DbMain, Tx_Nombre, "Nombre")
   End If
   
   If LCase(Cb_OrdenarPor) = "rut" Then
      Q1 = Q1 & " ORDER BY right('0' & RUT,8)"
   Else
      Q1 = Q1 & " ORDER BY " & Cb_OrdenarPor
   
   End If
   
   la_nReg = ""
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   i = Grid.FixedRows
   Grid.TopRow = i   ' 21 sep 2020 - pam: para que muestre desde el inicio
   Grid.Rows = i
   Do While Rs.EOF = False
      Grid.Rows = i + 1
      
      Grid.TextMatrix(i, C_RUT) = FmtStRut(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)       'FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)
      Grid.TextMatrix(i, C_CODIGO) = vFld(Rs("Codigo"))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"))
      Grid.TextMatrix(i, C_ESTADO) = gEstadoEntidad(vFld(Rs("Estado")))
      Grid.TextMatrix(i, C_DIRECCION) = vFld(Rs("Direccion")) & " " & vFld(Rs("Ciudad"))
      Grid.TextMatrix(i, C_TELEFONO) = vFld(Rs("Telefonos"))
      Grid.TextMatrix(i, C_FAX) = vFld(Rs("Fax"))
      Grid.TextMatrix(i, C_EMAIL) = vFld(Rs("Email"))
      Grid.TextMatrix(i, C_WEB) = vFld(Rs("Web"))
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      Grid.TextMatrix(i, C_ID) = vFld(Rs("idEntidad"))
      Grid.TextMatrix(i, C_NOTVALIDRUT) = vFld(Rs("NotValidRut"))
      
      i = i + 1
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   Grid.Redraw = True
   
   la_nReg = i - Grid.FixedRows
   
End Sub
Private Sub FillStruct(Row As Integer, Cb As ComboBox)

   lEntidad.Rut = Grid.TextMatrix(Row, C_RUT)
   lEntidad.Codigo = Grid.TextMatrix(Row, C_CODIGO)
   lEntidad.Nombre = Grid.TextMatrix(Row, C_NOMBRE)
   lEntidad.Estado = Val(Grid.TextMatrix(Row, C_IDESTADO))
   lEntidad.id = Grid.TextMatrix(Row, C_ID)
   lEntidad.Clasif = CbItemData(Cb)
   lEntidad.NotValidRut = Val(Grid.TextMatrix(Row, C_NOTVALIDRUT))
   lEntidad.email = Grid.TextMatrix(Row, C_EMAIL)

   
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - w.YCaption - Grid.Top - 435 - w.yFrame
   Grid.Width = Me.Width - Grid.Left - Fr_Edit.Width - 435 * 2 - w.xFrame * 2
   Fr_Edit.Left = Grid.Left + Grid.Width + 435
   
   la_nReg.Top = Grid.Top + Grid.Height - la_nReg.Height
   la_nReg.Left = Fr_Edit.Left

   Call FGrVRows(Grid)
 
End Sub

Private Sub Grid_DblClick()

   If lOper = O_VIEW Then
      Call Bt_Sel_Click(0)
   ElseIf lOper = O_EDIT Then
      Call Bt_Edit_Click(0)
   ElseIf lOper = O_SELEDIT Then
      Call Bt_Sel_Click(1)
   End If

End Sub
Private Sub FrmEnab(ByVal bool As Boolean)
   Dim i As Integer

   If Not ChkPriv(PRV_ADM_DEF) Then
      bool = False
   End If
   
   For i = 0 To 1
      Bt_New(i).Enabled = bool
      Bt_Edit(i).Enabled = bool
      Bt_Del(i).Enabled = bool
      Bt_CopyExcel(i).Enabled = bool
   Next i

End Sub
Private Sub FillCbClasifEnt(Cb As ComboBox, Optional ByVal TipoEnt As Integer = ENT_CLIENTE)
   Dim i As Integer
   
   For i = ENT_CLIENTE To ENT_OTRO
      Cb.AddItem gClasifEnt(i)
      Cb.ItemData(Cb.NewIndex) = i
      
   Next i
   Cb.ListIndex = TipoEnt
   
End Sub

Private Sub Tx_Rut_LostFocus()
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If vFmtCID(Tx_RUT, lNotValidRut = 0) = "" Then
      Tx_RUT = ""
      Tx_RUT.SetFocus
      Exit Sub
   End If
   
'   If Not MsgValidRut(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
'
   Tx_RUT = FmtCID(vFmtCID(Tx_RUT, lNotValidRut = 0), lNotValidRut = 0)
   
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If Trim(Tx_RUT) = "0-0" Then
      MsgBox1 "RUT Inválido.", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RUT, lNotValidRut = 0) Then
      Tx_RUT.SetFocus
      Cancel = True
      Exit Sub
      
   End If
   
   
End Sub
Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)
   If lNotValidRut = 0 Then
      Call KeyCID(KeyAscii)
   
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   
   End If
End Sub

