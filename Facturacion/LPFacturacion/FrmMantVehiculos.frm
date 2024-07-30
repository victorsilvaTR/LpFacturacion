VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMantVehiculos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehículos"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   Icon            =   "FrmMantVehiculos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   855
      Left            =   11100
      Picture         =   "FrmMantVehiculos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Seleccionar"
      Top             =   1980
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Print 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   11100
      Picture         =   "FrmMantVehiculos.frx":064E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir listado"
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   11100
      Picture         =   "FrmMantVehiculos.frx":0C7D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar Cláusula"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "&Copiar Excel"
      Height          =   855
      Left            =   11100
      Picture         =   "FrmMantVehiculos.frx":12DF
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar datos a Excel"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   10980
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton bt_Aceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   10980
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   4395
      Left            =   1500
      TabIndex        =   0
      Top             =   540
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   7752
      Cols            =   4
      Rows            =   20
      FixedCols       =   0
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
   Begin VB.Image Image1 
      Height          =   660
      Left            =   240
      Picture         =   "FrmMantVehiculos.frx":1894
      Top             =   540
      Width           =   855
   End
End
Attribute VB_Name = "FrmMantVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_PATENTE = 0
Const C_DESCRIP = 1
Const C_TIPO = 2
Const C_IDTIPO = 3
Const C_ID = 4
Const C_UPDATE = 5
Private Const NCOLS = C_UPDATE

Private FldLen(NCOLS) As Byte
Dim lOrientacion As Integer
Dim lOper As Integer
Dim lIdVehiculo As Long, lPatente As String
Dim lRc As Integer

Public Function FEdit()
   lOper = O_EDIT
   Me.Show vbModal
   
End Function

Public Function FView(IdVehiculo As Long, Patente As String) As Integer
   lOper = O_VIEW
   Me.Show vbModal
   
   If lRc = vbOK Then
      IdVehiculo = lIdVehiculo
      Patente = lPatente
   Else
      IdVehiculo = 0
      Patente = ""
   End If
   
   FView = lRc
   
End Function

Private Sub Bt_Aceptar_Click()
   Dim Q1 As String
   Dim i As Integer
   
   For i = 1 To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_PATENTE) <> "" Then
         
         If Len(Trim(Grid.TextMatrix(i, C_PATENTE))) < 6 Then
            MsgBox1 "Patente inválida '" & Grid.TextMatrix(i, C_PATENTE) & "'.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
         
         If Trim(Grid.TextMatrix(i, C_DESCRIP)) = "" Then
            MsgBox1 "Falta ingresar la descripción del vehículo.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
                           
         If vFmt(Grid.TextMatrix(i, C_IDTIPO)) = 0 Then
            MsgBox1 "Falta seleccionar el tipo del vehículo.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
                           
      End If
   Next i

   For i = 1 To Grid.rows - 1
      If Grid.TextMatrix(i, C_PATENTE) <> "" Then
      
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
            Q1 = "INSERT INTO Vehiculos (IdEmpresa, Patente, IdTipoVehiculo, Descrip)"
            Q1 = Q1 & " VALUES (" & gEmpresa.Id
            Q1 = Q1 & ", '" & ParaSQL(Grid.TextMatrix(i, C_PATENTE)) & "'"
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(i, C_IDTIPO))
            Q1 = Q1 & ", '" & ParaSQL(Grid.TextMatrix(i, C_DESCRIP)) & "' )"
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
            Q1 = "UPDATE Vehiculos SET "
            Q1 = Q1 & "  Patente ='" & ParaSQL(Grid.TextMatrix(i, C_PATENTE)) & "'"
            Q1 = Q1 & ", IdTipoVehiculo = " & Val(Grid.TextMatrix(i, C_IDTIPO))
            Q1 = Q1 & ", Descrip ='" & ParaSQL(Grid.TextMatrix(i, C_DESCRIP)) & "'"
            Q1 = Q1 & "  WHERE Id =" & Val(Grid.TextMatrix(i, C_ID)) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
         
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
         
            Q1 = "DELETE * FROM Vehiculos"
            Q1 = Q1 & " WHERE Id =" & Val(Grid.TextMatrix(i, C_ID)) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
         End If
      End If
   Next i
   
   Unload Me
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim i As Integer
   
   Call SetUpGrid
   
   Q1 = "SELECT Patente, Descrip, TipoVehiculo.Nombre, IdTipoVehiculo, Vehiculos.Id "
   Q1 = Q1 & " FROM Vehiculos INNER JOIN TipoVehiculo ON Vehiculos.IdTipoVehiculo = TipoVehiculo.Id "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.Id
   Q1 = Q1 & " ORDER BY TipoVehiculo.Nombre, Descrip"
   Call FillGrLista(Q1, Grid, C_ID, FldLen)
      
   Grid.rows = Grid.rows + 1
   
   Q1 = "SELECT TipoVehiculo.Nombre, Id FROM TipoVehiculo ORDER BY TipoVehiculo.Nombre "
   Call CbAddItem(Grid.CbList(C_TIPO), "", 0)
   Call FillCombo(Grid.CbList(C_TIPO), DbMain, Q1, 0)
   
   Call SetupPriv
   
   If lOper = O_VIEW Then
      Grid.Locked = True
      Bt_Del.Visible = False
      bt_Aceptar.Visible = False
   Else
      Bt_Sel.Visible = False
   End If
   
End Sub
Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_PATENTE) = 1200
   Grid.ColWidth(C_DESCRIP) = 4000
   Grid.ColWidth(C_TIPO) = 3500
   Grid.ColWidth(C_IDTIPO) = 0
   Grid.ColWidth(C_ID) = 600
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_DESCRIP) = flexAlignLeftCenter
   Grid.ColAlignment(C_PATENTE) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_PATENTE) = "Patente"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripcíón"
   Grid.TextMatrix(0, C_TIPO) = "Tipo de Vehiculo"
    
End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim r As Integer

   Value = Trim(Value)
   
   Select Case Col
      Case C_DESCRIP
      
         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_DESCRIP), vbTextCompare) = 0 Then
                  MsgBox1 "Este Vehículo ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Exit Sub
               End If
            End If
         Next r
   
      Case C_PATENTE
         If Value = "" Or Len(Value) < 6 Then
            MsgBox1 "Patente inválida.", vbExclamation
         End If

         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_PATENTE), vbTextCompare) = 0 And Val(Grid.TextMatrix(r, C_PATENTE)) <> 8 Then    'el código 8 se puede repetir
                  MsgBox1 "Esta patente ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Exit Sub
               End If
            End If
         Next r
         
      Case C_TIPO
         Grid.TextMatrix(Row, C_IDTIPO) = CbItemData(Grid.CbList(C_TIPO))
   
   End Select
    
   Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
   Call FGrEdRows(Grid, Row)
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Row > Grid.FixedRows Then
      If Grid.TextMatrix(Row - 1, C_DESCRIP) = "" Or Grid.TextMatrix(Row - 1, C_PATENTE) = "" Or vFmt(Grid.TextMatrix(Row - 1, C_IDTIPO)) = 0 Then
         MsgBox1 "Debe completar la línea anterior.", vbExclamation
         Exit Sub
      End If
   End If
   
   If (Col > C_PATENTE And Grid.TextMatrix(Row, C_PATENTE) = "") Or (Col > C_DESCRIP And Grid.TextMatrix(Row, C_DESCRIP) = "") Then
      MsgBox1 "Debe completar la línea campo a campo.", vbExclamation
      Exit Sub
   End If

   
   If Col = C_TIPO Then
      EdType = FEG_List
      
   Else
      Grid.TxBox.MaxLength = FldLen(Col)
      EdType = FEG_Edit
   End If
   
   If Grid.rows = Row + 1 Then
      Grid.rows = Grid.rows + 1
      Grid.FlxGrid.TopRow = Row
   End If
   
End Sub

Private Sub Bt_Sel_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Val(Grid.TextMatrix(Row, C_ID)) > 0 Then
      lIdVehiculo = Val(Grid.TextMatrix(Row, C_ID))
      lPatente = Grid.TextMatrix(Row, C_PATENTE)
      lRc = vbOK
      Unload Me
   End If
      
End Sub

Private Sub Grid_DblClick()

   If lOper = O_VIEW Then
      Call Bt_Sel_Click
   End If

End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)

   If Grid.Col = C_PATENTE Then
      Call KeyUCod(KeyAscii)
   End If
      
End Sub

Private Sub SetupPriv()

   Call EnableForm0(Me, ChkPriv(PRVF_MANT_DATOS))
         
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Id As Long
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Id = Val(Grid.TextMatrix(Row, C_ID))
   
   If Grid.TextMatrix(Row, C_DESCRIP) = "" And Grid.TextMatrix(Row, C_PATENTE) = "" Then
      Exit Sub
      
   End If
   
   If MsgBox1("¿Está seguro que desea elimnar este Vehículo?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Call FGrModRow(Grid, Row, FGR_D, C_ID, C_UPDATE)
   Grid.rows = Grid.rows + 1

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
   
   lOrientacion = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   gPrtReportes.Titulos = Titulos
    
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_DESCRIP
   gPrtReportes.NTotLines = 0
   

End Sub
Private Sub Bt_CopyExcel_Click()
   
   Call FGr2Clip(Grid, "Listado de Vehículos")
   
End Sub

