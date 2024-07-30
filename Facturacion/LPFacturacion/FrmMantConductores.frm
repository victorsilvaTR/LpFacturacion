VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMantConductores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antecedentes del Conductor"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "FrmMantConductores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   855
      Left            =   7920
      Picture         =   "FrmMantConductores.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Seleccionar"
      Top             =   1980
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Print 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   7920
      Picture         =   "FrmMantConductores.frx":064E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir lista de Países"
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   7920
      Picture         =   "FrmMantConductores.frx":0C7D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar País"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "&Copiar Excel"
      Height          =   855
      Left            =   7920
      Picture         =   "FrmMantConductores.frx":12DF
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar datos a Excel"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton bt_Aceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7800
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   4395
      Left            =   1140
      TabIndex        =   0
      Top             =   540
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7752
      Cols            =   3
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
      Height          =   630
      Left            =   240
      Picture         =   "FrmMantConductores.frx":1894
      Top             =   540
      Width           =   600
   End
End
Attribute VB_Name = "FrmMantConductores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_RUT = 0
Const C_NOMBRE = 1
Const C_ID = 2
Const C_UPDATE = 3
Private Const NCOLS = C_UPDATE

Private FldLen(NCOLS) As Byte
Dim lOrientacion As Integer
Dim lOper As Integer
Dim lIdConductor As Long, lNombre As String, lRut As String
Dim lRc As Integer

Public Function FEdit()
   lOper = O_EDIT
   Me.Show vbModal
   
End Function

Public Function FView(IdConductor As Long, Nombre As String, Rut As String) As Integer
   lOper = O_VIEW
   Me.Show vbModal
   
   If lRc = vbOK Then
      IdConductor = lIdConductor
      Nombre = lNombre
      Rut = lRut
   Else
      IdConductor = 0
      Nombre = ""
      Rut = ""
   End If
   
   FView = lRc
   
End Function

Private Sub Bt_Aceptar_Click()
   Dim Q1 As String
   Dim i As Integer
   
   For i = 1 To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_RUT) <> "" Then
         
         If Grid.TextMatrix(i, C_NOMBRE) = "" Then
            MsgBox1 "Falta ingresar nombre del chofer.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
                  
      End If
   Next i

   For i = 1 To Grid.rows - 1
      If Grid.TextMatrix(i, C_RUT) <> "" Then
      
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
            Q1 = "INSERT INTO Conductores (IdEmpresa, RUTChofer, NombreChofer)"
            Q1 = Q1 & " VALUES (" & gEmpresa.Id
            Q1 = Q1 & ", '" & vFmtRut(Grid.TextMatrix(i, C_RUT)) & "'"
            Q1 = Q1 & ", '" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "')"
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
            Q1 = "UPDATE Conductores SET NombreChofer ='" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ", RUTChofer ='" & vFmtRut(Grid.TextMatrix(i, C_RUT)) & "'"
            Q1 = Q1 & " WHERE Id =" & Val(Grid.TextMatrix(i, C_ID)) & " AND IdEmpresa = " & gEmpresa.Id
            Call ExecSQL(DbMain, Q1)
         
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
         
            Q1 = "DELETE * FROM Conductores "
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

Private Sub Bt_Sel_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Val(Grid.TextMatrix(Row, C_ID)) > 0 Then
      lIdConductor = Val(Grid.TextMatrix(Row, C_ID))
      lNombre = Grid.TextMatrix(Row, C_NOMBRE)
      lRut = Grid.TextMatrix(Row, C_RUT)
      lRc = vbOK
      Unload Me
   End If
      
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim i As Integer
   
   Call SetUpGrid
   
   Q1 = "SELECT RutChofer, NombreChofer, Id FROM Conductores WHERE IdEmpresa = " & gEmpresa.Id & " ORDER BY NombreChofer"
   Call FillGrLista(Q1, Grid, C_ID, FldLen)
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_RUT) <> "" Then
         Grid.TextMatrix(i, C_RUT) = FmtRut(Grid.TextMatrix(i, C_RUT))
      Else
         Exit For
      End If
   Next i
      
   Grid.rows = Grid.rows + 1
      
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
   
   Grid.ColWidth(C_RUT) = 1500
   Grid.ColWidth(C_NOMBRE) = 4500
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT Chofer"
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre Chofer"
    
End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim r As Integer

   Value = Trim(Value)
   
   Select Case Col
   
      Case C_RUT
      
         If Value <> "" Then
      
            If Value = "0-0" Then
               MsgBox1 "RUT inválido.", vbExclamation
               Action = vbRetry
               Exit Sub
            
            ElseIf Not ValidRut(Value) Then
               MsgBox1 "RUT inválido.", vbExclamation
               Grid.TextMatrix(Row, C_RUT) = Value
               Action = vbRetry
               Exit Sub
               
            Else
            
               Value = FmtRut(vFmtRut(Value))
            
               For r = Grid.FixedRows To Grid.rows - 1
                  If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
                     If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_RUT), vbTextCompare) = 0 Then
                        MsgBox1 "Este RUT de chofer ya existe.", vbExclamation
                        Call FGrSelRow(Grid, Row)
                        Action = vbRetry
                        Exit Sub
                     End If
                  End If
               Next r
            End If
            
         Else
            MsgBox1 "Debe ingresar un RUT válido.", vbExclamation
            Action = vbCancel
            Exit Sub
         End If
         
      Case C_NOMBRE
   
         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_NOMBRE), vbTextCompare) = 0 Then
                  MsgBox1 "Este chofer ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Exit Sub
               End If
            End If
         Next r
      
   
   End Select
    
    
   Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
   Call FGrEdRows(Grid, Row)
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
  
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row - 1, C_NOMBRE) = "" Or Grid.TextMatrix(Row - 1, C_RUT) = "" Then
      MsgBox1 "Debe completar la línea anterior.", vbExclamation
      Exit Sub
   End If
   
   If Col > C_RUT And Grid.TextMatrix(Row, C_RUT) = "" Then
      MsgBox1 "Debe ingresar un RUT válido antes de continuar.", vbExclamation
      Exit Sub
   End If
   
   Grid.TxBox.MaxLength = FldLen(Col)
   EdType = FEG_Edit
    
   If Grid.rows = Row + 1 Then
      Grid.rows = Grid.rows + 1
      Grid.FlxGrid.TopRow = Row
   End If
   
End Sub

Private Sub Grid_DblClick()

   If lOper = O_VIEW Then
      Call Bt_Sel_Click
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   If Grid.Col = C_RUT Then
      Call KeyRut(KeyAscii)
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
   
      
   If Grid.TextMatrix(Row, C_NOMBRE) = "" And Grid.TextMatrix(Row, C_RUT) = "" Then
      Exit Sub
      
   End If
   
   If MsgBox1("¿Está seguro que desea elimnar este Chofer?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
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
   gPrtReportes.ColObligatoria = C_NOMBRE
   gPrtReportes.NTotLines = 0
   

End Sub
Private Sub Bt_CopyExcel_Click()
   
   Call FGr2Clip(Grid, "Listado de Conductores")
   
End Sub
