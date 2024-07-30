VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMantTipoDocRef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Documentos de Referencia"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "FrmMantTipoDocRef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Print 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   8460
      Picture         =   "FrmMantTipoDocRef.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir listado"
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   8460
      Picture         =   "FrmMantTipoDocRef.frx":063B
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar Cláusula"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "&Copiar Excel"
      Height          =   855
      Left            =   8460
      Picture         =   "FrmMantTipoDocRef.frx":0C9D
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar datos a Excel"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8340
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton bt_Aceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8340
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   4395
      Left            =   1140
      TabIndex        =   0
      Top             =   540
      Width           =   6855
      _ExtentX        =   12091
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
      Height          =   570
      Left            =   240
      Picture         =   "FrmMantTipoDocRef.frx":1252
      Top             =   540
      Width           =   600
   End
End
Attribute VB_Name = "FrmMantTipoDocRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_NOMBRE = 1
Const C_ID = 2
Const C_ESFIJO = 3
Const C_UPDATE = 4
Private Const NCOLS = C_UPDATE

Private FldLen(NCOLS) As Byte
Dim lOrientacion As Integer

Private Sub Bt_Aceptar_Click()
   Dim Q1 As String
   Dim i As Integer
   
   For i = 1 To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_NOMBRE) <> "" Then
         
         If Trim(Grid.TextMatrix(i, C_CODIGO)) = "" Or Len(Trim(Grid.TextMatrix(i, C_CODIGO))) < 2 Then
            MsgBox1 "Código inválido '" & Grid.TextMatrix(i, C_CODIGO) & "'.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
                           
      End If
   Next i

   For i = 1 To Grid.rows - 1
      If Grid.TextMatrix(i, C_NOMBRE) <> "" Then
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
            Q1 = "INSERT INTO TipoDocRef (CodDocRefSII, Nombre, EsFijo)"
            Q1 = Q1 & " VALUES ('" & ParaSQL(Grid.TextMatrix(i, C_CODIGO)) & "'"
            Q1 = Q1 & ", '" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ",  0)"
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
            Q1 = "UPDATE TipoDocRef SET Nombre='" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ", CodDocRefSII='" & ParaSQL(Grid.TextMatrix(i, C_CODIGO)) & "'"
            Q1 = Q1 & ", EsFijo= " & Val(Grid.TextMatrix(i, C_ESFIJO))
            Q1 = Q1 & " WHERE IdTipoDocRef=" & Val(Grid.TextMatrix(i, C_ID))
            Call ExecSQL(DbMain, Q1)
         
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
         
            Q1 = "DELETE * FROM TipoDocRef"
            Q1 = Q1 & " WHERE IdTipoDocRef=" & Val(Grid.TextMatrix(i, C_ID))
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
   
   Q1 = "SELECT CodDocRefSII, Nombre, IdTipoDocRef, EsFijo FROM TipoDocRef ORDER BY EsFijo, Nombre"
   Call FillGrLista(Q1, Grid, C_ESFIJO, FldLen)
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_NOMBRE) = "" Then
         Exit For
      End If
      
      If Val(Grid.TextMatrix(i, C_ESFIJO)) <> 0 Then
         Call FGrSetRowStyle(Grid, i, "FC", vbBlue)
      End If
   Next i
   
   Grid.rows = Grid.rows + 1
   
   Call SetupPriv
   
End Sub
Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_CODIGO) = 1200
   Grid.ColWidth(C_NOMBRE) = 5200
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ESFIJO) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_NOMBRE) = "Descripción"
   Grid.TextMatrix(0, C_CODIGO) = "Tipo Doc."
    
End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim r As Integer

   Value = Trim(Value)
   
   Select Case Col
      Case C_NOMBRE
      
         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_NOMBRE), vbTextCompare) = 0 Then
                  MsgBox1 "Este Tipo de Documento de Referencia ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Exit Sub
               End If
            End If
         Next r
   
      Case C_CODIGO
         If Trim(Value) = "" Then
            MsgBox1 "Código de Tipo de Documento de Referencia inválido.", vbExclamation
         End If

         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_CODIGO), vbTextCompare) = 0 And Val(Grid.TextMatrix(r, C_CODIGO)) <> 8 Then    'el código 8 se puede repetir
                  MsgBox1 "Este Código de Documento de Referencia ya existe.", vbExclamation
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
   
   If Val(Grid.TextMatrix(Row, C_ESFIJO)) <> 0 Then
      MsgBox1 "Esta cláusula es fija", vbInformation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row - 1, C_NOMBRE) = "" Or Grid.TextMatrix(Row - 1, C_CODIGO) = "" Then
      MsgBox1 "Debe completar la línea anterior.", vbExclamation
      Exit Sub
   End If
   
   Grid.TxBox.MaxLength = FldLen(Col)
   If Col = C_NOMBRE Then
      Grid.TxBox.MaxLength = 65
   End If
   
   EdType = FEG_Edit
      
   
   If Grid.rows = Row + 1 Then
      Grid.rows = Grid.rows + 1
      Grid.FlxGrid.TopRow = Row
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)

   If Grid.Col = C_CODIGO Then
      Call KeyAlpha(KeyAscii)
      Call KeyUpper(KeyAscii)
      
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
   
   If Val(Grid.TextMatrix(Row, C_ESFIJO)) <> 0 Then
      MsgBox1 "No es posible eliminar este Tipo de Documento de Referencia.", vbExclamation
      Exit Sub
   End If
   
   If Id > 0 Then
      Q1 = "SELECT IdReferencia FROM Referencias WHERE IdTipoDocRef = " & Id & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         MsgBox1 "No es posible eliminar este Tipo de Documento de Referencia. Hay Documentos que incluyen referencias de este tipo.", vbExclamation
         Call CloseRs(Rs)
         Exit Sub
      End If
   
      Call CloseRs(Rs)
      
   ElseIf Grid.TextMatrix(Row, C_NOMBRE) = "" And Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Exit Sub
      
   End If
   
   If MsgBox1("¿Está seguro que desea elimnar este Tipo de Documento de Referencia?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
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
   
   Call FGr2Clip(Grid, "Listado de Tipos de Documentos de Referencia")
   
End Sub

