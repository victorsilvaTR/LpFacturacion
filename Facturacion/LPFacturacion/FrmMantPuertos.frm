VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMantPuertos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puertos"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "FrmMantPuertos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Paises 
      Caption         =   "Países"
      Height          =   855
      Left            =   10080
      Picture         =   "FrmMantPuertos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Mantención de Países"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Print 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   10080
      Picture         =   "FrmMantPuertos.frx":03F3
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir lista de Puertos"
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   10080
      Picture         =   "FrmMantPuertos.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar Puerto"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "&Copiar Excel"
      Height          =   855
      Left            =   10080
      Picture         =   "FrmMantPuertos.frx":1084
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar datos a Excel"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   9960
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton bt_Aceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   9960
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   4395
      Left            =   1140
      TabIndex        =   0
      Top             =   540
      Width           =   8355
      _ExtentX        =   14737
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
   Begin VB.Label la_LinkAduana2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.aduana.cl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1080
      MouseIcon       =   "FrmMantPuertos.frx":1639
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6000
      Width           =   6615
   End
   Begin VB.Label La_LinkAduana1 
      Caption         =   "Link Aduanas"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1080
      TabIndex        =   7
      Top             =   5760
      Width           =   6555
   End
   Begin VB.Label La_Info 
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   5160
      Width           =   6555
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   240
      Picture         =   "FrmMantPuertos.frx":178B
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "FrmMantPuertos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NOMBRE = 0
Const C_CODIGO = 1
Const C_ID = 2
Const C_IDPAIS = 3
Const C_PAIS = 4
Const C_UPDATE = 5
Private Const NCOLS = C_UPDATE

Private FldLen(NCOLS) As Byte
Dim lOrientacion As Integer

Private Sub Bt_Aceptar_Click()
   Dim Q1 As String
   Dim i As Integer
   
   For i = 1 To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_NOMBRE) <> "" Then
         
         If Trim(Grid.TextMatrix(i, C_CODIGO)) = "" Or (Grid.TextMatrix(i, C_CODIGO) <> "" And Val(Grid.TextMatrix(i, C_CODIGO)) <= 0) Then  ' se verifica que no pongan 000
            MsgBox1 "Código inválido '" & Grid.TextMatrix(i, C_CODIGO) & "'.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
         
         If Val(Grid.TextMatrix(i, C_IDPAIS)) = 0 Then
            MsgBox1 "Falta seleccionar el País.", vbExclamation
            Call FGrSelRow(Grid, i)
            Exit Sub
         End If
                  
      End If
   Next i

   For i = 1 To Grid.rows - 1
      If Grid.TextMatrix(i, C_NOMBRE) <> "" Then
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
            Q1 = "INSERT INTO Puertos (Codigo, Nombre, IdPais)"
            Q1 = Q1 & " VALUES ('" & ParaSQL(Grid.TextMatrix(i, C_CODIGO)) & "'"
            Q1 = Q1 & ", '" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ",  " & Grid.TextMatrix(i, C_IDPAIS) & ")"
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
            Q1 = "UPDATE Puertos SET Nombre='" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
            Q1 = Q1 & ", Codigo='" & ParaSQL(Grid.TextMatrix(i, C_CODIGO)) & "'"
            Q1 = Q1 & ", IdPais= " & Grid.TextMatrix(i, C_IDPAIS)
            Q1 = Q1 & " WHERE Id=" & Val(Grid.TextMatrix(i, C_ID))
            Call ExecSQL(DbMain, Q1)
         
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
         
            Q1 = "DELETE * FROM Puertos"
            Q1 = Q1 & " WHERE Id=" & Val(Grid.TextMatrix(i, C_ID))
            Call ExecSQL(DbMain, Q1)
         End If
      End If
   Next i
   
   Unload Me
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_Paises_Click()
   Dim Frm As FrmMantPaises
   Dim Q1 As String
   
   Set Frm = New FrmMantPaises
   Frm.Show vbModal
   Set Frm = Nothing
   
   Grid.CbList(C_PAIS).Clear
   Call CbAddItem(Grid.CbList(C_PAIS), " ", 0)
   Q1 = "SELECT Nombre, Id FROM Paises ORDER BY Nombre"
   Call FillCombo(Grid.CbList(C_PAIS), DbMain, Q1, -1)

   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   Call SetUpGrid
   
   Q1 = "SELECT Puertos.Nombre, Puertos.Codigo, Puertos.Id, Puertos.IdPais, Paises.Nombre FROM Puertos INNER JOIN Paises ON Puertos.IdPais = Paises.Id ORDER BY Puertos.Nombre"
   Call FillGrLista(Q1, Grid, C_PAIS, FldLen)
   
   Call CbAddItem(Grid.CbList(C_PAIS), " ", 0)
   Q1 = "SELECT Nombre, Id FROM Paises ORDER BY Nombre"
   Call FillCombo(Grid.CbList(C_PAIS), DbMain, Q1, -1)

      
   La_Info = "El Código de Aduana para los Puertos tiene tres dígitos y debe ser distinto de cero."
     
   La_LinkAduana1 = "Para ver los Códigos de Aduana de los Puertos ingrese a:"
   
'   la_LinkAduana2.Tag = "https://www.aduana.cl/aduana/site/artic/20080218/pags/20080218165942.html#vtxt_cuerpo_T11"
   la_LinkAduana2.Tag = "http://comext.aduana.cl:7001/codigos"
   
   Call SetupPriv
   
End Sub
Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_NOMBRE) = 3400
   Grid.ColWidth(C_CODIGO) = 1200
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_IDPAIS) = 0
   Grid.ColWidth(C_PAIS) = 3400
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid.ColAlignment(C_PAIS) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_NOMBRE) = "Puerto"
   Grid.TextMatrix(0, C_CODIGO) = "Código Aduana"
   Grid.TextMatrix(0, C_PAIS) = "País"
    
End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim r As Integer

   Value = Trim(Value)
   
   Select Case Col
      Case C_NOMBRE
      
         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_NOMBRE), vbTextCompare) = 0 Then
                  MsgBox1 "Este Puerto ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Exit Sub
               End If
            End If
         Next r
   
      Case C_CODIGO
         Value = Right("000" & Value, 3)
      
         If Val(Value) = 0 Then
            MsgBox1 "Código de Puerto inválido.", vbExclamation
         End If

         For r = Grid.FixedRows To Grid.rows - 1
            If (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_ID) <> "") And Grid.RowHeight(r) > 0 Then
               If r <> Row And StrComp(Value, Grid.TextMatrix(r, C_CODIGO), vbTextCompare) = 0 Then
                  MsgBox1 "Este Código de Puerto ya existe.", vbExclamation
                  Call FGrSelRow(Grid, Row)
                  Action = vbRetry
                  Exit Sub
               End If
            End If
         Next r

      Case C_PAIS
         Grid.TextMatrix(Row, C_IDPAIS) = CbItemData(Grid.CbList(C_PAIS))

   
   End Select
    
   Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
   Call FGrEdRows(Grid, Row)
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row - 1, C_NOMBRE) = "" Or Grid.TextMatrix(Row - 1, C_CODIGO) = "" Or Grid.TextMatrix(Row - 1, C_PAIS) = "" Then
      MsgBox1 "Debe completar la línea anterior.", vbExclamation
      Exit Sub
   End If
   
   Select Case Col
      Case C_NOMBRE, C_CODIGO
         Grid.TxBox.MaxLength = FldLen(Col)
         EdType = FEG_Edit
      Case C_PAIS
         EdType = FEG_List
   End Select
   
   If Grid.rows = Row + 1 Then
      Grid.rows = Grid.rows + 1
      Grid.FlxGrid.TopRow = Row
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   If Grid.Col = C_CODIGO Then
      Call KeyNumPos(KeyAscii)
   End If
      
End Sub

Private Sub SetupPriv()

   Call EnableForm0(Me, ChkPriv(PRVF_MANT_DATOS))
         
End Sub

Private Sub La_LinkAduana2_Click()
   Dim Rc As Long
   Dim Buf As String
   
   If la_LinkAduana2.Tag <> "" Then
      Buf = la_LinkAduana2.Tag
   Else
      Buf = la_LinkAduana2
   End If

   If LCase(Left(Buf, 4)) <> "http" Then
      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", SW_SHOWNORMAL)
   Else
      Rc = Shell(gHtmExt.OpenCmd & " " & Buf, vbNormalFocus)
   End If

End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdPuerto As Long
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdPuerto = Val(Grid.TextMatrix(Row, C_ID))
   
   If IdPuerto > 0 Then
      Q1 = "SELECT IdDTEFactExp FROM DTEFactExp WHERE (IdPuertoEmbarque = " & IdPuerto & " OR IdPuertoDesembarque = " & IdPuerto & ") AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         MsgBox1 "No es posible eliminar este puerto. Hay Documentos de Esportación cuyos antecedentes incluyen este puerto.", vbExclamation
         Call CloseRs(Rs)
         Exit Sub
      End If
   
      Call CloseRs(Rs)
      
   ElseIf Grid.TextMatrix(Row, C_NOMBRE) = "" And Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Exit Sub
      
   End If
   
   If MsgBox1("¿Está seguro que desea elimnar este Puerto?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
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
   
   Call FGr2Clip(Grid, "Listado de Puertos")
   
End Sub

