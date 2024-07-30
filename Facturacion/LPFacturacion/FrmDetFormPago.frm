VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDetFormPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor de Formas de Pago"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   12615
      Begin VB.Frame Frame2 
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6975
         Begin VB.CommandButton Bt_Limpiar 
            Caption         =   "Limpiar"
            Height          =   315
            Left            =   5040
            TabIndex        =   12
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Txt_Id 
            Height          =   285
            Left            =   4560
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Bt_Save 
            Caption         =   "Guardar"
            Height          =   840
            Left            =   5040
            Picture         =   "FrmDetFormPago.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar cambios componente seleccionada"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Bt_Cerrar 
            Caption         =   "Cerrar"
            Height          =   315
            Left            =   5040
            TabIndex        =   8
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox Tx_Descripcion 
            Height          =   315
            Left            =   240
            MaxLength       =   100
            TabIndex        =   6
            Top             =   720
            Width           =   4065
         End
         Begin VB.ComboBox Cb_Estado 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1560
            Width           =   1695
         End
         Begin VB.ComboBox Cb_FormaDePago 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1560
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3075
            Left            =   240
            TabIndex        =   10
            Top             =   2640
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   5424
            _Version        =   393216
            Cols            =   4
            FixedCols       =   3
            AllowUserResizing=   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Forma de Pago: "
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   5
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago:"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   3
            Top             =   1320
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "FrmDetFormPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_ID = 0
Const C_DESCRIPCION = 1
Const C_IDFORMAPAGO = 2
Const C_FORMAPAGO = 3
Const C_IDESTADO = 4
Const C_ESTADO = 5


Const NCOLS = C_ESTADO

Private Sub Bt_Cerrar_Click()
Unload Me
End Sub
Private Sub LoadAll()
  Dim Q1 As String
  Dim Rs As Recordset


  Q1 = " SELECT Id, Descripcion, FormaPago, Estado FROM DetFormaPago "
  Q1 = Q1 & " WHERE FormaPago = " & CbItemData(Me.Cb_FormaDePago)
  'Q1 = Q1 & " AND Estado = " & CbItemData(Me.Cb_Estado)
   
  Set Rs = OpenRs(DbMain, Q1)
   
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Grid.Redraw = False
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_ID) = vFld(Rs("Id"))
      Grid.TextMatrix(i, C_DESCRIPCION) = vFld(Rs("Descripcion"))
      Grid.TextMatrix(i, C_IDFORMAPAGO) = vFld(Rs("FormaPago"))
      Grid.TextMatrix(i, C_FORMAPAGO) = gFormaDePago(vFld(Rs("FormaPago")))
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      Grid.TextMatrix(i, C_ESTADO) = gEstado(vFld(Rs("Estado")))
      
      
      i = i + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
      
   Grid.Row = 0
   'Grid.Col = lOrdenSel
   Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Grid.Redraw = True
End Sub

Private Sub Bt_Limpiar_Click()
Me.Txt_Id.Text = 0
Me.Tx_Descripcion.Text = ""
End Sub

Private Sub Bt_Save_Click()
Dim Q1 As String

    If Not Valida() Then
        Exit Sub
    End If


      If Me.Txt_Id.Text <> "0" And Me.Txt_Id.Text <> "" Then
         Q1 = " UPDATE DETFORMAPAGO "
         Q1 = Q1 & " SET Descripcion = '" & Me.Tx_Descripcion.Text & "', "
         Q1 = Q1 & " FormaPago = " & CbItemData(Me.Cb_FormaDePago) & ", "
         Q1 = Q1 & " Estado = " & CbItemData(Me.Cb_Estado)
         Q1 = Q1 & " WHERE Id = " & Me.Txt_Id.Text
      Else
         Q1 = " INSERT INTO  DETFORMAPAGO (Descripcion, FormaPago, Estado) Values('" & Me.Tx_Descripcion.Text & "'," & CbItemData(Me.Cb_FormaDePago) & ", " & CbItemData(Me.Cb_Estado) & ") "
      End If
      Call ExecSQL(DbMain, Q1)
      
      Call LoadAll
End Sub


Private Sub Cb_FormaDePago_Click()
Call SetUpGrid
Call LoadAll
End Sub

Private Sub Form_Load()
Call CargaCbFormaPago
Call CargaCbEstado
Call SetUpGrid
Call LoadAll
Me.Txt_Id.Text = "0"
End Sub

Private Sub CargaCbFormaPago()
    For i = 1 To UBound(gFormaDePago)
       Call CbAddItem(Cb_FormaDePago, gFormaDePago(i), i)
    Next i
   Cb_FormaDePago.ListIndex = 0
End Sub
Private Sub CargaCbEstado()

    For i = 0 To UBound(gEstado)
       Call CbAddItem(Cb_Estado, gEstado(i), i)
    Next i
    Cb_Estado.ListIndex = ES_ACTIVO
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid, True)
   Grid.FixedCols = 0
 
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_DESCRIPCION) = 3000
   Grid.ColWidth(C_IDFORMAPAGO) = 0
   Grid.ColWidth(C_FORMAPAGO) = 1500
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_ESTADO) = 1300

   
   Grid.ColAlignment(C_DESCRIPCION) = flexAlignLeftCenter
   Grid.ColAlignment(C_FORMAPAGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_DESCRIPCION) = "Descripción"
   Grid.TextMatrix(0, C_FORMAPAGO) = "Forma de Pago"
   Grid.TextMatrix(0, C_ESTADO) = "Estado"

   
   Call FGrVRows(Grid, 1)
End Sub


Private Sub Grid_DblClick()
Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_ID) <> "" Then
        Me.Txt_Id.Text = Grid.TextMatrix(Row, C_ID)
        Me.Tx_Descripcion.Text = Grid.TextMatrix(Row, C_DESCRIPCION)
        Cb_FormaDePago.ListIndex = Grid.TextMatrix(Row, C_IDFORMAPAGO) - 1
        Cb_Estado.ListIndex = Grid.TextMatrix(Row, C_IDESTADO)
   End If
   
   
End Sub
Private Function Valida() As Boolean
   Valida = True
   
   If Tx_Descripcion.Text = "" Then
      MsgBox1 "Favor ingresar una Descripcion Forma de Pago", vbExclamation
      Valida = False
      Exit Function
   End If
   
'   If Not ValidRut(Tx_RUT) Then
'      MsgBox1 "Rut No es Valido", vbExclamation
'      Valida = False
'      Exit Function
'   End If
'
'   If Me.Tx_Codigo.Text = "" Then
'      MsgBox1 "Favor ingresar un Codigo", vbExclamation
'      Valida = False
'      Exit Function
'   End If
'
'   If Not ValidaCodigo() Then
'      Valida = False
'      Exit Function
'   End If


End Function
